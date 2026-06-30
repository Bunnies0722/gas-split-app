/**
 * 資源エネルギー庁「石油製品価格調査」から、その週のレギュラー/ハイオク/軽油の
 * 全国平均価格を自動取得して price.json に保存するスクリプト。
 *
 * ねらい：
 *   経済産業省(資源エネルギー庁)はこのデータをAPIとして公開していません。
 *   毎週水曜14時に、ページ上のExcelファイル（ファイル名に日付が入る）として
 *   更新されるだけなので、「結果ページを読む→最新のExcelリンクを見つける→
 *   ダウンロードして数値を読む」という処理を自前で行う必要があります。
 *
 * 重要な制約（必ず読んでください）：
 *   このサイトは一般的な政府サイトと同様、ブラウザからの直接fetch（JavaScript）に
 *   対応するCORSヘッダーを返していないため、ブラウザ上のJSから直接呼び出すと
 *   ほぼ確実にCORSエラーになります。
 *   → そのため、このスクリプトは「サーバー側（Node.js／Vercelのサーバーレス関数など）」
 *     で実行してください。CORSはブラウザだけが適用するルールなので、
 *     サーバー側からのfetchには影響しません。
 *
 *   また、このスクリプトのExcel解析部分（findFuelPrices関数）は、想定される
 *   一般的な統計表のレイアウト（油種名が並ぶヘッダー行 → その下に「全国」の行）
 *   に基づいて書いていますが、実際のファイルの正確なレイアウトを私の方で
 *   直接確認することができていません（私の実行環境からはenecho.meti.go.jpへの
 *   通信がブロックされていたため）。そのため、初回実行時は必ず
 *   `--debug` オプション付きで実行し、出力されるシート内容を見ながら
 *   キーワードやロジックを調整してください。
 *
 * 使い方：
 *   npm install xlsx
 *   node fetch-gas-price.js            … price.json を生成
 *   node fetch-gas-price.js --debug    … 見つけたシートの中身をそのまま表示（調整用）
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx'); // npm install xlsx で導入（SheetJS）

const RESULTS_PAGE_URL = 'https://www.enecho.meti.go.jp/statistics/petroleum_and_lpgas/pl007/results.html';
const BASE_URL = 'https://www.enecho.meti.go.jp';
const OUTPUT_PATH = path.join(__dirname, '..', 'prototype', 'price.json');

const FUEL_KEYWORDS = {
  regular: 'レギュラー',
  highoctane: 'ハイオク',
  diesel: '軽油',
};

// 経済産業局別シートの地域列（全国列の左側に並ぶ9地域）。
// 出発地・到着地の都道府県からこの9地域のどれかに割り当てて、
// 2地点の地域価格の平均をおすすめ単価として使う。
const REGION_LABELS = ['北海道', '東北', '関東', '中部', '近畿', '中国', '四国', '九州', '沖縄'];

async function main() {
  const debug = process.argv.includes('--debug');

  console.log('1. 結果ページを取得中...', RESULTS_PAGE_URL);
  const html = await fetchText(RESULTS_PAGE_URL);

  const xlsxUrl = findLatestXlsxUrl(html);
  if (!xlsxUrl) {
    throw new Error('結果ページから最新のExcelリンクが見つかりませんでした。サイト構成が変わった可能性があります。');
  }
  console.log('2. 最新ファイルを発見:', xlsxUrl);

  console.log('3. Excelファイルをダウンロード中...');
  const buffer = await fetchBuffer(xlsxUrl);

  console.log('4. Excelを解析中...');
  const workbook = XLSX.read(buffer, { type: 'buffer' });

  if (debug) {
    dumpWorkbook(workbook);
    return;
  }

  const extracted = findFuelPrices(workbook);
  if (!extracted) {
    throw new Error(
      '価格データを抽出できませんでした。`node fetch-gas-price.js --debug` でシート内容を確認し、' +
      'FUEL_KEYWORDS や findFuelPrices のロジックを実際のレイアウトに合わせて調整してください。'
    );
  }

  const result = {
    fetchedAt: new Date().toISOString(),
    weekOf: excelSerialToISODate(extracted.latestDateSerial),
    sourceUrl: xlsxUrl,
    unit: '円/L',
    prices: extracted.prices, // { regular, highoctane, diesel } … 全国平均
    regionPrices: extracted.regionPrices, // { regular: {北海道, 東北, ...}, highoctane: {...}, diesel: {...} }
  };

  fs.writeFileSync(OUTPUT_PATH, JSON.stringify(result, null, 2), 'utf-8');
  console.log('5. 保存しました:', OUTPUT_PATH);
  console.log(result);
}

// 結果ページのHTMLから、最新の「結果詳細版」Excelのリンクを抜き出す。
// 例: .../pl007/xlsx/260610.xlsx （6桁の日付＋.xlsxのみ。s5/t/t5/k/k5/a/a5/oなどの
// 接尾辞付きファイルは別カテゴリの過去データ・産業用価格などなので対象外にする）
function findLatestXlsxUrl(html) {
  const match = html.match(/href="([^"]*pl007\/xlsx\/(\d{6})\.xlsx)"/);
  if (!match) return null;
  const url = match[1];
  return url.startsWith('http') ? url : BASE_URL + url;
}

// ワークブックの中から燃料別の週次データを探して最新の全国平均価格を取り出す。
// 実ファイルのレイアウト（経済産業局別シート）:
//   - 列方向に地域が並び、「全国（円/ﾘｯﾄﾙ）」列から価格を読む
//   - 行方向に油種セクション（ハイオク→レギュラー→軽油）が並び、各6週分の行が続く
//   - 各セクション先頭で調査日シリアルがリセット（46160→...→46195 を繰り返す）
//   - 「軽油」はExcelの結合セルで「軽」と「油」が別行に分かれているため、
//     セクション内の全セルを結合してキーワード検索する
function findFuelPrices(workbook) {
  const normalize = (s) => String(s).replace(/[\s　]/g, '');
  const isDateSerial = (v) => typeof v === 'number' && v > 40000 && v < 50000;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    // '全国' を含む列インデックスを探す
    let nationalCol = -1;
    for (let r = 0; r < rows.length; r++) {
      const idx = rows[r].findIndex((cell) => String(cell).includes('全国'));
      if (idx !== -1) { nationalCol = idx; break; }
    }
    if (nationalCol === -1) continue;

    // 9地域（北海道〜沖縄）の列インデックスを探す。
    // 「九州及び沖縄」という結合列もあり「九州」を含んでしまうため、
    // includes ではなく完全一致（normalize後）で判定する。
    const regionCols = {};
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        const cellText = normalize(rows[r][c]);
        for (const label of REGION_LABELS) {
          if (regionCols[label] === undefined && cellText === label) {
            regionCols[label] = c;
          }
        }
      }
    }

    // 調査日シリアルが前週より小さくなる（リセット）たびに新セクション開始と判定
    const groups = [];
    let currentGroup = null;
    let prevDate = Infinity;
    for (let r = 0; r < rows.length; r++) {
      const val = rows[r][2];
      if (!isDateSerial(val)) continue;
      if (val < prevDate) {
        currentGroup = { rows: [r] };
        groups.push(currentGroup);
      } else {
        currentGroup.rows.push(r);
      }
      prevDate = val;
    }

    // 各セクション内の全col0+col1テキストを結合して燃料種別を特定し、最新価格を取得
    const prices = {};
    const regionPrices = {};
    let latestDateSerial = 0;
    for (const group of groups) {
      const combined = group.rows
        .map((r) => normalize(rows[r][0]) + normalize(rows[r][1]))
        .join('');

      for (const [key, label] of Object.entries(FUEL_KEYWORDS)) {
        if (prices[key] !== undefined) continue;
        if (!combined.includes(normalize(label))) continue;

        let latestDate = -1, latestPrice = NaN, latestRow = -1;
        for (const r of group.rows) {
          const d = rows[r][2], v = Number(rows[r][nationalCol]);
          if (!Number.isNaN(v) && v > 0 && d > latestDate) {
            latestDate = d;
            latestPrice = v;
            latestRow = r;
          }
        }
        if (!Number.isNaN(latestPrice)) {
          prices[key] = latestPrice;
          latestDateSerial = Math.max(latestDateSerial, latestDate);

          const regionsForFuel = {};
          for (const label of REGION_LABELS) {
            const col = regionCols[label];
            if (col === undefined) continue;
            const v = Number(rows[latestRow][col]);
            if (!Number.isNaN(v) && v > 0) regionsForFuel[label] = v;
          }
          if (Object.keys(regionsForFuel).length > 0) regionPrices[key] = regionsForFuel;
        }
        break;
      }

      if (Object.keys(prices).length === Object.keys(FUEL_KEYWORDS).length) break;
    }

    if (Object.keys(prices).length === Object.keys(FUEL_KEYWORDS).length) {
      return { prices, regionPrices, latestDateSerial };
    }
  }
  return null;
}

// Excelの日付シリアル値（1900年1月1日起点）をISO日付文字列(YYYY-MM-DD)に変換する
function excelSerialToISODate(serial) {
  if (!serial) return null;
  const ms = (serial - 25569) * 86400 * 1000;
  return new Date(ms).toISOString().slice(0, 10);
}

// --debug 用：各シートの先頭30行をそのまま表示する（実際のレイアウト確認用）
function dumpWorkbook(workbook) {
  for (const sheetName of workbook.SheetNames) {
    console.log(`\n===== シート: ${sheetName} =====`);
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' });
    rows.slice(0, 30).forEach((row, i) => console.log(i, row));
  }
}

async function fetchText(url) {
  const res = await fetch(url); // Node.js 18以降は標準でfetchが使えます
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${url}`);
  return await res.text();
}

async function fetchBuffer(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${url}`);
  const arrayBuffer = await res.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

main().catch((err) => {
  console.error('エラー:', err.message);
  process.exit(1);
});
