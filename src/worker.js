/**
 * Cloudflare Worker本体。
 *
 * 1. 通常のHTTPリクエスト（fetch）は静的アセット（prototypeフォルダ）にそのまま委譲する。
 * 2. 週次のCron Trigger（scheduled）で、資源エネルギー庁のガソリン価格Excelを取得・解析し、
 *    GitHub API経由で prototype/price.json を直接コミットする。
 *    → このコミットがCloudflareのGit連携の再デプロイを自動的にトリガーするので、
 *      手動でnode実行・git push・再デプロイをする必要がなくなる。
 *
 * 必要なSecret（Cloudflareダッシュボード or `wrangler secret put` で設定）：
 *   - GITHUB_TOKEN: prototype/price.json を書き込めるGitHub Personal Access Token
 *                   （Fine-grained PAT、対象リポジトリに Contents: Read and write 権限）
 *
 * 動作確認用に、GET /__cron-test?key=<GITHUB_TOKENと同じ値> を呼ぶと、
 * Cron同じ処理を即時実行できる（本番運用前のテスト用）。
 */

import * as XLSX from 'xlsx';

const RESULTS_PAGE_URL = 'https://www.enecho.meti.go.jp/statistics/petroleum_and_lpgas/pl007/results.html';
const BASE_URL = 'https://www.enecho.meti.go.jp';

const GITHUB_OWNER = 'Bunnies0722';
const GITHUB_REPO = 'gas-split-app';
const GITHUB_FILE_PATH = 'prototype/price.json';
const GITHUB_BRANCH = 'main';

const FUEL_KEYWORDS = {
  regular: 'レギュラー',
  highoctane: 'ハイオク',
  diesel: '軽油',
};

const REGION_LABELS = ['北海道', '東北', '関東', '中部', '近畿', '中国', '四国', '九州', '沖縄'];

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    // 動作確認用エンドポイント（本番運用前のテスト・トラブル時の手動再実行に使う）
    if (url.pathname === '/__cron-test') {
      const key = url.searchParams.get('key');
      if (!env.GITHUB_TOKEN || key !== env.GITHUB_TOKEN) {
        return new Response('Forbidden', { status: 403 });
      }
      try {
        const result = await updateGasPrice(env);
        return new Response(JSON.stringify(result, null, 2), {
          headers: { 'content-type': 'application/json; charset=utf-8' },
        });
      } catch (err) {
        return new Response(`Error: ${err.message}\n${err.stack || ''}`, { status: 500 });
      }
    }

    // それ以外は静的アセット（prototypeフォルダ）にそのまま委譲
    return env.ASSETS.fetch(request);
  },

  async scheduled(event, env, ctx) {
    ctx.waitUntil(
      updateGasPrice(env).catch((err) => {
        console.error('週次ガソリン価格の自動取得に失敗しました:', err.message, err.stack);
      })
    );
  },
};

async function updateGasPrice(env) {
  if (!env.GITHUB_TOKEN) {
    throw new Error('GITHUB_TOKEN シークレットが設定されていません。');
  }

  const html = await fetchText(RESULTS_PAGE_URL);
  const xlsxUrl = findLatestXlsxUrl(html);
  if (!xlsxUrl) {
    throw new Error('結果ページから最新のExcelリンクが見つかりませんでした。サイト構成が変わった可能性があります。');
  }

  const buffer = await fetchBuffer(xlsxUrl);
  const workbook = XLSX.read(buffer, { type: 'buffer' });

  const extracted = findFuelPrices(workbook);
  if (!extracted) {
    throw new Error('価格データを抽出できませんでした。Excelのレイアウトが変わった可能性があります。');
  }

  const result = {
    fetchedAt: new Date().toISOString(),
    weekOf: excelSerialToISODate(extracted.latestDateSerial),
    sourceUrl: xlsxUrl,
    unit: '円/L',
    prices: extracted.prices,
    regionPrices: extracted.regionPrices,
  };

  await commitToGitHub(env, result);
  return result;
}

function findLatestXlsxUrl(html) {
  const match = html.match(/href="([^"]*pl007\/xlsx\/(\d{6})\.xlsx)"/);
  if (!match) return null;
  const url = match[1];
  return url.startsWith('http') ? url : BASE_URL + url;
}

function findFuelPrices(workbook) {
  const normalize = (s) => String(s).replace(/[\s　]/g, '');
  const isDateSerial = (v) => typeof v === 'number' && v > 40000 && v < 50000;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    let nationalCol = -1;
    for (let r = 0; r < rows.length; r++) {
      const idx = rows[r].findIndex((cell) => String(cell).includes('全国'));
      if (idx !== -1) { nationalCol = idx; break; }
    }
    if (nationalCol === -1) continue;

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

function excelSerialToISODate(serial) {
  if (!serial) return null;
  const ms = (serial - 25569) * 86400 * 1000;
  return new Date(ms).toISOString().slice(0, 10);
}

async function fetchText(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${url}`);
  return await res.text();
}

async function fetchBuffer(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${url}`);
  const arrayBuffer = await res.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

// GitHub Contents APIで prototype/price.json を直接更新する。
// これにより、Cloudflareのpush連携が自動で再デプロイしてくれる。
async function commitToGitHub(env, jsonObject) {
  const apiBase = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${GITHUB_FILE_PATH}`;
  const headers = {
    Authorization: `Bearer ${env.GITHUB_TOKEN}`,
    'User-Agent': 'gas-split-app-price-bot',
    Accept: 'application/vnd.github+json',
  };

  // 既存ファイルのSHAを取得（更新には必須）
  const getRes = await fetch(`${apiBase}?ref=${GITHUB_BRANCH}`, { headers });
  if (!getRes.ok) {
    throw new Error(`GitHubから既存ファイル情報の取得に失敗: HTTP ${getRes.status}`);
  }
  const getData = await getRes.json();
  const sha = getData.sha;

  const content = JSON.stringify(jsonObject, null, 2) + '\n';
  const base64Content = Buffer.from(content, 'utf-8').toString('base64');

  const putRes = await fetch(apiBase, {
    method: 'PUT',
    headers: { ...headers, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      message: `chore: update gas price data (${jsonObject.weekOf || jsonObject.fetchedAt})`,
      content: base64Content,
      sha,
      branch: GITHUB_BRANCH,
    }),
  });

  if (!putRes.ok) {
    const errText = await putRes.text();
    throw new Error(`GitHubへのコミットに失敗: HTTP ${putRes.status} ${errText}`);
  }
}
