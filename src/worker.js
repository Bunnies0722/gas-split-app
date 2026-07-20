/**
 * Cloudflare Worker本体。
 *
 * 1. 通常のHTTPリクエスト（fetch）は静的アセット（prototypeフォルダ）にそのまま委譲する。
 * 2. 週次のCron Trigger（scheduled）で、資源エネルギー庁のガソリン価格Excelを取得・解析し、
 *    GitHub API経由で prototype/price.json を直接コミットする。
 *    → このコミットがCloudflareのGit連携の再デプロイを自動的にトリガーするので、
 *      手動でnode実行・git push・再デプロイをする必要がなくなる。
 *
 * 3. GET /api/route で、出発地点・到着地点の緯度経度から道路に沿ったルート・距離を返す。
 *    avoidHighways=false → OSRM公開インスタンス（APIキー不要）
 *    avoidHighways=true  → Valhalla公開インスタンス（APIキー不要、use_highways:0で高速回避）
 *    いずれも外部APIキー不要。
 *
 * 4. /api/groups 以下のエンドポイントで、グループ（メンバー名）と精算履歴をSupabase（Postgres）に保存する。
 *    ログイン不要。グループ作成時に発行されるランダムな共有コードが、そのままアクセス権のような役割を果たす
 *    （コードを知っている人だけが読み書きできる、URL共有型の設計）。SupabaseのService Role Keyは
 *    Worker側だけが保持し、フロントには一切渡さない。
 *      - POST   /api/groups                  グループ作成（name, members[]） → { id, code, name, members }
 *      - GET    /api/groups/:code             グループ情報＋メンバー＋履歴を取得
 *      - PUT    /api/groups/:code/members     メンバー一覧を入れ替え
 *      - POST   /api/groups/:code/trips       精算履歴を1件追加
 *
 * 必要なSecret（Cloudflareダッシュボード or `wrangler secret put` で設定）：
 *   - GITHUB_TOKEN: prototype/price.json を書き込めるGitHub Personal Access Token
 *                   （Fine-grained PAT、対象リポジトリに Contents: Read and write 権限）
 *   - ORS_API_KEY: OpenRouteService（openrouteservice.org）の無料APIキー
 *                  （Directions APIを使うために必要。無料枠で1日2500回まで利用可）
 *   - SUPABASE_URL: SupabaseプロジェクトのURL（例：https://xxxxx.supabase.co）
 *   - SUPABASE_SERVICE_ROLE_KEY: SupabaseのService Role Key（Project Settings → API）
 *                  ※RLSを有効にしテーブルへのポリシーを作らないことで、このキーを持たない
 *                    フロント側からは一切アクセスできないようにしている。
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

    // 道路距離・道路沿いの経路取得（OpenRouteService Directions APIへのプロキシ）
    // フロント側からAPIキーを隠すため、Worker経由で呼び出す。
    if (url.pathname === '/api/route') {
      try {
        return await getRoute(url, env);
      } catch (err) {
        return new Response(JSON.stringify({ error: err.message }), {
          status: 502,
          headers: { 'content-type': 'application/json; charset=utf-8' },
        });
      }
    }

    // ルーティングAPIの診断エンドポイント
    // /api/route-debug にアクセスすると Valhalla / OSRM の動作を直接確認できる
    if (url.pathname === '/api/route-debug') {
      return await routeDebug();
    }

    // グループ作成（メンバー名を保存する“箱”を作る）
    if (url.pathname === '/api/groups' && request.method === 'POST') {
      try {
        return await createGroup(request, env);
      } catch (err) {
        return jsonResponse({ error: err.message }, 502);
      }
    }

    // グループ情報＋メンバー＋履歴の取得
    const groupMatch = url.pathname.match(/^\/api\/groups\/([A-Za-z0-9]+)$/);
    if (groupMatch && request.method === 'GET') {
      try {
        return await getGroupInfo(groupMatch[1], env);
      } catch (err) {
        return jsonResponse({ error: err.message }, 502);
      }
    }

    // メンバー一覧の入れ替え
    const membersMatch = url.pathname.match(/^\/api\/groups\/([A-Za-z0-9]+)\/members$/);
    if (membersMatch && request.method === 'PUT') {
      try {
        return await updateGroupMembers(membersMatch[1], request, env);
      } catch (err) {
        return jsonResponse({ error: err.message }, 502);
      }
    }

    // 精算履歴の追加
    const tripsMatch = url.pathname.match(/^\/api\/groups\/([A-Za-z0-9]+)\/trips$/);
    if (tripsMatch && request.method === 'POST') {
      try {
        return await createTrip(tripsMatch[1], request, env);
      } catch (err) {
        return jsonResponse({ error: err.message }, 502);
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

// /api/route-debug: Valhalla・OSRM の動作確認（東京駅→大阪駅で両エンジンをテスト）
async function routeDebug() {
  const startLat = 35.6812, startLng = 139.7671; // 東京駅
  const endLat = 34.7024, endLng = 135.4959;     // 大阪駅

  const results = {};

  // OSRM テスト
  try {
    const osrmUrl = `https://router.project-osrm.org/route/v1/driving/${startLng},${startLat};${endLng},${endLat}?overview=false`;
    const res = await fetch(osrmUrl, { headers: { 'User-Agent': 'gas-split-app/1.0' } });
    const data = await res.json();
    results.osrm = {
      status: res.status,
      code: data.code,
      distanceKm: data.routes && data.routes[0] ? (data.routes[0].distance / 1000).toFixed(1) : null,
    };
  } catch (err) {
    results.osrm = { error: err.message };
  }

  // Valhalla テスト（高速回避）
  try {
    const body = {
      locations: [{ lat: startLat, lon: startLng }, { lat: endLat, lon: endLng }],
      costing: 'auto',
      costing_options: { auto: { use_highways: 0.0, use_tolls: 0.0 } },
    };
    const res = await fetch('https://valhalla.openstreetmap.de/route', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'User-Agent': 'gas-split-app/1.0' },
      body: JSON.stringify(body),
    });
    const text = await res.text();
    let data;
    try { data = JSON.parse(text); } catch { data = { raw: text.slice(0, 300) }; }
    results.valhalla_no_highway = {
      status: res.status,
      distanceKm: data.trip && data.trip.summary ? data.trip.summary.length.toFixed(1) : null,
      error: data.error || null,
      statusMessage: data.trip && data.trip.status_message || null,
    };
  } catch (err) {
    results.valhalla_no_highway = { error: err.message };
  }

  // Valhalla テスト（高速あり、比較用）
  try {
    const body = {
      locations: [{ lat: startLat, lon: startLng }, { lat: endLat, lon: endLng }],
      costing: 'auto',
    };
    const res = await fetch('https://valhalla.openstreetmap.de/route', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'User-Agent': 'gas-split-app/1.0' },
      body: JSON.stringify(body),
    });
    const data = await res.json();
    results.valhalla_with_highway = {
      status: res.status,
      distanceKm: data.trip && data.trip.summary ? data.trip.summary.length.toFixed(1) : null,
      error: data.error || null,
    };
  } catch (err) {
    results.valhalla_with_highway = { error: err.message };
  }

  return new Response(JSON.stringify(results, null, 2), {
    headers: { 'content-type': 'application/json; charset=utf-8' },
  });
}

// /api/route?startLat=..&startLng=..&endLat=..&endLng=..&avoidHighways=true を受け取り、
// 実際の道路に沿ったルートと距離を返す。
//
// ルーティングの優先順位:
//   avoidHighways=false（デフォルト）: OSRM → 失敗時はORS → 失敗時はエラー
//   avoidHighways=true（一般道のみ）:  ORS（avoid_features:highways） → 失敗時はOSRM → 失敗時はエラー
//
// OSRMはAPIキー不要で安定しているので優先的に使う。
// ORSは高速道路回避オプションに必要なためAPIキーを持つ場合のみ利用。
//
// レスポンス: { distanceKm: number, geometry: [[lat, lng], ...], source: 'osrm'|'ors' }
async function getRoute(url, env) {
  const startLat = parseFloat(url.searchParams.get('startLat'));
  const startLng = parseFloat(url.searchParams.get('startLng'));
  const endLat = parseFloat(url.searchParams.get('endLat'));
  const endLng = parseFloat(url.searchParams.get('endLng'));
  const avoidHighways = url.searchParams.get('avoidHighways') === 'true';

  if ([startLat, startLng, endLat, endLng].some((v) => Number.isNaN(v))) {
    throw new Error('startLat/startLng/endLat/endLng が不正です。');
  }

  const errors = [];

  if (!avoidHighways) {
    // 高速道路あり: OSRM（APIキー不要、安定）
    try {
      return await fetchRouteOSRM(startLat, startLng, endLat, endLng, { avoidanceRequested: false, avoidanceApplied: false });
    } catch (err) {
      errors.push(err.message);
    }
  } else {
    // 一般道のみ: Valhalla公式（APIキー不要、use_highways:0 で確実に高速回避）
    try {
      return await fetchRouteValhalla(startLat, startLng, endLat, endLng);
    } catch (err) {
      errors.push(err.message);
      // Valhallaが失敗した場合はOSRMにフォールバック（高速回避なし・要通知）
      try {
        return await fetchRouteOSRM(startLat, startLng, endLat, endLng, { avoidanceRequested: true, avoidanceApplied: false });
      } catch (err2) {
        errors.push(err2.message);
      }
    }
  }

  throw new Error(`ルート取得失敗: ${errors.join(' / ')}`);
}

// Valhalla公式公開インスタンス（APIキー不要）でルートを取得する。
// use_highways: 0.0 にすることで高速道路・有料道路を回避する。
// エンドポイント: https://valhalla.openstreetmap.de (Geofabrik運営)
async function fetchRouteValhalla(startLat, startLng, endLat, endLng) {
  const body = {
    locations: [
      { lat: startLat, lon: startLng },
      { lat: endLat, lon: endLng },
    ],
    costing: 'auto',
    costing_options: {
      auto: {
        use_highways: 0.0,  // 0=完全回避, 1=積極利用
        use_tolls: 0.0,     // 有料道路も回避
      },
    },
  };

  const res = await fetch('https://valhalla.openstreetmap.de/route', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'User-Agent': 'gas-split-app/1.0 (hobby project)',
    },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Valhalla HTTP ${res.status}: ${text.slice(0, 200)}`);
  }

  const data = await res.json();
  if (!data.trip || !data.trip.legs || !data.trip.legs[0]) {
    throw new Error('Valhalla: ルートなし');
  }

  const leg = data.trip.legs[0];
  const distanceKm = leg.summary.length; // Valhallaはkm単位
  // shapeはpolyline6エンコード（精度1e-6）。Leafletで使う[lat, lng]の順に返す。
  const geometry = decodePolyline6(leg.shape);

  return new Response(
    JSON.stringify({ distanceKm, geometry, source: 'valhalla', avoidanceRequested: true, avoidanceApplied: true }),
    { headers: { 'content-type': 'application/json; charset=utf-8' } }
  );
}

// Valhallaのpolyline6エンコードをデコードして [[lat, lng], ...] の配列に変換する
function decodePolyline6(encoded) {
  const coords = [];
  let index = 0, lat = 0, lng = 0;
  while (index < encoded.length) {
    let shift = 0, result = 0, b;
    do { b = encoded.charCodeAt(index++) - 63; result |= (b & 0x1f) << shift; shift += 5; } while (b >= 0x20);
    lat += (result & 1) ? ~(result >> 1) : (result >> 1);
    shift = 0; result = 0;
    do { b = encoded.charCodeAt(index++) - 63; result |= (b & 0x1f) << shift; shift += 5; } while (b >= 0x20);
    lng += (result & 1) ? ~(result >> 1) : (result >> 1);
    coords.push([lat / 1e6, lng / 1e6]);
  }
  return coords;
}

// OSRM（APIキー不要）でルートを取得する
// レスポンスに source と avoidanceApplied フラグを含める
async function fetchRouteOSRM(startLat, startLng, endLat, endLng, { avoidanceRequested = false, avoidanceApplied = false } = {}) {
  // OSRMは経度,緯度の順
  const url = `https://router.project-osrm.org/route/v1/driving/${startLng},${startLat};${endLng},${endLat}?overview=full&geometries=geojson`;
  const res = await fetch(url, {
    headers: { 'User-Agent': 'gas-split-app/1.0 (hobby project)' },
  });
  if (!res.ok) throw new Error(`OSRM HTTP ${res.status}`);
  const data = await res.json();
  if (data.code !== 'Ok' || !data.routes || !data.routes[0]) {
    throw new Error(`OSRM: ${data.code || 'ルートなし'}`);
  }
  const route = data.routes[0];
  const distanceKm = route.distance / 1000;
  // GeoJSONは[lng, lat]の順なので、Leafletで使う[lat, lng]の順に入れ替える
  const geometry = route.geometry.coordinates.map(([lng, lat]) => [lat, lng]);
  return new Response(
    JSON.stringify({ distanceKm, geometry, source: 'osrm', avoidanceRequested, avoidanceApplied }),
    { headers: { 'content-type': 'application/json; charset=utf-8' } }
  );
}


function jsonResponse(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { 'content-type': 'application/json; charset=utf-8' },
  });
}

// 0/O/1/I/l など見間違えやすい文字を除いたコード用アルファベット
const GROUP_CODE_ALPHABET = '23456789ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz';

function generateGroupCode(length = 8) {
  const bytes = new Uint8Array(length);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, (b) => GROUP_CODE_ALPHABET[b % GROUP_CODE_ALPHABET.length]).join('');
}

function numOrNull(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

// SupabaseのREST API（PostgREST）をService Role Keyで呼び出す共通ヘルパー。
// このキーはRLSをバイパスできるため、Worker内だけで使い、フロントには絶対に渡さない。
async function supabaseFetch(env, path, options = {}) {
  if (!env.SUPABASE_URL || !env.SUPABASE_SERVICE_ROLE_KEY) {
    throw new Error('SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY シークレットが設定されていません。');
  }
  return fetch(`${env.SUPABASE_URL}/rest/v1/${path}`, {
    ...options,
    headers: {
      apikey: env.SUPABASE_SERVICE_ROLE_KEY,
      Authorization: `Bearer ${env.SUPABASE_SERVICE_ROLE_KEY}`,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  });
}

// 共有コードからグループ行（id, code, name）を1件取得する。見つからなければnull。
async function getGroupByCode(code, env) {
  const res = await supabaseFetch(env, `groups?code=eq.${encodeURIComponent(code)}&select=id,code,name,created_at`);
  if (!res.ok) throw new Error(`Supabaseエラー（グループ取得）: HTTP ${res.status} ${await res.text()}`);
  const rows = await res.json();
  return rows[0] || null;
}

// POST /api/groups: グループを新規作成し、共有コードとメンバーを返す。
async function createGroup(request, env) {
  const body = await request.json().catch(() => ({}));
  const name = typeof body.name === 'string' && body.name.trim() ? body.name.trim().slice(0, 100) : null;
  const members = Array.isArray(body.members)
    ? body.members.map((m) => String(m).trim()).filter(Boolean).slice(0, 30)
    : [];

  let group = null;
  for (let attempt = 0; attempt < 5 && !group; attempt++) {
    const code = generateGroupCode();
    const res = await supabaseFetch(env, 'groups', {
      method: 'POST',
      headers: { Prefer: 'return=representation' },
      body: JSON.stringify({ code, name }),
    });
    if (res.status === 201) {
      const rows = await res.json();
      group = rows[0];
    } else if (res.status === 409) {
      continue; // コード重複（極めて稀）。別のコードで再試行。
    } else {
      throw new Error(`Supabaseエラー（グループ作成）: HTTP ${res.status} ${await res.text()}`);
    }
  }
  if (!group) throw new Error('グループコードの生成に失敗しました。もう一度お試しください。');

  if (members.length > 0) {
    const memberRows = members.map((memberName, i) => ({ group_id: group.id, name: memberName, sort_order: i }));
    const res = await supabaseFetch(env, 'group_members', {
      method: 'POST',
      headers: { Prefer: 'return=representation' },
      body: JSON.stringify(memberRows),
    });
    if (!res.ok) throw new Error(`Supabaseエラー（メンバー登録）: HTTP ${res.status} ${await res.text()}`);
  }

  return jsonResponse({
    id: group.id,
    code: group.code,
    name: group.name,
    members: members.map((memberName, i) => ({ name: memberName, sort_order: i })),
  });
}

// GET /api/groups/:code: グループ情報・メンバー一覧・履歴（最新50件）を返す。
async function getGroupInfo(code, env) {
  const group = await getGroupByCode(code, env);
  if (!group) return jsonResponse({ error: 'グループが見つかりません。' }, 404);

  const [membersRes, tripsRes] = await Promise.all([
    supabaseFetch(env, `group_members?group_id=eq.${group.id}&select=id,name,sort_order&order=sort_order.asc`),
    supabaseFetch(env, `trips?group_id=eq.${group.id}&select=*&order=created_at.desc&limit=50`),
  ]);
  if (!membersRes.ok) throw new Error(`Supabaseエラー（メンバー取得）: HTTP ${membersRes.status} ${await membersRes.text()}`);
  if (!tripsRes.ok) throw new Error(`Supabaseエラー（履歴取得）: HTTP ${tripsRes.status} ${await tripsRes.text()}`);

  const members = await membersRes.json();
  const trips = await tripsRes.json();

  return jsonResponse({ id: group.id, code: group.code, name: group.name, members, trips });
}

// PUT /api/groups/:code/members: メンバー一覧を丸ごと入れ替える（小規模なので削除→再挿入の単純な方式）。
async function updateGroupMembers(code, request, env) {
  const group = await getGroupByCode(code, env);
  if (!group) return jsonResponse({ error: 'グループが見つかりません。' }, 404);

  const body = await request.json().catch(() => ({}));
  const members = Array.isArray(body.members)
    ? body.members.map((m) => String(m).trim()).filter(Boolean).slice(0, 30)
    : [];

  const delRes = await supabaseFetch(env, `group_members?group_id=eq.${group.id}`, { method: 'DELETE' });
  if (!delRes.ok) throw new Error(`Supabaseエラー（メンバー削除）: HTTP ${delRes.status} ${await delRes.text()}`);

  if (members.length > 0) {
    const memberRows = members.map((memberName, i) => ({ group_id: group.id, name: memberName, sort_order: i }));
    const insRes = await supabaseFetch(env, 'group_members', {
      method: 'POST',
      headers: { Prefer: 'return=representation' },
      body: JSON.stringify(memberRows),
    });
    if (!insRes.ok) throw new Error(`Supabaseエラー（メンバー登録）: HTTP ${insRes.status} ${await insRes.text()}`);
  }

  return jsonResponse({ ok: true, members: members.map((memberName, i) => ({ name: memberName, sort_order: i })) });
}

// POST /api/groups/:code/trips: 精算結果を履歴として1件追加する。
async function createTrip(code, request, env) {
  const group = await getGroupByCode(code, env);
  if (!group) return jsonResponse({ error: 'グループが見つかりません。' }, 404);

  const body = await request.json().catch(() => ({}));
  const row = {
    group_id: group.id,
    start_label: body.startLabel || null,
    start_lat: numOrNull(body.startLat),
    start_lng: numOrNull(body.startLng),
    end_label: body.endLabel || null,
    end_lat: numOrNull(body.endLat),
    end_lng: numOrNull(body.endLng),
    distance_km: numOrNull(body.distanceKm),
    fuel_type: body.fuelType || null,
    unit_price_per_l: numOrNull(body.unitPricePerL),
    efficiency_km_per_l: numOrNull(body.efficiencyKmPerL),
    people_count: numOrNull(body.peopleCount),
    total_cost_yen: numOrNull(body.totalCostYen),
    per_person_yen: numOrNull(body.perPersonYen),
  };

  const res = await supabaseFetch(env, 'trips', {
    method: 'POST',
    headers: { Prefer: 'return=representation' },
    body: JSON.stringify(row),
  });
  if (!res.ok) throw new Error(`Supabaseエラー（履歴保存）: HTTP ${res.status} ${await res.text()}`);
  const rows = await res.json();
  return jsonResponse(rows[0], 201);
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
