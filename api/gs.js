export const config = { runtime: 'edge' };

const GAS_URL = 'https://script.google.com/macros/s/AKfycbyeW_1LFg6GNEPnpQZwI2a8VxAHhhEicm2vlyjbFbvmeY84we2c4C_FaPonZXg2Z1DoPg/exec';
const FETCH_TIMEOUT = 55000; // 55초 (Edge 제한 60초 이내)

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Cache-Control': 'no-store',
};

function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json', ...CORS_HEADERS },
  });
}

async function fetchGAS(url, options = {}) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), FETCH_TIMEOUT);
  try {
    const res = await fetch(url, { ...options, signal: controller.signal, redirect: 'follow' });
    const body = await res.text();
    return new Response(body, {
      status: 200,
      headers: { 'Content-Type': 'application/json', ...CORS_HEADERS },
    });
  } catch (err) {
    if (err.name === 'AbortError') {
      return jsonResponse({ ok: false, error: '서버 응답 시간 초과. 잠시 후 다시 시도해주세요.' }, 504);
    }
    return jsonResponse({ ok: false, error: '서버 연결 실패: ' + err.message }, 502);
  } finally {
    clearTimeout(timer);
  }
}

export default async function handler(req) {
  if (req.method === 'OPTIONS') {
    return new Response(null, { status: 204, headers: CORS_HEADERS });
  }

  if (req.method === 'GET') {
    const url = new URL(req.url);
    const gasUrl = new URL(GAS_URL);
    url.searchParams.forEach((v, k) => gasUrl.searchParams.set(k, v));
    return fetchGAS(gasUrl.toString());
  }

  if (req.method === 'POST') {
    const body = await req.text();
    return fetchGAS(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body,
    });
  }

  return jsonResponse({ error: 'Method not allowed' }, 405);
}
