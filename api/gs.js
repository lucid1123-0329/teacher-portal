export const config = { runtime: 'edge' };

const GAS_URL = 'https://script.google.com/macros/s/AKfycbyeW_1LFg6GNEPnpQZwI2a8VxAHhhEicm2vlyjbFbvmeY84we2c4C_FaPonZXg2Z1DoPg/exec';

export default async function handler(req) {
  const url = new URL(req.url);
  
  if (req.method === 'GET') {
    const gasUrl = new URL(GAS_URL);
    url.searchParams.forEach((v, k) => gasUrl.searchParams.set(k, v));
    
    const res = await fetch(gasUrl.toString(), { redirect: 'follow' });
    const body = await res.text();
    
    return new Response(body, {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
        'Cache-Control': 'no-store',
      },
    });
  }
  
  if (req.method === 'POST') {
    const body = await req.text();
    
    const res = await fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body,
      redirect: 'follow',
    });
    const resBody = await res.text();
    
    return new Response(resBody, {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
        'Cache-Control': 'no-store',
      },
    });
  }
  
  if (req.method === 'OPTIONS') {
    return new Response(null, {
      status: 204,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
      },
    });
  }
  
  return new Response(JSON.stringify({ error: 'Method not allowed' }), { status: 405 });
}
