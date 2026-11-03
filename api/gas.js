export default async function handler(req, res) {
  // URL Web App GAS - deployment @85 aktif dengan doGet dan doPost API router
  const GAS_URL = "https://script.google.com/macros/s/AKfycbwDSnMqpHCO3Yg_YHjizrx_588KKXFFEbZAuY7tprarOE_f0HhlnkdZbRRZEkdr4dE0/exec";

  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Accept');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  let action = '';
  let params = [];

  if (req.method === 'POST') {
    const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
    action = body.action || req.query.action || '';
    params = body.params || (req.query.params ? JSON.parse(req.query.params) : []);
  } else {
    action = req.query.action || '';
    if (req.query.params) {
      try { params = JSON.parse(req.query.params); } catch(e) { params = [req.query.params]; }
    }
  }

  if (!action) {
    return res.status(400).json({ status: 'error', message: 'Parameter action wajib' });
  }

  try {
    let fetchOptions = { redirect: 'follow' };
    let fetchUrl = GAS_URL;

    if (req.method === 'POST') {
      fetchOptions = {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action, params }),
        redirect: 'follow'
      };
    } else {
      fetchUrl = `${GAS_URL}?action=${encodeURIComponent(action)}&params=${encodeURIComponent(JSON.stringify(params))}`;
    }

    let response = await fetch(fetchUrl, fetchOptions);
    let text = await response.text();

    // Jika Google mengembalikan HTML (redirect page), coba extract redirect URL
    if (text.includes('<!DOCTYPE html') || text.includes('<!doctype html')) {
      const metaMatch = text.match(/content=["']0;\s*url=(.*?)["']/i);
      const jsMatch = text.match(/location\.(?:replace|href)\s*[\(=]\s*["'](https:\/\/[^"']+)["']/);
      const redirectUrl = metaMatch && metaMatch[1] ? metaMatch[1] : (jsMatch && jsMatch[1] ? jsMatch[1] : null);

      if (redirectUrl) {
        const redir = await fetch(redirectUrl, { redirect: 'follow' });
        text = await redir.text();
      }
    }

    res.setHeader('Content-Type', 'application/json');

    try {
      JSON.parse(text); // validasi JSON
      return res.status(200).send(text);
    } catch (e) {
      return res.status(502).json({
        status: 'error',
        message: 'GAS tidak mengembalikan JSON. Pastikan deployment Web App sudah di-update dan akses = Anyone.',
        preview: text.slice(0, 300)
      });
    }
  } catch (error) {
    return res.status(500).json({ status: 'error', message: error.message });
  }
}
