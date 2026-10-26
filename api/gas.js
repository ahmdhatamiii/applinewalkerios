export default async function handler(req, res) {
  // URL Web App GAS - deployment @85 aktif dengan doGet API router
  const GAS_URL = "https://script.google.com/macros/s/AKfycbwDSnMqpHCO3Yg_YHjizrx_588KKXFFEbZAuY7tprarOE_f0HhlnkdZbRRZEkdr4dE0/exec";

  const action = req.query.action || '';
  const params = req.query.params || '[]';

  if (!action) {
    return res.status(400).json({ status: 'error', message: 'Parameter action wajib' });
  }

  const url = `${GAS_URL}?action=${encodeURIComponent(action)}&params=${encodeURIComponent(params)}`;

  try {
    // Step 1: GET ke GAS URL, tangkap redirect manual
    let response = await fetch(url, { redirect: 'follow' });
    let text = await response.text();

    // Step 2: Jika Google mengembalikan HTML (redirect page), coba extract redirect URL
    if (text.includes('<!DOCTYPE html') || text.includes('<!doctype html')) {
      // Google sering pakai meta refresh atau JS redirect
      const metaMatch = text.match(/content=["']0;\s*url=(.*?)["']/i);
      const jsMatch = text.match(/location\.(?:replace|href)\s*[\(=]\s*["'](https:\/\/[^"']+)["']/);
      const redirectUrl = metaMatch && metaMatch[1] ? metaMatch[1] : (jsMatch && jsMatch[1] ? jsMatch[1] : null);

      if (redirectUrl) {
        const redir = await fetch(redirectUrl, { redirect: 'follow' });
        text = await redir.text();
      }
    }

    // Step 3: Coba parse JSON
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Content-Type', 'application/json');

    if (req.method === 'OPTIONS') {
      return res.status(200).end();
    }

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
