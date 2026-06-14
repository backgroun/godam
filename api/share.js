// Proxy to jsonblob.com for cross-device plan sharing
const https = require('https');

function req(opts, body) {
  return new Promise((resolve, reject) => {
    const r = https.request(opts, res => {
      let d = '';
      res.on('data', c => d += c);
      res.on('end', () => resolve({ status: res.statusCode, headers: res.headers, body: d }));
    });
    r.on('error', reject);
    if (body) r.write(body);
    r.end();
  });
}

module.exports = function(request, response) {
  response.setHeader('Access-Control-Allow-Origin', '*');

  if (request.method === 'POST') {
    let body = '';
    request.on('data', c => body += c);
    request.on('end', () => {
      req({
        hostname: 'jsonblob.com',
        path: '/api/jsonBlob',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'Content-Length': Buffer.byteLength(body)
        }
      }, body).then(r => {
        const location = r.headers['location'] || '';
        const id = location.split('/').pop();
        if (id) {
          response.json({ id });
        } else {
          response.status(500).json({ error: 'no id', status: r.status, body: r.body.slice(0, 200) });
        }
      }).catch(e => response.status(500).json({ error: e.message }));
    });

  } else if (request.method === 'GET') {
    const id = (request.query || {}).id;
    if (!id) { response.status(400).json({ error: 'no id' }); return; }
    req({
      hostname: 'jsonblob.com',
      path: `/api/jsonBlob/${id}`,
      method: 'GET',
      headers: { 'Accept': 'application/json' }
    }).then(r => {
      response.setHeader('Cache-Control', 's-maxage=3600, stale-while-revalidate');
      try { response.json(JSON.parse(r.body)); }
      catch(e) { response.status(500).json({ error: 'parse error', raw: r.body.slice(0, 200) }); }
    }).catch(e => response.status(500).json({ error: e.message }));

  } else {
    response.status(405).end();
  }
};
