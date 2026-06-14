// CommonJS + built-in https: works on any Vercel Node.js runtime without package.json
const https = require('https');

module.exports = function (req, res) {
  const q = (req.query || {}).q;
  const type = (req.query || {}).type || 'address';

  if (!q) { res.status(400).json({ error: 'no query' }); return; }

  const ep = type === 'keyword' ? 'search/keyword' : 'search/address';
  const path = `/v2/local/${ep}.json?query=${encodeURIComponent(q)}&size=1`;

  const req2 = https.request(
    { hostname: 'dapi.kakao.com', path, method: 'GET',
      headers: { Authorization: 'KakaoAK ecb0c29ccb572cc5a0d4280e73beacdd' } },
    (r) => {
      let body = '';
      r.on('data', d => { body += d; });
      r.on('end', () => {
        res.setHeader('Cache-Control', 's-maxage=86400, stale-while-revalidate');
        try { res.json(JSON.parse(body)); }
        catch (e) { res.status(500).json({ error: 'parse error', raw: body.slice(0, 200) }); }
      });
    }
  );
  req2.on('error', e => res.status(500).json({ error: e.message }));
  req2.end();
};
