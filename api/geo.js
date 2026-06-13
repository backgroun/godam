export default async function handler(req, res) {
  const { q, type = 'address' } = req.query;
  if (!q) return res.status(400).json({ error: 'no query' });

  const ep = type === 'keyword' ? 'search/keyword' : 'search/address';
  try {
    const r = await fetch(
      `https://dapi.kakao.com/v2/local/${ep}.json?query=${encodeURIComponent(q)}&size=1`,
      { headers: { 'Authorization': 'KakaoAK ecb0c29ccb572cc5a0d4280e73beacdd' } }
    );
    const data = await r.json();
    res.setHeader('Cache-Control', 's-maxage=86400, stale-while-revalidate');
    return res.json(data);
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
}
