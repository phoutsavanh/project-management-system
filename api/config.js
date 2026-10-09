module.exports = (req, res) => {
  res.setHeader('Cache-Control', 's-maxage=60, stale-while-revalidate');
  res.status(200).json({
    apiUrl: process.env.GAS_API_URL || process.env.NEXT_PUBLIC_GAS_API_URL || ''
  });
};
