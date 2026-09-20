const { generateMergedWorkbook } = require('../lib/excelGenerator');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Credentials', true);
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Hanya menerima request POST' });
  }

  try {
    const {
      datasets,
      mode = 'append', // 'append' atau 'multi-sheet'
      filename = 'gabungan_rekap_excel.xlsx',
      title = 'MASTER REKAP GABUNGAN EXCEL'
    } = req.body || {};

    if (!Array.isArray(datasets) || datasets.length < 2) {
      return res.status(400).json({ error: 'Minimal pilih 2 file untuk digabungkan.' });
    }

    const excelBuffer = await generateMergedWorkbook(datasets, mode, title);

    const safeFilename = encodeURIComponent(filename.replace(/[^a-zA-Z0-9_\-.]/g, '_'));

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${safeFilename}"`);
    res.setHeader('Content-Length', excelBuffer.length);

    return res.status(200).send(excelBuffer);
  } catch (error) {
    console.error('Error in /api/merge-excel:', error);
    return res.status(500).json({
      error: error.message || 'Gagal menggabungkan file Excel'
    });
  }
};
