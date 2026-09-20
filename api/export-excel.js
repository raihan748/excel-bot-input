const { generateSingleWorkbook } = require('../lib/excelGenerator');

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
    const { tableData, filename = 'tabel_hasil_ai.xlsx' } = req.body || {};

    if (!tableData || !tableData.columns) {
      return res.status(400).json({ error: 'Data tabel tidak valid atau kosong.' });
    }

    const excelBuffer = await generateSingleWorkbook(tableData);

    const safeFilename = encodeURIComponent(filename.replace(/[^a-zA-Z0-9_\-.]/g, '_'));

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${safeFilename}"`);
    res.setHeader('Content-Length', excelBuffer.length);

    return res.status(200).send(excelBuffer);
  } catch (error) {
    console.error('Error in /api/export-excel:', error);
    return res.status(500).json({
      error: error.message || 'Gagal membuat file Excel'
    });
  }
};
