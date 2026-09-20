const fs = require('fs');
const path = require('path');
const https = require('https');
const ExcelJS = require('exceljs');

const API_KEY = process.env.API_KEY || 'sk-ts-VB0BNV245K445QF7ZCRVCN6B7ADS';
const ENDPOINT = 'https://api.thirtystore.com/v1/chat/completions';
const INPUT_DIR = path.join(__dirname, 'input_foto_mobil');
const OUTPUT_DIR = path.join(__dirname, 'hasil_excel');

const MODELS = [
  'thirty/gpt-6-astra',
  'thirty/deepseek-v4-pro',
  'thirty/deepseek-v4.1-flash'
];

/**
 * Ekstraksi data spesifik lembar identitas mobil Kodim & Koperasi Desa
 */
function extractCarDetails(base64Image, filename) {
  const prompt = `Foto ini adalah lembaran kertas identitas kendaraan dinas/bantuan yang ditempel pada kaca mobil.
Teks dicetak dengan format seperti ini:
- Kodim / Satuan (misal: Kodim 1402/Polewali Mandar, Kodim 1407/Bone, Kodim 1417/Kendari)
- Model / Tipe Kendaraan (misal: YODHA, YODHA SINGLE CAB 4X4, TRUCK)
- Koperasi Desa / Penerima (misal: Koperasi Desa Campurjo, Koperasi Desa Pattimpa, Koperasi Desa Landono Dua)
- CH/NO : Nomor Rangka Chassis (misal: MAT464844TSR06479)
- ENGINE NO : Nomor Mesin (misal: VARICOR 16FTXJ11040, 16GTXJ13323)
- Wilayah / Provinsi (jika ada tertera di bawah, misal: Sulawesi Tenggara, Sulawesi Barat, Sulawesi Selatan)

TUGAS:
Baca teks pada kertas dengan sangat akurat dan persis huruf per huruf (terutama CH/NO dan ENGINE NO).
Kembalikan HANYA string JSON valid tanpa penjelasan markdown tambahan:
{
  "kodim": "Nama Kodim / Satuan",
  "tipe": "Model / Tipe Kendaraan (contoh: YODHA SINGLE CAB 4X4 atau TRUCK atau YODHA)",
  "koperasi": "Nama Koperasi Desa",
  "chassis_no": "Nomor CH/NO",
  "engine_no": "Nomor ENGINE NO",
  "wilayah": "Nama Wilayah / Provinsi jika ada, atau strip '-'"
}`;

  return new Promise(async (resolve, reject) => {
    let lastErr = null;

    for (const model of MODELS) {
      try {
        const payload = JSON.stringify({
          model: model,
          messages: [{
            role: 'user',
            content: [
              { type: 'text', text: prompt },
              {
                type: 'image_url',
                image_url: { url: `data:image/jpeg;base64,${base64Image}` }
              }
            ]
          }],
          temperature: 0.1,
          max_tokens: 500
        });

        const rawResp = await new Promise((res, rej) => {
          const req = https.request(ENDPOINT, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': `Bearer ${API_KEY}`,
              'User-Agent': 'Mozilla/5.0'
            },
            timeout: 35000
          }, (resp) => {
            let data = '';
            resp.on('data', c => data += c);
            resp.on('end', () => {
              if (resp.statusCode >= 200 && resp.statusCode < 300) {
                try {
                  const parsed = JSON.parse(data);
                  const content = parsed.choices?.[0]?.message?.content;
                  if (!content) return rej(new Error('Respon kosong'));
                  const l = content.toLowerCase();
                  if (l.includes('tidak dukung') || l.includes('gak bisa') || l.includes('image omitted')) {
                    return rej(new Error('Model menolak vision'));
                  }
                  res(content);
                } catch (e) {
                  rej(e);
                }
              } else {
                rej(new Error(`HTTP ${resp.statusCode}: ${data.slice(0, 100)}`));
              }
            });
          });

          req.on('timeout', () => { req.destroy(); rej(new Error('Timeout')); });
          req.on('error', rej);
          req.write(payload);
          req.end();
        });

        let clean = rawResp.trim();
        if (clean.includes('```json')) {
          clean = clean.split('```json')[1].split('```')[0].trim();
        } else if (clean.includes('```')) {
          clean = clean.split('```')[1].split('```')[0].trim();
        }
        const data = JSON.parse(clean);
        return resolve(data);
      } catch (err) {
        lastErr = err;
        console.warn(`   [WARN] Model ${model} gagal: ${err.message}. Mencoba model cadangan...`);
      }
    }

    reject(lastErr || new Error('Gagal mengekstrak data mobil dari foto'));
  });
}

/**
 * Styling bantu auto-fit kolom Excel
 */
function autoFit(worksheet) {
  worksheet.columns.forEach(col => {
    let max = 12;
    col.eachCell({ includeEmpty: false }, cell => {
      let str = '';
      if (cell.value !== null && cell.value !== undefined) {
        if (typeof cell.value === 'object' && cell.value.result) str = String(cell.value.result);
        else str = String(cell.value);
      }
      if (str.length > max) max = str.length;
    });
    col.width = Math.min(max + 4, 45);
  });
}

async function main() {
  console.log('================================================================');
  console.log('🚗 PEMROSESAN OTONOM 25 FOTO MOBIL KODIM & KOPERASI DESA');
  console.log('================================================================\n');

  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

  const validExts = ['.jpg', '.jpeg', '.png', '.webp'];
  const files = fs.readdirSync(INPUT_DIR).filter(f => {
    return validExts.includes(path.extname(f).toLowerCase());
  });

  if (files.length === 0) {
    console.log('❌ Tidak ada file foto di folder input_foto_mobil!');
    return;
  }

  console.log(`📁 Ditemukan ${files.length} foto kendaraan yang siap diproses.\n`);

  const extractedData = [];

  for (let i = 0; i < files.length; i++) {
    const filename = files[i];
    const filepath = path.join(INPUT_DIR, filename);
    const base64 = fs.readFileSync(filepath).toString('base64');

    console.log(`[${i + 1}/${files.length}] Membaca "${filename}"...`);
    try {
      const item = await extractCarDetails(base64, filename);
      item.filename = filename;
      extractedData.push(item);
      console.log(`   ✅ Kodim   : ${item.kodim || '-'}`);
      console.log(`      Tipe    : ${item.tipe || '-'}`);
      console.log(`      Koperasi: ${item.koperasi || '-'}`);
      console.log(`      NoRangka: ${item.chassis_no || '-'}`);
      console.log(`      NoMesin : ${item.engine_no || '-'}`);
    } catch (e) {
      console.error(`   ❌ Gagal membaca foto ke-${i + 1}:`, e.message);
    }
  }

  if (extractedData.length === 0) {
    console.error('\n❌ Tidak ada data yang berhasil diekstrak.');
    return;
  }

  console.log(`\n================================================================`);
  console.log(`📊 TOTAL KENDARAAN TERBACA: ${extractedData.length} DARI ${files.length} FOTO`);
  console.log(`================================================================\n`);

  // Simpan JSON mentah untuk backup
  const backupJsonPath = path.join(OUTPUT_DIR, 'hasil_ekstraksi_mentah.json');
  fs.writeFileSync(backupJsonPath, JSON.stringify(extractedData, null, 2));
  console.log(`💾 Backup data mentah tersimpan di: ${backupJsonPath}`);

  const now = new Date();
  const dateStr = now.toISOString().split('T')[0];

  // ================================================================
  // FILE 1: REKAP TERKELOMPOK PER KODIM (1 SHEET + SUBTOTAL UNIT)
  // ================================================================
  console.log('\n📄 Membuat File 1: REKAP_KENDARAAN_PER_KODIM.xlsx ...');
  const wb1 = new ExcelJS.Workbook();
  wb1.creator = 'ExcelBot AI Antigravity';
  wb1.created = now;
  const ws1 = wb1.addWorksheet('Rekap Per Kodim', { views: [{ showGridLines: true }] });

  // Header Title
  ws1.mergeCells('A1:G1');
  const t1 = ws1.getCell('A1');
  t1.value = 'DATA REKAPITULASI KENDARAAN DINAS / BANTUAN KOPERASI DESA';
  t1.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  t1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  t1.alignment = { horizontal: 'center', vertical: 'middle' };
  ws1.getRow(1).height = 32;

  ws1.mergeCells('A2:G2');
  const s1 = ws1.getCell('A2');
  s1.value = `Dikelompokkan Berdasarkan: Kodim / Satuan Wilayah  |  Tanggal: ${dateStr}  |  Total Unit: ${extractedData.length} Mobil`;
  s1.font = { name: 'Segoe UI', size: 10, italic: true, color: { argb: 'FF475569' } };
  s1.alignment = { horizontal: 'center', vertical: 'middle' };
  ws1.getRow(2).height = 20;

  ws1.getRow(3).height = 10;

  // Table Columns Header
  const colHeaders = ['No', 'Kodim / Satuan', 'Tipe / Model Kendaraan', 'Koperasi Desa Penerima', 'Nomor Rangka (CH/NO)', 'Nomor Mesin (ENGINE NO)', 'Wilayah / Provinsi'];
  const hRow1 = ws1.getRow(4);
  hRow1.values = colHeaders;
  hRow1.height = 26;
  hRow1.eachCell(c => {
    c.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'medium', color: { argb: 'FF1D4ED8' } }, right: { style: 'thin' } };
  });

  // Grouping by Kodim
  const kodimMap = new Map();
  extractedData.forEach(item => {
    const k = (item.kodim || 'KODIM LAINNYA').trim();
    if (!kodimMap.has(k)) kodimMap.set(k, []);
    kodimMap.get(k).push(item);
  });

  let rowIdx1 = 5;
  let unitCounter = 1;

  for (const [kodimName, items] of kodimMap.entries()) {
    // Group Header Row
    ws1.mergeCells(`A${rowIdx1}:G${rowIdx1}`);
    const gh = ws1.getCell(`A${rowIdx1}`);
    gh.value = `🏛️ KELOMPOK [ ${kodimName.toUpperCase()} ] — (${items.length} UNIT KENDARAAN)`;
    gh.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF0F172A' } };
    gh.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
    gh.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    ws1.getRow(rowIdx1).height = 24;
    rowIdx1++;

    // Group items
    items.forEach((item, idx) => {
      const r = ws1.getRow(rowIdx1);
      r.values = [
        unitCounter++,
        item.kodim || '-',
        item.tipe || '-',
        item.koperasi || '-',
        item.chassis_no || '-',
        item.engine_no || '-',
        item.wilayah || '-'
      ];
      r.height = 22;
      r.eachCell((c, cIdx) => {
        c.font = { name: 'Segoe UI', size: 10 };
        c.border = { top: { style: 'thin', color: { argb: 'FFE2E8F0' } }, left: { style: 'thin', color: { argb: 'FFE2E8F0' } }, bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } }, right: { style: 'thin', color: { argb: 'FFE2E8F0' } } };
        if (idx % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        if (cIdx === 1) c.alignment = { horizontal: 'center' };
        if (cIdx === 5 || cIdx === 6) c.font = { name: 'Consolas', size: 9.5, bold: true }; // Chassis & Engine font mono
      });
      rowIdx1++;
    });

    // Subtotal Row
    ws1.mergeCells(`A${rowIdx1}:F${rowIdx1}`);
    const stLabel = ws1.getCell(`A${rowIdx1}`);
    stLabel.value = `SUBTOTAL UNIT ${kodimName.toUpperCase()}:`;
    stLabel.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF1E40AF' } };
    stLabel.alignment = { horizontal: 'right', vertical: 'middle' };
    stLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF6FF' } };

    const stVal = ws1.getCell(`G${rowIdx1}`);
    stVal.value = `${items.length} Unit`;
    stVal.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF1E40AF' } };
    stVal.alignment = { horizontal: 'center', vertical: 'middle' };
    stVal.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF6FF' } };

    for (let c = 1; c <= 7; c++) {
      ws1.getRow(rowIdx1).getCell(c).border = { top: { style: 'thin', color: { argb: 'FFBFDBFE' } }, bottom: { style: 'thin', color: { argb: 'FFBFDBFE' } } };
    }
    ws1.getRow(rowIdx1).height = 22;
    rowIdx1++;

    // Spacer
    ws1.getRow(rowIdx1).height = 8;
    rowIdx1++;
  }

  // Grand Total
  ws1.mergeCells(`A${rowIdx1}:F${rowIdx1}`);
  const gtLabel = ws1.getCell(`A${rowIdx1}`);
  gtLabel.value = 'TOTAL KESELURUHAN UNIT ARMADA MOBIL:';
  gtLabel.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF0F172A' } };
  gtLabel.alignment = { horizontal: 'right', vertical: 'middle' };
  gtLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };

  const gtVal = ws1.getCell(`G${rowIdx1}`);
  gtVal.value = `${extractedData.length} Unit Mobil`;
  gtVal.font = { name: 'Segoe UI', size: 12, bold: true, color: { argb: 'FF059669' } };
  gtVal.alignment = { horizontal: 'center', vertical: 'middle' };
  gtVal.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFECFDF5' } };
  for (let c = 1; c <= 7; c++) {
    ws1.getRow(rowIdx1).getCell(c).border = { top: { style: 'medium', color: { argb: 'FF94A3B8' } }, bottom: { style: 'double', color: { argb: 'FF0F172A' } } };
  }
  ws1.getRow(rowIdx1).height = 28;

  autoFit(ws1);
  const outPath1 = path.join(OUTPUT_DIR, `REKAP_KENDARAAN_PER_KODIM_${dateStr}.xlsx`);
  await wb1.xlsx.writeFile(outPath1);
  console.log(`   ✅ File 1 selesai dibuat: ${outPath1}`);

  // ================================================================
  // FILE 2: REKAP TERKELOMPOK PER TIPE KENDARAAN (YODHA / TRUCK)
  // ================================================================
  console.log('\n📄 Membuat File 2: REKAP_KENDARAAN_PER_TIPE_MOBIL.xlsx ...');
  const wb2 = new ExcelJS.Workbook();
  wb2.creator = 'ExcelBot AI Antigravity';
  wb2.created = now;
  const ws2 = wb2.addWorksheet('Rekap Per Tipe Mobil', { views: [{ showGridLines: true }] });

  ws2.mergeCells('A1:G1');
  const t2 = ws2.getCell('A1');
  t2.value = 'DATA REKAPITULASI KENDARAAN (PER MODEL / TIPE UNIT)';
  t2.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  t2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF065F46' } }; // Hijau tua emerald
  t2.alignment = { horizontal: 'center', vertical: 'middle' };
  ws2.getRow(1).height = 32;

  ws2.mergeCells('A2:G2');
  const s2 = ws2.getCell('A2');
  s2.value = `Dikelompokkan Berdasarkan: Model / Tipe Kendaraan  |  Tanggal: ${dateStr}  |  Total: ${extractedData.length} Unit`;
  s2.font = { name: 'Segoe UI', size: 10, italic: true, color: { argb: 'FF475569' } };
  s2.alignment = { horizontal: 'center', vertical: 'middle' };
  ws2.getRow(2).height = 20;

  ws2.getRow(3).height = 10;

  const hRow2 = ws2.getRow(4);
  hRow2.values = ['No', 'Tipe / Model Kendaraan', 'Kodim / Satuan', 'Koperasi Desa Penerima', 'Nomor Rangka (CH/NO)', 'Nomor Mesin (ENGINE NO)', 'Wilayah'];
  hRow2.height = 26;
  hRow2.eachCell(c => {
    c.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF059669' } };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'medium', color: { argb: 'FF047857' } }, right: { style: 'thin' } };
  });

  const tipeMap = new Map();
  extractedData.forEach(item => {
    const t = (item.tipe || 'TIPE LAINNYA').trim();
    if (!tipeMap.has(t)) tipeMap.set(t, []);
    tipeMap.get(t).push(item);
  });

  let rowIdx2 = 5;
  let counter2 = 1;

  for (const [tipeName, items] of tipeMap.entries()) {
    ws2.mergeCells(`A${rowIdx2}:G${rowIdx2}`);
    const gh = ws2.getCell(`A${rowIdx2}`);
    gh.value = `🚗 MODEL / TIPE [ ${tipeName.toUpperCase()} ] — (${items.length} UNIT)`;
    gh.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF064E3B' } };
    gh.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
    gh.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    ws2.getRow(rowIdx2).height = 24;
    rowIdx2++;

    items.forEach((item, idx) => {
      const r = ws2.getRow(rowIdx2);
      r.values = [
        counter2++,
        item.tipe || '-',
        item.kodim || '-',
        item.koperasi || '-',
        item.chassis_no || '-',
        item.engine_no || '-',
        item.wilayah || '-'
      ];
      r.height = 22;
      r.eachCell((c, cIdx) => {
        c.font = { name: 'Segoe UI', size: 10 };
        c.border = { top: { style: 'thin', color: { argb: 'FFE2E8F0' } }, left: { style: 'thin', color: { argb: 'FFE2E8F0' } }, bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } }, right: { style: 'thin', color: { argb: 'FFE2E8F0' } } };
        if (idx % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0FDF4' } };
        if (cIdx === 1) c.alignment = { horizontal: 'center' };
        if (cIdx === 5 || cIdx === 6) c.font = { name: 'Consolas', size: 9.5, bold: true };
      });
      rowIdx2++;
    });

    // Subtotal
    ws2.mergeCells(`A${rowIdx2}:F${rowIdx2}`);
    const stLabel = ws2.getCell(`A${rowIdx2}`);
    stLabel.value = `SUBTOTAL MODEL ${tipeName.toUpperCase()}:`;
    stLabel.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF065F46' } };
    stLabel.alignment = { horizontal: 'right', vertical: 'middle' };
    stLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFECFDF5' } };

    const stVal = ws2.getCell(`G${rowIdx2}`);
    stVal.value = `${items.length} Unit`;
    stVal.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF065F46' } };
    stVal.alignment = { horizontal: 'center', vertical: 'middle' };
    stVal.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFECFDF5' } };

    for (let c = 1; c <= 7; c++) {
      ws2.getRow(rowIdx2).getCell(c).border = { top: { style: 'thin', color: { argb: 'FFA7F3D0' } }, bottom: { style: 'thin', color: { argb: 'FFA7F3D0' } } };
    }
    ws2.getRow(rowIdx2).height = 22;
    rowIdx2++;

    ws2.getRow(rowIdx2).height = 8;
    rowIdx2++;
  }

  // Grand Total
  ws2.mergeCells(`A${rowIdx2}:F${rowIdx2}`);
  const gtLabel2 = ws2.getCell(`A${rowIdx2}`);
  gtLabel2.value = 'TOTAL KESELURUHAN UNIT SEMUA TIPE:';
  gtLabel2.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF0F172A' } };
  gtLabel2.alignment = { horizontal: 'right', vertical: 'middle' };
  gtLabel2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };

  const gtVal2 = ws2.getCell(`G${rowIdx2}`);
  gtVal2.value = `${extractedData.length} Unit Mobil`;
  gtVal2.font = { name: 'Segoe UI', size: 12, bold: true, color: { argb: 'FF059669' } };
  gtVal2.alignment = { horizontal: 'center', vertical: 'middle' };
  gtVal2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFECFDF5' } };
  for (let c = 1; c <= 7; c++) {
    ws2.getRow(rowIdx2).getCell(c).border = { top: { style: 'medium', color: { argb: 'FF94A3B8' } }, bottom: { style: 'double', color: { argb: 'FF0F172A' } } };
  }
  ws2.getRow(rowIdx2).height = 28;

  autoFit(ws2);
  const outPath2 = path.join(OUTPUT_DIR, `REKAP_KENDARAAN_PER_TIPE_MOBIL_${dateStr}.xlsx`);
  await wb2.xlsx.writeFile(outPath2);
  console.log(`   ✅ File 2 selesai dibuat: ${outPath2}`);

  // ================================================================
  // FILE 3: REKAP MULTI-TAB (TAB PER KODIM + TAB RINGKASAN REKAP)
  // ================================================================
  console.log('\n📄 Membuat File 3: REKAP_KENDARAAN_MULTITAB_LENGKAP.xlsx ...');
  const wb3 = new ExcelJS.Workbook();
  wb3.creator = 'ExcelBot AI Antigravity';
  wb3.created = now;

  // 1. Tab Ringkasan / Dashboard
  const sumWs = wb3.addWorksheet('RINGKASAN_REKAP', { views: [{ showGridLines: true }] });
  sumWs.mergeCells('A1:D1');
  const stitle = sumWs.getCell('A1');
  stitle.value = 'RINGKASAN ALOKASI KENDARAAN PER KODIM';
  stitle.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  stitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  stitle.alignment = { horizontal: 'center', vertical: 'middle' };
  sumWs.getRow(1).height = 32;

  const sHeaders = sumWs.getRow(3);
  sHeaders.values = ['No', 'Nama Kodim / Satuan', 'Jumlah Unit Mobil', 'Keterangan'];
  sHeaders.height = 24;
  sHeaders.eachCell(c => {
    c.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  let sumRowIdx = 4;
  let sNo = 1;
  for (const [kodimName, items] of kodimMap.entries()) {
    const r = sumWs.getRow(sumRowIdx);
    r.values = [sNo++, kodimName, `${items.length} Unit`, `Tersebar di ${items.length} Koperasi Desa`];
    r.height = 22;
    r.getCell(1).alignment = { horizontal: 'center' };
    r.getCell(3).alignment = { horizontal: 'center', bold: true };
    sumRowIdx++;
  }

  const sGrand = sumWs.getRow(sumRowIdx);
  sumWs.mergeCells(`A${sumRowIdx}:B${sumRowIdx}`);
  sGrand.getCell(1).value = 'TOTAL KESELURUHAN ARMADA:';
  sGrand.getCell(1).font = { bold: true };
  sGrand.getCell(1).alignment = { horizontal: 'right' };
  sGrand.getCell(3).value = `${extractedData.length} Unit Mobil`;
  sGrand.getCell(3).font = { bold: true, color: { argb: 'FF059669' }, size: 11 };
  sGrand.getCell(3).alignment = { horizontal: 'center' };
  autoFit(sumWs);

  // 2. Tab untuk masing-masing Kodim
  let sheetNum = 1;
  for (const [kodimName, items] of kodimMap.entries()) {
    const cleanSheetName = kodimName.replace(/[\\/*?:[\]]/g, '').slice(0, 24);
    const tabName = `${sheetNum++}. ${cleanSheetName}`.slice(0, 31);
    const ws = wb3.addWorksheet(tabName, { views: [{ showGridLines: true }] });

    ws.mergeCells('A1:F1');
    const tCell = ws.getCell('A1');
    tCell.value = `DATA KENDARAAN — ${kodimName.toUpperCase()}`;
    tCell.font = { name: 'Segoe UI', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
    tCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    tCell.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(1).height = 28;

    const hr = ws.getRow(3);
    hr.values = ['No', 'Tipe Kendaraan', 'Koperasi Desa Penerima', 'Nomor Rangka (CH/NO)', 'Nomor Mesin (ENGINE NO)', 'Wilayah'];
    hr.height = 24;
    hr.eachCell(c => {
      c.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF475569' } };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    items.forEach((item, idx) => {
      const r = ws.getRow(4 + idx);
      r.values = [
        idx + 1,
        item.tipe || '-',
        item.koperasi || '-',
        item.chassis_no || '-',
        item.engine_no || '-',
        item.wilayah || '-'
      ];
      r.height = 21;
      r.getCell(1).alignment = { horizontal: 'center' };
      r.getCell(4).font = { name: 'Consolas', size: 9.5, bold: true };
      r.getCell(5).font = { name: 'Consolas', size: 9.5, bold: true };
    });

    autoFit(ws);
  }

  // 3. Tab Master Keseluruhan
  const allWs = wb3.addWorksheet('DATA_MASTER_SEMUA_UNIT', { views: [{ showGridLines: true }] });
  const allH = allWs.getRow(1);
  allH.values = ['No', 'Kodim / Satuan', 'Tipe Kendaraan', 'Koperasi Desa', 'No. Rangka (CH/NO)', 'No. Mesin (ENGINE NO)', 'Wilayah', 'Nama File Foto'];
  allH.height = 26;
  allH.eachCell(c => {
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    c.alignment = { horizontal: 'center' };
  });

  extractedData.forEach((item, idx) => {
    const r = allWs.getRow(2 + idx);
    r.values = [
      idx + 1,
      item.kodim || '-',
      item.tipe || '-',
      item.koperasi || '-',
      item.chassis_no || '-',
      item.engine_no || '-',
      item.wilayah || '-',
      item.filename || '-'
    ];
    r.height = 20;
    r.getCell(1).alignment = { horizontal: 'center' };
    r.getCell(5).font = { name: 'Consolas', size: 9 };
    r.getCell(6).font = { name: 'Consolas', size: 9 };
  });
  autoFit(allWs);

  const outPath3 = path.join(OUTPUT_DIR, `REKAP_KENDARAAN_MULTITAB_LENGKAP_${dateStr}.xlsx`);
  await wb3.xlsx.writeFile(outPath3);
  console.log(`   ✅ File 3 selesai dibuat: ${outPath3}`);

  console.log('\n================================================================');
  console.log('🎉 SEMUA 25 FOTO MOBIL BERHASIL DIEKSTRAK & DIBUATKAN EXCEL!');
  console.log(`📂 Seluruh file tersimpan rapi di: ${OUTPUT_DIR}`);
  console.log('================================================================\n');
}

main().catch(err => console.error('FATAL ERROR:', err));
