const fs = require('fs');
const path = require('path');
const https = require('https');
const ExcelJS = require('exceljs');
const { generateSingleWorkbook } = require('./lib/excelGenerator');

const API_KEY = process.env.API_KEY || 'sk-ts-VB0BNV245K445QF7ZCRVCN6B7ADS';
const ENDPOINT = 'https://api.thirtystore.com/v1/chat/completions';
const INPUT_DIR = path.join(__dirname, 'input_foto_mobil');
const OUTPUT_DIR = path.join(__dirname, 'hasil_excel');

// Model Vision
const MODELS = [
  'thirty/gpt-6-astra',
  'thirty/deepseek-v4-pro',
  'thirty/deepseek-v4.1-flash'
];

function callVision(base64Image, mimeType, filename) {
  const systemPrompt = `Anda adalah asisten AI spesialis ekstraksi dokumen, nota, buku kas, dan daftar armada kendaraan / mobil.

TUGAS:
1. Baca teks foto ini dengan sangat teliti.
2. Identifikasi data unit mobil (Merk, Tipe Mobil, Plat Nomor jika ada, Penyewa/Supir/Pelanggan/Uraian Servis, dan Biaya/Tarif/Nominal).
3. Kolom yang diwajibkan: ["No", "Mobil / Unit", "Plat Nomor", "Penyewa / Driver / Uraian", "Tarif / Biaya", "Total"]
   (Boleh disesuaikan jika isi dokumen spesifik berbeda, tapi kolom nama Mobil WAJIB ada di kolom index ke-1 agar bisa dikelompokkan).
4. Pastikan angka nominal hanya berupa integer murni tanpa tulisan Rp/titik/koma.

OUTPUT HANYA JSON VALID:
{
  "document_title": "Judul data / toko / rental",
  "document_type": "Data Mobil",
  "date": "YYYY-MM-DD",
  "columns": ["No", "Mobil / Kendaraan", "Plat Nomor", "Penyewa / Driver / Uraian", "Tarif / Biaya", "Total"],
  "rows": [
    [1, "Toyota Avanza", "B 1234 CD", "Sewa 2 Hari", 800000, 800000]
  ],
  "total": 800000
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
              { type: 'text', text: systemPrompt },
              {
                type: 'image_url',
                image_url: { url: `data:${mimeType};base64,${base64Image}` }
              }
            ]
          }],
          temperature: 0.1,
          max_tokens: 2500
        });

        const result = await new Promise((res, rej) => {
          const req = https.request(ENDPOINT, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': `Bearer ${API_KEY}`,
              'User-Agent': 'Mozilla/5.0'
            },
            timeout: 45000
          }, (resp) => {
            let data = '';
            resp.on('data', c => data += c);
            resp.on('end', () => {
              if (resp.statusCode >= 200 && resp.statusCode < 300) {
                try {
                  const json = JSON.parse(data);
                  const content = json.choices?.[0]?.message?.content;
                  if (!content) return rej(new Error('Respon teks kosong'));
                  
                  const lower = content.toLowerCase();
                  if (lower.includes('tidak dukung') || lower.includes('gak bisa') || lower.includes('image omitted')) {
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

        // Ekstrak JSON
        let clean = result.trim();
        if (clean.includes('```json')) {
          clean = clean.split('```json')[1].split('```')[0].trim();
        } else if (clean.includes('```')) {
          clean = clean.split('```')[1].split('```')[0].trim();
        }
        const parsed = JSON.parse(clean);
        return resolve(parsed);
      } catch (err) {
        lastErr = err;
        console.warn(`   [WARN] Model ${model} gagal pada ${filename}: ${err.message}. Mencoba cadangan...`);
      }
    }

    reject(lastErr || new Error('Semua model vision gagal'));
  });
}

async function main() {
  console.log('====================================================');
  console.log('🚗 PROSES OTOMATIS FOTO MOBIL KE EXCEL TERKELOMPOK');
  console.log('====================================================\n');

  if (!fs.existsSync(INPUT_DIR)) fs.mkdirSync(INPUT_DIR, { recursive: true });
  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

  const validExts = ['.jpg', '.jpeg', '.png', '.webp', '.bmp'];
  const files = fs.readdirSync(INPUT_DIR).filter(f => {
    const ext = path.extname(f).toLowerCase();
    return validExts.includes(ext);
  });

  if (files.length === 0) {
    console.log(`📁 Belum ada foto di dalam folder:\n   ${INPUT_DIR}\n`);
    console.log('👉 Masukkan foto-foto mobil (JPG, PNG) ke dalam folder tersebut,');
    console.log('   lalu jalankan lagi script ini atau beritahu saya!');
    return;
  }

  console.log(`🔍 Ditemukan ${files.length} file foto mobil:`);
  files.forEach((f, i) => console.log(`   [${i+1}] ${f}`));
  console.log('\n⏳ Memulai pembacaan AI Vision...\n');

  let combinedColumns = ['No', 'Mobil / Kendaraan', 'Plat Nomor', 'Penyewa / Driver / Uraian', 'Tarif / Biaya', 'Total'];
  let allRows = [];
  let docTitle = 'REKAP ARMADA MOBIL';
  let totalSum = 0;

  for (let i = 0; i < files.length; i++) {
    const filename = files[i];
    const filepath = path.join(INPUT_DIR, filename);
    const ext = path.extname(filename).toLowerCase();
    const mime = ext === '.png' ? 'image/png' : 'image/jpeg';
    const base64 = fs.readFileSync(filepath).toString('base64');

    console.log(`[${i+1}/${files.length}] Memproses "${filename}"...`);
    try {
      const data = await callVision(base64, mime, filename);
      if (i === 0 && data.columns && data.columns.length > 0) {
        combinedColumns = data.columns;
        if (data.document_title) docTitle = data.document_title;
      }
      if (Array.isArray(data.rows)) {
        data.rows.forEach(r => allRows.push(r));
        console.log(`   ✅ Berhasil ekstrak ${data.rows.length} baris data`);
      }
      if (typeof data.total === 'number') totalSum += data.total;
    } catch (err) {
      console.error(`   ❌ Gagal pada ${filename}:`, err.message);
    }
  }

  if (allRows.length === 0) {
    console.error('\n❌ Tidak ada baris data yang berhasil diekstrak.');
    return;
  }

  // Normalisasi kolom nomor
  allRows.forEach((r, idx) => {
    if (typeof r[0] === 'number') r[0] = idx + 1;
  });

  // Cari kolom nama mobil
  let mobilColIdx = combinedColumns.findIndex(c => {
    const l = String(c).toLowerCase();
    return l.includes('mobil') || l.includes('kendaraan') || l.includes('plat') || l.includes('tipe');
  });
  if (mobilColIdx === -1 && combinedColumns.length > 1) mobilColIdx = 1;

  console.log(`\n📌 Dikelompokkan berdasarkan: Kolom "${combinedColumns[mobilColIdx]}" (Index ${mobilColIdx})`);
  console.log(`📊 Total Baris Data Terkumpul: ${allRows.length} baris\n`);

  const now = new Date();
  const dateStr = now.toISOString().split('T')[0];
  const timeStr = `${now.getHours()}${now.getMinutes()}${now.getSeconds()}`;

  const tableData = {
    document_title: docTitle,
    document_type: 'Data Mobil',
    date: dateStr,
    columns: combinedColumns,
    rows: allRows,
    total: totalSum
  };

  // 1. Generate Versi 1: Single Sheet Terkelompok dengan Subtotal
  const file1Name = `REKAP_MOBIL_TERKELOMPOK_${dateStr}_${timeStr}.xlsx`;
  const file1Path = path.join(OUTPUT_DIR, file1Name);
  const buf1 = await generateSingleWorkbook(tableData, {
    isGrouped: true,
    groupColIdx: mobilColIdx,
    groupMode: 'single-sheet'
  });
  fs.writeFileSync(file1Path, buf1);
  console.log(`✅ [1/2] File Excel Terkelompok (1 Sheet + Subtotal) Berhasil Dibuat:`);
  console.log(`   📄 ${file1Path}`);

  // 2. Generate Versi 2: Multi-Tab (Tiap Mobil Punya Tab Sendiri + Ringkasan)
  const file2Name = `REKAP_MOBIL_MULTITAB_${dateStr}_${timeStr}.xlsx`;
  const file2Path = path.join(OUTPUT_DIR, file2Name);
  const buf2 = await generateSingleWorkbook(tableData, {
    isGrouped: true,
    groupColIdx: mobilColIdx,
    groupMode: 'multi-sheet'
  });
  fs.writeFileSync(file2Path, buf2);
  console.log(`✅ [2/2] File Excel Multi-Tab (Tab per Unit Mobil) Berhasil Dibuat:`);
  console.log(`   📄 ${file2Path}`);

  console.log('\n🎉 ====================================================');
  console.log('   SEMUA FILE EXCEL SELESAI DIBUAT DENGAN SEMPURNA!');
  console.log('   Lokasi file ada di folder: hasil_excel/');
  console.log('====================================================\n');
}

main().catch(err => console.error('FATAL ERROR:', err));
