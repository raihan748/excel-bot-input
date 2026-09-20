const https = require('https');

const DEFAULT_API_KEY = process.env.API_KEY || 'sk-ts-VB0BNV245K445QF7ZCRVCN6B7ADS';
const DEFAULT_MODEL = process.env.AI_MODEL || 'thirty/gpt-6-astra';
const DEFAULT_ENDPOINT = 'https://api.thirtystore.com/v1/chat/completions';

// Daftar model vision cadangan jika model utama sedang sibuk / 502
const FALLBACK_MODELS = [
  'thirty/gpt-6-astra',
  'thirty/deepseek-v4-pro',
  'thirty/deepseek-v4.1-flash'
];

function callSingleModel(targetModel, systemPrompt, mimeType, cleanBase64, apiKey) {
  const payload = JSON.stringify({
    model: targetModel,
    messages: [
      {
        role: 'user',
        content: [
          { type: 'text', text: systemPrompt },
          {
            type: 'image_url',
            image_url: {
              url: `data:${mimeType};base64,${cleanBase64}`
            }
          }
        ]
      }
    ],
    temperature: 0.1,
    max_tokens: 2500
  });

  return new Promise((resolve, reject) => {
    const reqAi = https.request(DEFAULT_ENDPOINT, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
        'User-Agent': 'Mozilla/5.0'
      },
      timeout: 35000
    }, (resp) => {
      let raw = '';
      resp.on('data', chunk => raw += chunk);
      resp.on('end', () => {
        if (resp.statusCode >= 200 && resp.statusCode < 300) {
          try {
            const parsed = JSON.parse(raw);
            const content = parsed.choices?.[0]?.message?.content;
            if (!content) {
              return reject(new Error('Respon AI kosong atau tidak memiliki pilihan teks.'));
            }
            // Jika model mengembalikan pesan bahwa tidak dukung/bisa lihat gambar, picu fallback
            const lowerContent = content.toLowerCase();
            if (
              lowerContent.includes('tidak dukung') ||
              lowerContent.includes('gak dukung') ||
              lowerContent.includes('tidak bisa') ||
              lowerContent.includes('gak bisa') ||
              lowerContent.includes('gagal terkirim') ||
              lowerContent.includes('image omitted') ||
              lowerContent.includes('tidak memiliki kemampuan untuk melihat') ||
              lowerContent.includes('cannot view or analyze images')
            ) {
              return reject(new Error(`Model ${targetModel} menolak/tidak membaca gambar: ${content.slice(0, 100)}`));
            }
            resolve(content);
          } catch (e) {
            reject(new Error('Gagal parse respon API: ' + e.message));
          }
        } else {
          reject(new Error(`API Error HTTP ${resp.statusCode}: ${raw.slice(0, 150)}`));
        }
      });
    });

    reqAi.on('timeout', () => {
      reqAi.destroy();
      reject(new Error(`Timeout 35 detik pada model ${targetModel}`));
    });

    reqAi.on('error', (e) => reject(e));
    reqAi.write(payload);
    reqAi.end();
  });
}

module.exports = async function handler(req, res) {
  // Set CORS headers
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
      image,
      preset = 'auto',
      extraPrompt = '',
      apiKey = DEFAULT_API_KEY,
      model = DEFAULT_MODEL
    } = req.body || {};

    if (!image) {
      return res.status(400).json({ error: 'Gambar tidak boleh kosong. Harap lampirkan gambar base64.' });
    }

    // Bersihkan prefix data URL jika ada
    let cleanBase64 = image;
    let mimeType = 'image/jpeg';
    if (image.startsWith('data:')) {
      const parts = image.split(',');
      const mimeMatch = parts[0].match(/:(.*?);/);
      if (mimeMatch) mimeType = mimeMatch[1];
      cleanBase64 = parts[1] || parts[0];
    }
    // Hapus seluruh spasi dan karakter newline yang bisa merusak payload JSON
    cleanBase64 = cleanBase64.replace(/\s+/g, '');

    // Deteksi otomatis jika gambar berformat PNG
    if (cleanBase64.startsWith('iVBORw0KGgo')) {
      mimeType = 'image/png';
    }

    // Susun instruksi khusus sesuai Preset Dokumen
    let presetInstruction = '';
    switch (preset) {
      case 'nota':
        presetInstruction = `Dokumen ini adalah NOTA / STRUK / KWITANSI TRANSAKSI.
Prioritaskan kolom: ["No", "Nama Barang / Jasa", "Qty", "Harga Satuan", "Total"].
Pastikan angka harga dikonversi menjadi integer murni tanpa tulisan Rp/titik, contoh: 50000.
Jika ada tanggal nota dan nama toko, masukkan ke document_title dan date.`;
        break;
      case 'kas':
        presetInstruction = `Dokumen ini adalah CATATAN KAS / LAPORAN KEUANGAN / ARUS KAS.
Prioritaskan kolom: ["No", "Tanggal", "Keterangan", "Pemasukan", "Pengeluaran", "Saldo"].
Pastikan nominal keuangan berupa angka positif murni.`;
        break;
      case 'stok':
        presetInstruction = `Dokumen ini adalah DAFTAR STOK / INVENTARIS BARANG / GUDANG.
Prioritaskan kolom: ["No", "Kode Barang", "Nama Barang", "Satuan", "Stok Awal", "Masuk", "Keluar", "Stok Akhir"].`;
        break;
      case 'absensi':
        presetInstruction = `Dokumen ini adalah DAFTAR ABSENSI / DAFTAR NAMA / PESERTA.
Prioritaskan kolom: ["No", "Nama Lengkap", "Jabatan/Bagian", "Status Kehadiran", "Keterangan"].`;
        break;
      case 'mobil':
        presetInstruction = `Dokumen ini adalah DATA ARMADA MOBIL / KENDARAAN / RENTAL / SERVIS & BENGKEL MOBIL.
Prioritaskan kolom: ["No", "Mobil / Kendaraan", "Plat Nomor", "Penyewa / Driver / Uraian", "Biaya / Tarif", "Total / Status"].
Pastikan nama atau merk/tipe mobil (misal Avanza, Xenia, Innova, Hiace, Brio, Pajero, dsb) ditulis konsisten agar mudah dikelompokkan.`;
        break;
      default:
        presetInstruction = `Deteksi dan baca seluruh tabel, daftar, kolom, atau tulisan terstruktur yang ada di foto. Buat judul kolom yang paling masuk akal dan relevan sesuai isi dokumen.`;
    }

    const systemPrompt = `Anda adalah asisten AI spesialis ekstraksi dokumen fisik, nota, tabel cetak, dan tulisan tangan ke dalam format tabel spreadsheet Excel.

TUGAS UTAMA:
1. Baca teks dan angka di dalam gambar dengan sangat cermat dan akurat.
2. Identifikasi baris dan kolom tabel. Jika ada tulisan tangan yang agak miring/buram, gunakan konteks kalimat/angka untuk menebak dengan akurasi terbaik.
3. ${presetInstruction}
${extraPrompt ? 'INSTRUKSI KHUSUS PENGGUNA: ' + extraPrompt : ''}

OUTPUT WAJIB:
Kembalikan HANYA string JSON valid (tanpa teks penjelasan pembuka/penutup) dengan struktur:
{
  "document_title": "Judul dokumen / nama toko / judul tabel",
  "document_type": "${preset === 'auto' ? 'Tabel Umum' : preset.toUpperCase()}",
  "date": "YYYY-MM-DD (jika tidak tertera di gambar, gunakan tanggal hari ini)",
  "columns": ["Kolom 1", "Kolom 2", "Kolom 3", "..."],
  "rows": [
    [baris_1_kolom_1, baris_1_kolom_2, baris_1_kolom_3],
    [baris_2_kolom_1, baris_2_kolom_2, baris_2_kolom_3]
  ],
  "total": total_angka_kumulatif_jika_ada_atau_0
}`;

    // Susun antrian model: utamakan model pilihan user, lalu cadangan jika 502
    const modelsQueue = [model, ...FALLBACK_MODELS.filter(m => m !== model)];

    let aiResponse = null;
    let modelSuccess = model;
    let lastError = null;

    for (const currentModel of modelsQueue) {
      try {
        aiResponse = await callSingleModel(currentModel, systemPrompt, mimeType, cleanBase64, apiKey);
        modelSuccess = currentModel;
        break; // Berhasil, keluar dari loop
      } catch (err) {
        console.warn(`[AI WARN] Model ${currentModel} gagal: ${err.message}. Mencoba model cadangan...`);
        lastError = err;
      }
    }

    if (!aiResponse) {
      throw lastError || new Error('Gagal menghubungi AI Vision pada seluruh model.');
    }

    // Ekstrak JSON dari teks balasan AI
    let jsonStr = aiResponse.trim();
    const jsonMatch = aiResponse.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
    if (jsonMatch) {
      jsonStr = jsonMatch[1].trim();
    } else {
      const firstBrace = jsonStr.indexOf('{');
      const lastBrace = jsonStr.lastIndexOf('}');
      if (firstBrace !== -1 && lastBrace !== -1) {
        jsonStr = jsonStr.slice(firstBrace, lastBrace + 1);
      }
    }

    let parsedTableData;
    try {
      parsedTableData = JSON.parse(jsonStr);
    } catch (parseErr) {
      return res.status(422).json({
        error: 'AI memberikan output yang tidak dapat di-parse sebagai JSON valid.',
        raw: aiResponse
      });
    }

    // Validasi struktur minimal
    if (!Array.isArray(parsedTableData.columns)) {
      parsedTableData.columns = ['Kolom 1', 'Kolom 2', 'Kolom 3'];
    }
    if (!Array.isArray(parsedTableData.rows)) {
      parsedTableData.rows = [];
    }

    return res.status(200).json({
      success: true,
      modelUsed: modelSuccess,
      data: parsedTableData
    });

  } catch (error) {
    console.error('Error in /api/process:', error);
    return res.status(500).json({
      error: error.message || 'Terjadi kesalahan sistem saat memproses gambar.'
    });
  }
};
