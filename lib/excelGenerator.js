const ExcelJS = require('exceljs');

/**
 * Membuat worksheet bergaya rapi dan profesional dari satu dataset tabel
 */
function populateWorksheet(worksheet, tableData, sheetName = 'Data') {
  const columns = Array.isArray(tableData.columns) ? tableData.columns : ['No', 'Item', 'Jumlah', 'Harga', 'Total'];
  const rows = Array.isArray(tableData.rows) ? tableData.rows : [];
  const colCount = Math.max(columns.length, 1);
  const lastColLetter = String.fromCharCode(64 + Math.min(colCount, 26));

  // 1. Judul Dokumen (Header Utama)
  worksheet.mergeCells(`A1:${lastColLetter}1`);
  const titleCell = worksheet.getCell('A1');
  titleCell.value = (tableData.document_title || 'REKAP DATA TABEL OTOMATIS').toUpperCase();
  titleCell.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } }; // Biru tua elegan
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  worksheet.getRow(1).height = 32;

  // 2. Sub-header (Tanggal & Tipe Dokumen)
  worksheet.mergeCells(`A2:${lastColLetter}2`);
  const subCell = worksheet.getCell('A2');
  const dateStr = tableData.date || new Date().toISOString().split('T')[0];
  const typeStr = tableData.document_type || 'Ekstraksi AI Vision';
  subCell.value = `Tanggal: ${dateStr}  |  Kategori: ${typeStr}  |  Total Baris: ${rows.length}`;
  subCell.font = { name: 'Segoe UI', size: 10, italic: true, color: { argb: 'FF475569' } };
  subCell.alignment = { horizontal: 'center', vertical: 'middle' };
  worksheet.getRow(2).height = 20;

  // Baris kosong pemisah
  worksheet.getRow(3).height = 10;

  // 3. Header Kolom
  const headerRow = worksheet.getRow(4);
  headerRow.values = columns;
  headerRow.height = 26;
  headerRow.eachCell((cell) => {
    cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } }; // Biru cerah
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      bottom: { style: 'medium', color: { argb: 'FF1D4ED8' } },
      right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
    };
  });

  // Cari index kolom yang mengandung kata 'harga', 'total', 'nominal', 'biaya', 'bayar', 'subtotal'
  const numericCurrencyColIndices = new Set();
  columns.forEach((col, idx) => {
    const colLower = String(col).toLowerCase();
    if (
      colLower.includes('harga') ||
      colLower.includes('total') ||
      colLower.includes('nominal') ||
      colLower.includes('biaya') ||
      colLower.includes('bayar') ||
      colLower.includes('rp') ||
      colLower.includes('subtotal')
    ) {
      numericCurrencyColIndices.add(idx + 1);
    }
  });

  // 4. Data Rows
  let startRow = 5;
  rows.forEach((row, rIdx) => {
    const rowNum = startRow + rIdx;
    const dataRow = worksheet.getRow(rowNum);
    dataRow.values = row;
    dataRow.height = 22;

    dataRow.eachCell((cell, colNumber) => {
      cell.font = { name: 'Segoe UI', size: 10 };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
      };

      // Zebra striping
      if (rIdx % 2 === 1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
      }

      // Bersihkan string angka yang mengandung format "Rp", titik ribuan, atau koma jika AI mengirim teks
      let rawVal = cell.value;
      if (typeof rawVal === 'string' && /^[Rp\s\d.,\-]+$/.test(rawVal.trim())) {
        const cleanVal = rawVal.replace(/[^0-9\-]/g, '');
        if (cleanVal && !isNaN(Number(cleanVal))) {
          rawVal = Number(cleanVal);
          cell.value = rawVal;
        }
      }

      // Formatting angka dan mata uang
      if (typeof cell.value === 'number') {
        if (numericCurrencyColIndices.has(colNumber) || cell.value >= 1000) {
          cell.numFmt = '"Rp" #,##0';
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        } else {
          cell.numFmt = '#,##0';
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
      } else {
        // Kolom pertama biasanya Nomor Urut -> center
        cell.alignment = { horizontal: colNumber === 1 ? 'center' : 'left', vertical: 'middle' };
      }
    });
  });

  // 5. Baris Total Keseluruhan (Jika ada baris data)
  if (rows.length > 0) {
    const totalRowIdx = startRow + rows.length;
    const totalRow = worksheet.getRow(totalRowIdx);
    totalRow.height = 26;

    // Kolom label TOTAL
    if (colCount > 1) {
      const prevColLetter = String.fromCharCode(64 + colCount - 1);
      worksheet.mergeCells(`A${totalRowIdx}:${prevColLetter}${totalRowIdx}`);
    }
    const labelCell = worksheet.getCell(`A${totalRowIdx}`);
    labelCell.value = 'TOTAL KESELURUHAN';
    labelCell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF0F172A' } };
    labelCell.alignment = { horizontal: 'right', vertical: 'middle' };
    labelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };

    // Nilai / Rumus SUM di kolom terakhir
    const formulaCell = worksheet.getCell(`${lastColLetter}${totalRowIdx}`);
    const lastColIndex = colCount;
    if (numericCurrencyColIndices.has(lastColIndex) || true) {
      formulaCell.value = {
        formula: `SUM(${lastColLetter}5:${lastColLetter}${totalRowIdx - 1})`,
        result: tableData.total || 0
      };
      formulaCell.numFmt = '"Rp" #,##0';
    }
    formulaCell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF059669' } }; // Hijau emerald
    formulaCell.alignment = { horizontal: 'right', vertical: 'middle' };
    formulaCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFECFDF5' } };

    // Border ganda di bawah baris total ala akuntansi
    for (let c = 1; c <= colCount; c++) {
      const cell = worksheet.getRow(totalRowIdx).getCell(c);
      cell.border = {
        top: { style: 'medium', color: { argb: 'FF94A3B8' } },
        bottom: { style: 'double', color: { argb: 'FF0F172A' } }
      };
    }
  }

  // 6. Auto-fit Lebar Kolom
  worksheet.columns.forEach((column) => {
    let maxLength = 12;
    column.eachCell({ includeEmpty: false }, (cell) => {
      let cellText = cell.value ? String(cell.value) : '';
      if (typeof cell.value === 'object' && cell.value.result) {
        cellText = String(cell.value.result);
      }
      if (cellText.length > maxLength) {
        maxLength = cellText.length;
      }
    });
    column.width = Math.min(maxLength + 5, 45);
  });
}

/**
 * Membuat Workbook Excel dari satu file tabel
 */
async function generateSingleWorkbook(tableData) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Bot Otonom Excel Vercel';
  workbook.created = new Date();

  const sheetName = (tableData.document_type || 'Data').replace(/[\\/*?:[\]]/g, '').slice(0, 28);
  const worksheet = workbook.addWorksheet(sheetName, {
    views: [{ showGridLines: true }]
  });

  populateWorksheet(worksheet, tableData, sheetName);
  return await workbook.xlsx.writeBuffer();
}

/**
 * Menggabungkan beberapa dataset tabel menjadi satu Workbook Excel
 * mode: 'append' (semua baris digabung jadi 1 lembar) atau 'multi-sheet' (tiap tabel jadi tab sheet sendiri)
 */
async function generateMergedWorkbook(datasets, mode = 'append', customTitle = 'HASIL GABUNGAN (MERGE)') {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Bot Otonom Excel Vercel';
  workbook.created = new Date();

  if (mode === 'multi-sheet') {
    // Tiap dataset dibuatkan Sheet terpisah
    datasets.forEach((data, idx) => {
      const rawName = data.document_title || data.title || `Dokumen ${idx + 1}`;
      const cleanName = rawName.replace(/[\\/*?:[\]]/g, '').slice(0, 25);
      const sheetName = `${idx + 1}. ${cleanName}`.slice(0, 31);
      const ws = workbook.addWorksheet(sheetName, {
        views: [{ showGridLines: true }]
      });
      populateWorksheet(ws, data, sheetName);
    });
  } else {
    // Mode 'append': Gabung semua baris jadi 1 lembar master
    // Ambil kolom dari dataset pertama sebagai patokan
    const baseColumns = datasets[0]?.columns || ['No', 'Nama Barang / Uraian', 'Qty', 'Harga Satuan', 'Total'];
    const mergedRows = [];
    let rowCounter = 1;

    datasets.forEach((data, fileIdx) => {
      const rows = Array.isArray(data.rows) ? data.rows : [];
      rows.forEach((r) => {
        const copyRow = [...r];
        // Jika kolom pertama adalah nomor urut, buat nomor urut berlanjut
        if (typeof copyRow[0] === 'number') {
          copyRow[0] = rowCounter++;
        }
        mergedRows.push(copyRow);
      });
    });

    const combinedData = {
      document_title: customTitle,
      document_type: `Gabungan dari ${datasets.length} File`,
      date: new Date().toISOString().split('T')[0],
      columns: baseColumns,
      rows: mergedRows
    };

    const ws = workbook.addWorksheet('Master Rekap Gabungan', {
      views: [{ showGridLines: true }]
    });
    populateWorksheet(ws, combinedData, 'Master Rekap');
  }

  return await workbook.xlsx.writeBuffer();
}

module.exports = {
  generateSingleWorkbook,
  generateMergedWorkbook,
  populateWorksheet
};
