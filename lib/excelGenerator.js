const ExcelJS = require('exceljs');

/**
 * Helper untuk auto-fit lebar kolom secara aman tanpa crash jika ada cell null
 */
function safeAutoFitColumns(worksheet) {
  worksheet.columns.forEach((column) => {
    let maxLength = 12;
    column.eachCell({ includeEmpty: false }, (cell) => {
      let cellText = '';
      if (cell.value !== null && cell.value !== undefined) {
        if (typeof cell.value === 'object' && cell.value.result !== undefined && cell.value.result !== null) {
          cellText = String(cell.value.result);
        } else if (typeof cell.value === 'object' && cell.value.text !== undefined && cell.value.text !== null) {
          cellText = String(cell.value.text);
        } else if (typeof cell.value !== 'object') {
          cellText = String(cell.value);
        }
      }
      if (cellText.length > maxLength) {
        maxLength = cellText.length;
      }
    });
    column.width = Math.min(maxLength + 5, 45);
  });
}

/**
 * Membuat worksheet standar bergaya rapi dan profesional dari satu dataset tabel
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

  // Cari index kolom yang mengandung kata harga, total, nominal, dsb
  const numericCurrencyColIndices = new Set();
  columns.forEach((col, idx) => {
    const colLower = String(col).toLowerCase();
    if (
      colLower.includes('harga') ||
      colLower.includes('total') ||
      colLower.includes('nominal') ||
      colLower.includes('biaya') ||
      colLower.includes('bayar') ||
      colLower.includes('tarif') ||
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
    formulaCell.value = {
      formula: `SUM(${lastColLetter}5:${lastColLetter}${totalRowIdx - 1})`,
      result: tableData.total || 0
    };
    formulaCell.numFmt = '"Rp" #,##0';
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

  // 6. Auto-fit Lebar Kolom secara aman
  safeAutoFitColumns(worksheet);
}

/**
 * Membuat worksheet TERKELOMPOK (Grouped by Column, misal per Mobil / Kategori)
 * Menampilkan sub-header tiap grup dan subtotal otomatis per grup
 */
function populateGroupedWorksheet(worksheet, tableData, groupColIdx = 1) {
  const columns = Array.isArray(tableData.columns) ? tableData.columns : ['No', 'Mobil / Item', 'Qty', 'Harga', 'Total'];
  const rows = Array.isArray(tableData.rows) ? tableData.rows : [];
  const colCount = Math.max(columns.length, 1);
  const lastColLetter = String.fromCharCode(64 + Math.min(colCount, 26));
  const groupColName = columns[groupColIdx] || 'Kelompok';

  // 1. Judul Dokumen (Header Utama)
  worksheet.mergeCells(`A1:${lastColLetter}1`);
  const titleCell = worksheet.getCell('A1');
  titleCell.value = (tableData.document_title || 'REKAP LAPORAN TERKELOMPOK').toUpperCase();
  titleCell.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } }; // Slate dark
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  worksheet.getRow(1).height = 32;

  // 2. Sub-header (Tanggal & Keterangan Pengelompokan)
  worksheet.mergeCells(`A2:${lastColLetter}2`);
  const subCell = worksheet.getCell('A2');
  const dateStr = tableData.date || new Date().toISOString().split('T')[0];
  subCell.value = `Dikelompokkan Berdasarkan Kolom: [ ${groupColName} ]  |  Tanggal: ${dateStr}  |  Total: ${rows.length} Data`;
  subCell.font = { name: 'Segoe UI', size: 10, italic: true, color: { argb: 'FF475569' } };
  subCell.alignment = { horizontal: 'center', vertical: 'middle' };
  worksheet.getRow(2).height = 20;

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

  // Cari kolom nominal uang
  const numericCurrencyColIndices = new Set();
  columns.forEach((col, idx) => {
    const colLower = String(col).toLowerCase();
    if (
      colLower.includes('harga') ||
      colLower.includes('total') ||
      colLower.includes('nominal') ||
      colLower.includes('biaya') ||
      colLower.includes('bayar') ||
      colLower.includes('tarif') ||
      colLower.includes('rp') ||
      colLower.includes('subtotal')
    ) {
      numericCurrencyColIndices.add(idx + 1);
    }
  });

  // Kelompokkan data baris berdasarkan nilai di groupColIdx
  const groupsMap = new Map();
  rows.forEach((r) => {
    const rawVal = r[groupColIdx];
    const key = (rawVal !== null && rawVal !== undefined && String(rawVal).trim() !== '')
      ? String(rawVal).trim()
      : 'Lainnya / Tanpa Kategori';
    if (!groupsMap.has(key)) {
      groupsMap.set(key, []);
    }
    groupsMap.get(key).push(r);
  });

  let currentRowIdx = 5;
  let grandTotal = 0;

  for (const [groupName, groupRows] of groupsMap.entries()) {
    // A. Baris Header Grup (Misal: Toyota Avanza B 1234 CD)
    worksheet.mergeCells(`A${currentRowIdx}:${lastColLetter}${currentRowIdx}`);
    const gHeaderCell = worksheet.getCell(`A${currentRowIdx}`);
    gHeaderCell.value = `🚗 KELOMPOK [ ${groupName.toUpperCase()} ] — (${groupRows.length} Data / Transaksi)`;
    gHeaderCell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF0F172A' } };
    gHeaderCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } }; // Light Slate
    gHeaderCell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    worksheet.getRow(currentRowIdx).height = 24;
    currentRowIdx++;

    const groupStartRow = currentRowIdx;
    let groupSubtotal = 0;

    // B. Baris Data dalam grup
    groupRows.forEach((r, rIdx) => {
      const dataRow = worksheet.getRow(currentRowIdx);
      dataRow.values = r;
      dataRow.height = 21;

      dataRow.eachCell((cell, colNumber) => {
        cell.font = { name: 'Segoe UI', size: 10 };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
        };

        if (rIdx % 2 === 1) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        }

        // Clean numeric strings
        let rawVal = cell.value;
        if (typeof rawVal === 'string' && /^[Rp\s\d.,\-]+$/.test(rawVal.trim())) {
          const cleanVal = rawVal.replace(/[^0-9\-]/g, '');
          if (cleanVal && !isNaN(Number(cleanVal))) {
            rawVal = Number(cleanVal);
            cell.value = rawVal;
          }
        }

        if (typeof cell.value === 'number') {
          if (numericCurrencyColIndices.has(colNumber) || cell.value >= 1000) {
            cell.numFmt = '"Rp" #,##0';
            cell.alignment = { horizontal: 'right', vertical: 'middle' };
          } else {
            cell.numFmt = '#,##0';
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
          }

          if (colNumber === colCount) {
            groupSubtotal += cell.value;
          }
        } else {
          cell.alignment = { horizontal: colNumber === 1 ? 'center' : 'left', vertical: 'middle' };
        }
      });

      currentRowIdx++;
    });

    const groupEndRow = currentRowIdx - 1;

    // C. Baris Subtotal Grup
    const subtotalRow = worksheet.getRow(currentRowIdx);
    subtotalRow.height = 24;

    if (colCount > 1) {
      const prevColLetter = String.fromCharCode(64 + colCount - 1);
      worksheet.mergeCells(`A${currentRowIdx}:${prevColLetter}${currentRowIdx}`);
    }
    const subLabelCell = worksheet.getCell(`A${currentRowIdx}`);
    subLabelCell.value = `SUBTOTAL ${groupName.toUpperCase()}:`;
    subLabelCell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF1E40AF' } };
    subLabelCell.alignment = { horizontal: 'right', vertical: 'middle' };
    subLabelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF6FF' } }; // Light blue

    const subFormulaCell = worksheet.getCell(`${lastColLetter}${currentRowIdx}`);
    subFormulaCell.value = {
      formula: `SUM(${lastColLetter}${groupStartRow}:${lastColLetter}${groupEndRow})`,
      result: groupSubtotal
    };
    subFormulaCell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF1E40AF' } };
    subFormulaCell.numFmt = '"Rp" #,##0';
    subFormulaCell.alignment = { horizontal: 'right', vertical: 'middle' };
    subFormulaCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF6FF' } };

    for (let c = 1; c <= colCount; c++) {
      worksheet.getRow(currentRowIdx).getCell(c).border = {
        top: { style: 'thin', color: { argb: 'FFBFDBFE' } },
        bottom: { style: 'thin', color: { argb: 'FFBFDBFE' } }
      };
    }

    grandTotal += groupSubtotal;
    currentRowIdx++;

    // Spacer antar grup
    worksheet.getRow(currentRowIdx).height = 8;
    currentRowIdx++;
  }

  // 4. Baris GRAND TOTAL di paling bawah
  const grandTotalRow = worksheet.getRow(currentRowIdx);
  grandTotalRow.height = 28;

  if (colCount > 1) {
    const prevColLetter = String.fromCharCode(64 + colCount - 1);
    worksheet.mergeCells(`A${currentRowIdx}:${prevColLetter}${currentRowIdx}`);
  }
  const grandLabelCell = worksheet.getCell(`A${currentRowIdx}`);
  grandLabelCell.value = 'TOTAL KESELURUHAN (SEMUA KELOMPOK)';
  grandLabelCell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF0F172A' } };
  grandLabelCell.alignment = { horizontal: 'right', vertical: 'middle' };
  grandLabelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };

  const grandFormulaCell = worksheet.getCell(`${lastColLetter}${currentRowIdx}`);
  grandFormulaCell.value = grandTotal;
  grandFormulaCell.font = { name: 'Segoe UI', size: 12, bold: true, color: { argb: 'FF059669' } };
  grandFormulaCell.numFmt = '"Rp" #,##0';
  grandFormulaCell.alignment = { horizontal: 'right', vertical: 'middle' };
  grandFormulaCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFECFDF5' } };

  for (let c = 1; c <= colCount; c++) {
    worksheet.getRow(currentRowIdx).getCell(c).border = {
      top: { style: 'medium', color: { argb: 'FF94A3B8' } },
      bottom: { style: 'double', color: { argb: 'FF0F172A' } }
    };
  }

  // 5. Auto-fit Lebar Kolom secara aman
  safeAutoFitColumns(worksheet);
}

/**
 * Membuat Workbook Excel dari satu file tabel (Mendukung Grouped / Standar)
 */
async function generateSingleWorkbook(tableData, options = {}) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'ExcelBot AI Otonom';
  workbook.created = new Date();

  const isGrouped = Boolean(tableData.isGrouped || options.isGrouped);
  const groupColIdx = typeof options.groupColIdx === 'number' ? options.groupColIdx : (tableData.groupColIdx || 1);
  const groupMode = options.groupMode || tableData.groupMode || 'single-sheet'; // 'single-sheet' atau 'multi-sheet'

  if (isGrouped && groupMode === 'multi-sheet') {
    // Mode Multi-Tab: Tiap Mobil / Kelompok dibuatkan sheet tab tersendiri
    const columns = Array.isArray(tableData.columns) ? tableData.columns : [];
    const rows = Array.isArray(tableData.rows) ? tableData.rows : [];
    const groupsMap = new Map();

    rows.forEach(r => {
      const rawVal = r[groupColIdx];
      const key = (rawVal !== null && rawVal !== undefined && String(rawVal).trim() !== '')
        ? String(rawVal).trim()
        : 'Lainnya';
      if (!groupsMap.has(key)) groupsMap.set(key, []);
      groupsMap.get(key).push(r);
    });

    // 1. Tab Ringkasan / Index
    const summaryWs = workbook.addWorksheet('RINGKASAN KELOMPOK', { views: [{ showGridLines: true }] });
    summaryWs.mergeCells('A1:D1');
    const sTitle = summaryWs.getCell('A1');
    sTitle.value = 'RINGKASAN REKAP PER KELOMPOK / MOBIL';
    sTitle.font = { name: 'Segoe UI', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
    sTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
    sTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    summaryWs.getRow(1).height = 28;

    const sHeader = summaryWs.getRow(3);
    sHeader.values = ['No', 'Nama Kelompok / Mobil', 'Jumlah Data', 'Total Biaya / Nominal'];
    sHeader.height = 24;
    sHeader.eachCell(c => {
      c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    let sRowIdx = 4;
    let sNo = 1;
    let grandSum = 0;

    for (const [gName, gRows] of groupsMap.entries()) {
      let gSum = 0;
      gRows.forEach(r => {
        const lastVal = r[r.length - 1];
        const num = typeof lastVal === 'number' ? lastVal : Number(String(lastVal || '').replace(/[^0-9\-]/g, ''));
        if (!isNaN(num)) gSum += num;
      });
      grandSum += gSum;

      const r = summaryWs.getRow(sRowIdx);
      r.values = [sNo++, gName, `${gRows.length} transaksi`, gSum];
      r.getCell(4).numFmt = '"Rp" #,##0';
      r.getCell(4).alignment = { horizontal: 'right' };
      r.getCell(3).alignment = { horizontal: 'center' };
      sRowIdx++;
    }

    const grandRow = summaryWs.getRow(sRowIdx);
    summaryWs.mergeCells(`A${sRowIdx}:C${sRowIdx}`);
    grandRow.getCell(1).value = 'TOTAL KESELURUHAN:';
    grandRow.getCell(1).font = { bold: true };
    grandRow.getCell(1).alignment = { horizontal: 'right' };
    grandRow.getCell(4).value = grandSum;
    grandRow.getCell(4).font = { bold: true, color: { argb: 'FF059669' } };
    grandRow.getCell(4).numFmt = '"Rp" #,##0';
    safeAutoFitColumns(summaryWs);

    // 2. Tab untuk masing-masing mobil/kelompok
    let tabIdx = 1;
    for (const [gName, gRows] of groupsMap.entries()) {
      const cleanTab = gName.replace(/[\\/*?:[\]]/g, '').slice(0, 25);
      const tabTitle = `${tabIdx++}. ${cleanTab}`.slice(0, 31);
      const ws = workbook.addWorksheet(tabTitle, { views: [{ showGridLines: true }] });
      populateWorksheet(ws, {
        document_title: `${tableData.document_title || 'REKAP'} - ${gName.toUpperCase()}`,
        document_type: `Kelompok: ${gName}`,
        date: tableData.date,
        columns: columns,
        rows: gRows
      }, tabTitle);
    }
  } else if (isGrouped) {
    // Mode Single-Sheet Terkelompok
    const ws = workbook.addWorksheet('Rekap Terkelompok', { views: [{ showGridLines: true }] });
    populateGroupedWorksheet(ws, tableData, groupColIdx);
  } else {
    // Mode Normal Single-Sheet
    const sheetName = (tableData.document_type || 'Data').replace(/[\\/*?:[\]]/g, '').slice(0, 28);
    const ws = workbook.addWorksheet(sheetName, { views: [{ showGridLines: true }] });
    populateWorksheet(ws, tableData, sheetName);
  }

  return await workbook.xlsx.writeBuffer();
}

/**
 * Menggabungkan beberapa dataset tabel menjadi satu Workbook Excel
 */
async function generateMergedWorkbook(datasets, mode = 'append', customTitle = 'HASIL GABUNGAN (MERGE)') {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'ExcelBot AI Otonom';
  workbook.created = new Date();

  if (mode === 'multi-sheet') {
    datasets.forEach((data, idx) => {
      const rawName = data.document_title || data.title || `Dokumen ${idx + 1}`;
      const cleanName = rawName.replace(/[\\/*?:[\]]/g, '').slice(0, 25);
      const sheetName = `${idx + 1}. ${cleanName}`.slice(0, 31);
      const ws = workbook.addWorksheet(sheetName, { views: [{ showGridLines: true }] });
      populateWorksheet(ws, data, sheetName);
    });
  } else {
    const baseColumns = datasets[0]?.columns || ['No', 'Nama Barang / Uraian', 'Qty', 'Harga Satuan', 'Total'];
    const mergedRows = [];
    let rowCounter = 1;

    datasets.forEach((data) => {
      const rows = Array.isArray(data.rows) ? data.rows : [];
      rows.forEach((r) => {
        const copyRow = [...r];
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

    const ws = workbook.addWorksheet('Master Rekap Gabungan', { views: [{ showGridLines: true }] });
    populateWorksheet(ws, combinedData, 'Master Rekap');
  }

  return await workbook.xlsx.writeBuffer();
}

module.exports = {
  generateSingleWorkbook,
  generateMergedWorkbook,
  populateWorksheet,
  populateGroupedWorksheet
};
