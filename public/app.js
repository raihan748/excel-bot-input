// State Global Aplikasi
const state = {
  selectedPhotos: [], // Array: { id, file, name, size, base64 }
  selectedPreset: 'nota',
  extraPrompt: '',
  apiKey: localStorage.getItem('excelbot_api_key') || 'sk-ts-VB0BNV245K445QF7ZCRVCN6B7ADS',
  model: localStorage.getItem('excelbot_model') || 'thirty/gpt-6-astra',
  currentTableData: null,
  history: JSON.parse(localStorage.getItem('excelbot_history') || '[]'),
  selectedHistoryIds: new Set(),
  groupColIdx: -1, // -1 = tanpa pengelompokan (standar), >= 0 index kolom pengelompokan
  groupMode: 'single-sheet' // 'single-sheet' atau 'multi-sheet'
};

// DOM Elements
const el = {
  // Tabs
  tabBtnUpload: document.getElementById('tabBtnUpload'),
  tabBtnHistory: document.getElementById('tabBtnHistory'),
  tabBtnMerge: document.getElementById('tabBtnMerge'),
  tabUpload: document.getElementById('tabUpload'),
  tabHistory: document.getElementById('tabHistory'),
  tabMerge: document.getElementById('tabMerge'),
  historyBadge: document.getElementById('historyBadge'),

  // Preset
  presetButtons: document.querySelectorAll('.preset-btn'),

  // Upload Zone
  dropZone: document.getElementById('dropZone'),
  fileInput: document.getElementById('fileInput'),
  btnBrowseFiles: document.getElementById('btnBrowseFiles'),
  selectedPhotosContainer: document.getElementById('selectedPhotosContainer'),
  photosGrid: document.getElementById('photosGrid'),
  photoCountBadge: document.getElementById('photoCountBadge'),
  btnClearAllPhotos: document.getElementById('btnClearAllPhotos'),
  extraPromptInput: document.getElementById('extraPromptInput'),
  btnStartProcessing: document.getElementById('btnStartProcessing'),

  // Progress
  progressCard: document.getElementById('progressCard'),
  progressStatusText: document.getElementById('progressStatusText'),
  progressDetailText: document.getElementById('progressDetailText'),
  progressPercentText: document.getElementById('progressPercentText'),
  progressBarFill: document.getElementById('progressBarFill'),

  // Result Section
  resultSection: document.getElementById('resultSection'),
  resultDocTitle: document.getElementById('resultDocTitle'),
  resultDocType: document.getElementById('resultDocType'),
  resultDocMeta: document.getElementById('resultDocMeta'),
  liveTable: document.getElementById('liveTable'),
  tableHead: document.getElementById('tableHead'),
  tableBody: document.getElementById('tableBody'),
  tableFoot: document.getElementById('tableFoot'),
  btnAddRowBtn: document.getElementById('btnAddRowBtn'),
  btnDownloadCurrentExcel: document.getElementById('btnDownloadCurrentExcel'),
  btnDownloadMultiTabExcel: document.getElementById('btnDownloadMultiTabExcel'),

  // Grouping Toolbar
  groupColumnSelect: document.getElementById('groupColumnSelect'),
  btnQuickGroupMobil: document.getElementById('btnQuickGroupMobil'),
  btnResetGrouping: document.getElementById('btnResetGrouping'),
  groupStatusBadge: document.getElementById('groupStatusBadge'),

  // History Tab
  historyEmptyState: document.getElementById('historyEmptyState'),
  historyItemsTable: document.getElementById('historyItemsTable'),
  historyTableBody: document.getElementById('historyTableBody'),
  selectAllHistoryCheckbox: document.getElementById('selectAllHistoryCheckbox'),
  btnMergeSelectedFromHistory: document.getElementById('btnMergeSelectedFromHistory'),
  btnClearHistory: document.getElementById('btnClearHistory'),

  // Merge Tab
  mergeSelectedCount: document.getElementById('mergeSelectedCount'),
  mergeSelectedList: document.getElementById('mergeSelectedList'),
  mergeTitleInput: document.getElementById('mergeTitleInput'),
  btnExecuteMerge: document.getElementById('btnExecuteMerge'),

  // Settings
  btnOpenSettings: document.getElementById('btnOpenSettings'),
  btnCloseSettings: document.getElementById('btnCloseSettings'),
  settingsModal: document.getElementById('settingsModal'),
  settingApiKey: document.getElementById('settingApiKey'),
  settingModel: document.getElementById('settingModel'),
  btnSaveSettings: document.getElementById('btnSaveSettings'),

  // Header & Guide
  activeModelText: document.getElementById('activeModelText'),
  heroHistoryCount: document.getElementById('heroHistoryCount'),
  btnOpenGuide: document.getElementById('btnOpenGuide'),
  btnCloseGuide: document.getElementById('btnCloseGuide'),
  btnUnderstandGuide: document.getElementById('btnUnderstandGuide'),
  guideModal: document.getElementById('guideModal'),

  // Toast
  toast: document.getElementById('toast'),
  toastMessage: document.getElementById('toastMessage')
};

// ============================================================
// 1. Inisialisasi & Tab Navigation
// ============================================================
function init() {
  setupEventListeners();
  updateHistoryUI();
  updateActiveModelDisplay();
  if (el.settingApiKey) el.settingApiKey.value = state.apiKey;
  if (el.settingModel) el.settingModel.value = state.model;
}

function updateActiveModelDisplay() {
  let name = 'GPT-6 Astra';
  if (state.model.includes('deepseek-v4-pro')) name = 'DeepSeek V4 Pro';
  else if (state.model.includes('deepseek-v4.1-flash')) name = 'DeepSeek V4.1 Flash';
  else if (state.model.includes('claude-sonnet-5')) name = 'Claude Sonnet 5';
  else if (state.model.includes('gpt-6-astra')) name = 'GPT-6 Astra';
  if (el.activeModelText) el.activeModelText.textContent = name;
}

function switchTab(targetTab) {
  [el.tabUpload, el.tabHistory, el.tabMerge].forEach(t => t.classList.add('hidden'));
  [el.tabBtnUpload, el.tabBtnHistory, el.tabBtnMerge].forEach(b => {
    b.classList.remove('tab-active');
    b.classList.add('tab-inactive');
  });

  if (targetTab === 'upload') {
    el.tabUpload.classList.remove('hidden');
    el.tabBtnUpload.classList.add('tab-active');
    el.tabBtnUpload.classList.remove('tab-inactive');
  } else if (targetTab === 'history') {
    el.tabHistory.classList.remove('hidden');
    el.tabBtnHistory.classList.add('tab-active');
    el.tabBtnHistory.classList.remove('tab-inactive');
    updateHistoryUI();
  } else if (targetTab === 'merge') {
    el.tabMerge.classList.remove('hidden');
    el.tabBtnMerge.classList.add('tab-active');
    el.tabBtnMerge.classList.remove('tab-inactive');
    updateMergeTabUI();
  }
}

// ============================================================
// 2. Event Listeners
// ============================================================
function setupEventListeners() {
  // Navigasi Tab
  el.tabBtnUpload.addEventListener('click', () => switchTab('upload'));
  el.tabBtnHistory.addEventListener('click', () => switchTab('history'));
  el.tabBtnMerge.addEventListener('click', () => switchTab('merge'));

  // Preset Buttons
  el.presetButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      el.presetButtons.forEach(b => b.classList.remove('active-preset'));
      btn.classList.add('active-preset');
      state.selectedPreset = btn.getAttribute('data-preset');
    });
  });

  // Dropzone & File Input
  el.btnBrowseFiles.addEventListener('click', (e) => {
    e.stopPropagation();
    el.fileInput.click();
  });
  el.dropZone.addEventListener('click', () => el.fileInput.click());

  el.dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    el.dropZone.classList.add('border-emerald-500', 'bg-emerald-100/50');
  });

  el.dropZone.addEventListener('dragleave', () => {
    el.dropZone.classList.remove('border-emerald-500', 'bg-emerald-100/50');
  });

  el.dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    el.dropZone.classList.remove('border-emerald-500', 'bg-emerald-100/50');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFilesSelected(e.dataTransfer.files);
    }
  });

  el.fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleFilesSelected(e.target.files);
    }
  });

  el.btnClearAllPhotos.addEventListener('click', clearAllSelectedPhotos);

  // Proses Bot Otonom
  el.btnStartProcessing.addEventListener('click', startBatchProcessing);

  // Tabel Aksi
  el.btnAddRowBtn.addEventListener('click', addNewTableRow);

  // Download Standar / Terkelompok (1 Sheet)
  el.btnDownloadCurrentExcel.addEventListener('click', () => {
    if (state.currentTableData) {
      downloadTableAsExcel(state.currentTableData, { groupMode: 'single-sheet' });
    }
  });

  // Download Multi-Tab (Tiap Mobil/Grup Buat Sheet Sendiri)
  if (el.btnDownloadMultiTabExcel) {
    el.btnDownloadMultiTabExcel.addEventListener('click', () => {
      if (state.currentTableData) {
        downloadTableAsExcel(state.currentTableData, { groupMode: 'multi-sheet' });
      }
    });
  }

  // Grouping Events
  if (el.groupColumnSelect) {
    el.groupColumnSelect.addEventListener('change', (e) => {
      state.groupColIdx = parseInt(e.target.value, 10);
      if (state.currentTableData) {
        renderLiveTable(state.currentTableData, false);
      }
    });
  }

  if (el.btnQuickGroupMobil) {
    el.btnQuickGroupMobil.addEventListener('click', () => {
      if (!state.currentTableData || !state.currentTableData.columns) return;
      // Cari kolom mobil / kendaraan
      let targetColIdx = state.currentTableData.columns.findIndex(c => {
        const lower = String(c).toLowerCase();
        return lower.includes('mobil') || lower.includes('kendaraan') || lower.includes('plat') || lower.includes('tipe');
      });
      // Jika tidak ketemu, cari kolom ke-1 (setelah nomor)
      if (targetColIdx === -1 && state.currentTableData.columns.length > 1) {
        targetColIdx = 1;
      }
      if (targetColIdx >= 0) {
        state.groupColIdx = targetColIdx;
        el.groupColumnSelect.value = targetColIdx;
        renderLiveTable(state.currentTableData, false);
        showToast(`Data berhasil dikelompokkan per ${state.currentTableData.columns[targetColIdx]} 🚗`);
      }
    });
  }

  if (el.btnResetGrouping) {
    el.btnResetGrouping.addEventListener('click', () => {
      state.groupColIdx = -1;
      el.groupColumnSelect.value = -1;
      if (state.currentTableData) {
        renderLiveTable(state.currentTableData, false);
        showToast('Pengelompokan di-reset ke tampilan standar');
      }
    });
  }

  // Riwayat File
  el.selectAllHistoryCheckbox.addEventListener('change', (e) => {
    const checked = e.target.checked;
    state.selectedHistoryIds.clear();
    if (checked) {
      state.history.forEach(item => state.selectedHistoryIds.add(item.id));
    }
    updateHistoryUI();
  });

  el.btnMergeSelectedFromHistory.addEventListener('click', () => {
    switchTab('merge');
  });

  el.btnClearHistory.addEventListener('click', () => {
    if (confirm('Yakin ingin menghapus semua riwayat file?')) {
      state.history = [];
      state.selectedHistoryIds.clear();
      localStorage.setItem('excelbot_history', JSON.stringify([]));
      updateHistoryUI();
      showToast('Semua riwayat berhasil dibersihkan');
    }
  });

  // Merge Tab
  el.btnExecuteMerge.addEventListener('click', executeMerge);

  // Settings
  el.btnOpenSettings.addEventListener('click', () => el.settingsModal.classList.remove('hidden'));
  el.btnCloseSettings.addEventListener('click', () => el.settingsModal.classList.add('hidden'));
  el.btnSaveSettings.addEventListener('click', () => {
    state.apiKey = el.settingApiKey.value.trim() || state.apiKey;
    state.model = el.settingModel.value;
    localStorage.setItem('excelbot_api_key', state.apiKey);
    localStorage.setItem('excelbot_model', state.model);
    updateActiveModelDisplay();
    el.settingsModal.classList.add('hidden');
    showToast('Pengaturan API Key & Model berhasil disimpan!');
  });

  // Modal Panduan
  if (el.btnOpenGuide) {
    el.btnOpenGuide.addEventListener('click', () => {
      el.guideModal.classList.remove('hidden');
      lucide.createIcons();
    });
  }
  if (el.btnCloseGuide) el.btnCloseGuide.addEventListener('click', () => el.guideModal.classList.add('hidden'));
  if (el.btnUnderstandGuide) el.btnUnderstandGuide.addEventListener('click', () => el.guideModal.classList.add('hidden'));
}

// ============================================================
// 3. File Handling & Preview
// ============================================================
function handleFilesSelected(fileList) {
  const newFiles = Array.from(fileList).filter(f => f.type.startsWith('image/'));
  if (newFiles.length === 0) {
    showToast('Harap pilih file gambar (JPG, PNG, WEBP, dll)', 'warning');
    return;
  }

  let loaded = 0;
  newFiles.forEach(file => {
    const reader = new FileReader();
    reader.onload = (e) => {
      state.selectedPhotos.push({
        id: 'photo_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9),
        file,
        name: file.name,
        size: formatBytes(file.size),
        base64: e.target.result
      });
      loaded++;
      if (loaded === newFiles.length) {
        renderPhotosGrid();
      }
    };
    reader.readAsDataURL(file);
  });
}

function renderPhotosGrid() {
  if (state.selectedPhotos.length === 0) {
    el.selectedPhotosContainer.classList.add('hidden');
    el.btnStartProcessing.disabled = true;
    return;
  }

  el.selectedPhotosContainer.classList.remove('hidden');
  el.btnStartProcessing.disabled = false;
  el.photoCountBadge.textContent = `${state.selectedPhotos.length} Foto`;

  el.photosGrid.innerHTML = '';
  state.selectedPhotos.forEach((photo, idx) => {
    const item = document.createElement('div');
    item.className = 'group relative bg-slate-100 rounded-xl overflow-hidden border border-slate-200 shadow-sm aspect-square flex flex-col';
    item.innerHTML = `
      <img src="${photo.base64}" alt="${photo.name}" class="w-full h-full object-cover" />
      <div class="absolute inset-0 bg-slate-900/60 opacity-0 group-hover:opacity-100 transition flex flex-col justify-between p-2">
        <button type="button" onclick="removePhoto('${photo.id}')" class="self-end bg-rose-600 hover:bg-rose-700 text-white rounded-lg p-1.5 shadow transition" title="Hapus foto">
          <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
        </button>
        <div class="text-[10px] text-white font-medium truncate bg-slate-950/80 px-1.5 py-0.5 rounded">
          ${idx + 1}. ${photo.name}
        </div>
      </div>
      <div class="absolute top-1.5 left-1.5 bg-slate-900/80 text-white text-[10px] font-bold px-1.5 py-0.5 rounded shadow">
        #${idx + 1}
      </div>
    `;
    el.photosGrid.appendChild(item);
  });

  lucide.createIcons();
}

window.removePhoto = function(id) {
  state.selectedPhotos = state.selectedPhotos.filter(p => p.id !== id);
  renderPhotosGrid();
};

function clearAllSelectedPhotos() {
  state.selectedPhotos = [];
  el.fileInput.value = '';
  renderPhotosGrid();
}

// ============================================================
// 4. Proses Bot Otonom (Batch AI Extractor)
// ============================================================
async function startBatchProcessing() {
  if (state.selectedPhotos.length === 0) return;

  const total = state.selectedPhotos.length;
  el.btnStartProcessing.disabled = true;
  el.dropZone.classList.add('pointer-events-none', 'opacity-50');
  el.progressCard.classList.remove('hidden');
  el.resultSection.classList.add('hidden');

  let combinedColumns = null;
  let allRows = [];
  let totalSum = 0;
  let docTitle = 'REKAP EXCEL HASIL FOTO';
  const extraPrompt = el.extraPromptInput.value.trim();

  // Proses setiap foto berurutan secara otonom
  for (let i = 0; i < total; i++) {
    const current = state.selectedPhotos[i];
    const progressPercent = Math.round(((i + 1) / total) * 100);

    el.progressStatusText.textContent = `Memproses foto: ${current.name}`;
    el.progressDetailText.textContent = `Menganalisis foto ${i + 1} dari ${total}...`;
    el.progressPercentText.textContent = `${progressPercent}%`;
    el.progressBarFill.style.width = `${progressPercent}%`;

    try {
      const res = await fetch('/api/process', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          image: current.base64,
          preset: state.selectedPreset,
          extraPrompt: extraPrompt,
          apiKey: state.apiKey,
          model: state.model
        })
      });

      const json = await res.json();
      if (!res.ok || !json.success) {
        throw new Error(json.error || `Gagal membaca foto ke-${i + 1}`);
      }

      const tableData = json.data;
      if (i === 0) {
        docTitle = tableData.document_title || docTitle;
        combinedColumns = tableData.columns || ['No', 'Item', 'Qty', 'Harga', 'Total'];
      }

      // Gabungkan baris-baris data
      if (Array.isArray(tableData.rows)) {
        tableData.rows.forEach(r => {
          allRows.push(r);
        });
      }

      if (tableData.total && typeof tableData.total === 'number') {
        totalSum += tableData.total;
      }
    } catch (err) {
      console.warn(`Peringatan pada foto ${i + 1}:`, err);
      showToast(`Peringatan foto ${i + 1}: ${err.message}`, 'warning');
    }
  }

  // Jika kolom belum terisi, buat default
  if (!combinedColumns) {
    if (state.selectedPreset === 'mobil') {
      combinedColumns = ['No', 'Mobil / Kendaraan', 'Plat Nomor', 'Driver / Uraian', 'Biaya / Tarif', 'Total'];
    } else {
      combinedColumns = ['No', 'Uraian Barang', 'Jumlah', 'Harga Satuan', 'Total'];
    }
  }

  // Normalisasi nomor urut di kolom pertama jika berupa angka
  allRows.forEach((r, idx) => {
    if (typeof r[0] === 'number') {
      r[0] = idx + 1;
    }
  });

  // Hitung ulang total jika totalSum masih 0
  if (totalSum === 0) {
    totalSum = calculateTotalFromRows(allRows, combinedColumns.length - 1);
  }

  // Simpan data tabel aktif
  state.currentTableData = {
    id: 'doc_' + Date.now(),
    document_title: docTitle,
    document_type: getPresetLabel(state.selectedPreset),
    date: new Date().toISOString().split('T')[0],
    createdAt: new Date().toLocaleString('id-ID'),
    columns: combinedColumns,
    rows: allRows,
    total: totalSum,
    photoCount: total
  };

  // Cek jika preset mobil, otomatis aktifkan kelompokkan mobil
  if (state.selectedPreset === 'mobil') {
    const mobilCol = combinedColumns.findIndex(c => String(c).toLowerCase().includes('mobil'));
    state.groupColIdx = mobilCol >= 0 ? mobilCol : 1;
  } else {
    state.groupColIdx = -1;
  }

  // Simpan ke Histori
  saveToHistory(state.currentTableData);

  // Selesai
  el.progressCard.classList.add('hidden');
  el.dropZone.classList.remove('pointer-events-none', 'opacity-50');
  el.btnStartProcessing.disabled = false;

  renderLiveTable(state.currentTableData);
  showToast(`Selesai! ${total} foto berhasil diubah menjadi tabel Excel.`);
}

// ============================================================
// 5. Live In-Browser Spreadsheet Editor & Grouping Engine
// ============================================================
function renderLiveTable(tableData, resetGroupOptions = true) {
  el.resultSection.classList.remove('hidden');
  el.resultDocTitle.textContent = tableData.document_title;
  el.resultDocType.textContent = tableData.document_type;
  el.resultDocMeta.textContent = `Tanggal: ${tableData.date} | Total ${tableData.rows.length} Baris Data | Diproses dari ${tableData.photoCount || 1} Foto`;

  // Render Header Kolom
  el.tableHead.innerHTML = `
    <tr>
      ${tableData.columns.map(c => `<th>${escapeHtml(c)}</th>`).join('')}
      <th class="w-12 text-center">Aksi</th>
    </tr>
  `;

  // Update Grouping Toolbar Dropdown
  if (resetGroupOptions && el.groupColumnSelect) {
    el.groupColumnSelect.innerHTML = '<option value="-1">-- Tanpa Pengelompokan (Standar) --</option>';
    tableData.columns.forEach((col, idx) => {
      // Abaikan kolom nomor urut sebagai grouping key
      if (idx === 0 && String(col).toLowerCase().includes('no')) return;
      const opt = document.createElement('option');
      opt.value = idx;
      opt.textContent = `Kolom: ${col}`;
      if (idx === state.groupColIdx) opt.selected = true;
      el.groupColumnSelect.appendChild(opt);
    });
  }

  // Tampilkan / Sembunyikan Shortcut Tombol Mobil
  const hasMobilCol = tableData.columns.some(c => {
    const l = String(c).toLowerCase();
    return l.includes('mobil') || l.includes('kendaraan') || l.includes('plat');
  });
  if (el.btnQuickGroupMobil) {
    if (hasMobilCol || state.selectedPreset === 'mobil') {
      el.btnQuickGroupMobil.classList.remove('hidden');
    } else {
      el.btnQuickGroupMobil.classList.add('hidden');
    }
  }

  // Update status badge dan tombol download
  const isGrouped = state.groupColIdx >= 0;
  if (isGrouped) {
    if (el.groupStatusBadge) el.groupStatusBadge.classList.remove('hidden');
    if (el.btnResetGrouping) el.btnResetGrouping.classList.remove('hidden');
    if (el.btnDownloadMultiTabExcel) el.btnDownloadMultiTabExcel.classList.remove('hidden');
    el.btnDownloadCurrentExcel.innerHTML = `
      <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
      <span>Unduh Excel Terkelompok (.xlsx)</span>
    `;
  } else {
    if (el.groupStatusBadge) el.groupStatusBadge.classList.add('hidden');
    if (el.btnResetGrouping) el.btnResetGrouping.classList.add('hidden');
    if (el.btnDownloadMultiTabExcel) el.btnDownloadMultiTabExcel.classList.add('hidden');
    el.btnDownloadCurrentExcel.innerHTML = `
      <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
      <span>Unduh File Excel (.xlsx)</span>
    `;
  }

  // Render Body & Footer
  renderTableRows(tableData);
  renderTableFooter(tableData);

  lucide.createIcons();
}

function renderTableRows(tableData) {
  el.tableBody.innerHTML = '';
  const isGrouped = state.groupColIdx >= 0 && state.groupColIdx < tableData.columns.length;

  if (!isGrouped) {
    // Mode Standar (Tanpa Grup)
    tableData.rows.forEach((row, rIdx) => {
      el.tableBody.appendChild(createTableRowElement(row, rIdx, tableData.columns));
    });
  } else {
    // Mode Terkelompok (Grouped by Column, e.g. Mobil)
    const groupsMap = new Map();
    tableData.rows.forEach((row, rIdx) => {
      const rawVal = row[state.groupColIdx];
      const key = (rawVal !== null && rawVal !== undefined && String(rawVal).trim() !== '')
        ? String(rawVal).trim()
        : 'Lainnya / Tanpa Kategori';
      if (!groupsMap.has(key)) {
        groupsMap.set(key, []);
      }
      groupsMap.get(key).push({ row, originalIndex: rIdx });
    });

    const lastColIdx = tableData.columns.length - 1;

    for (const [groupName, items] of groupsMap.entries()) {
      // 1. Group Header Row
      const headerTr = document.createElement('tr');
      headerTr.className = 'group-header-row';
      headerTr.innerHTML = `
        <td colspan="${tableData.columns.length + 1}">
          <div class="flex items-center justify-between">
            <div class="flex items-center space-x-2">
              <span class="text-base">🚗</span>
              <span class="font-bold text-slate-900 tracking-wide">KELOMPOK [ ${escapeHtml(groupName.toUpperCase())} ]</span>
            </div>
            <span class="text-xs bg-slate-200 text-slate-700 px-2 py-0.5 rounded-full font-bold">
              ${items.length} Data
            </span>
          </div>
        </td>
      `;
      el.tableBody.appendChild(headerTr);

      // 2. Data Rows in Group
      let groupSubtotal = 0;
      items.forEach(({ row, originalIndex }) => {
        el.tableBody.appendChild(createTableRowElement(row, originalIndex, tableData.columns));
        
        // Hitung subtotal dari kolom nominal/terakhir
        const val = row[lastColIdx];
        const num = typeof val === 'number' ? val : Number(String(val || '').replace(/[^0-9\-]/g, ''));
        if (!isNaN(num)) groupSubtotal += num;
      });

      // 3. Group Subtotal Row
      const subTr = document.createElement('tr');
      subTr.className = 'group-subtotal-row';
      subTr.innerHTML = `
        <td colspan="${Math.max(tableData.columns.length - 1, 1)}" class="text-right text-xs uppercase tracking-wider font-bold">
          SUBTOTAL ${escapeHtml(groupName.toUpperCase())}:
        </td>
        <td class="text-right font-black text-blue-800 text-sm">
          ${formatRupiah(groupSubtotal)}
        </td>
        <td class="bg-blue-50/50"></td>
      `;
      el.tableBody.appendChild(subTr);
    }
  }

  // Attach event listener untuk inline editing
  attachCellEditListeners();
}

function createTableRowElement(row, rIdx, columns) {
  const tr = document.createElement('tr');
  tr.className = 'hover:bg-slate-50 transition';

  let rowHtml = '';
  columns.forEach((col, cIdx) => {
    let cellVal = row[cIdx] !== undefined ? row[cIdx] : '';
    let displayVal = cellVal;

    // Format tampilan uang jika angka besar
    if (typeof cellVal === 'number' && cellVal >= 1000) {
      displayVal = formatRupiah(cellVal);
    }

    rowHtml += `
      <td
        contenteditable="true"
        data-row="${rIdx}"
        data-col="${cIdx}"
        class="editable-cell text-slate-700"
      >${escapeHtml(String(displayVal))}</td>
    `;
  });

  // Tombol hapus baris
  rowHtml += `
    <td class="text-center p-1">
      <button type="button" onclick="deleteTableRow(${rIdx})" class="text-slate-400 hover:text-rose-600 p-1 rounded hover:bg-rose-50 transition" title="Hapus baris ini">
        <svg class="w-4 h-4 mx-auto" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
      </button>
    </td>
  `;

  tr.innerHTML = rowHtml;
  return tr;
}

function attachCellEditListeners() {
  el.tableBody.querySelectorAll('.editable-cell').forEach(cell => {
    cell.addEventListener('blur', (e) => {
      const r = parseInt(e.target.getAttribute('data-row'));
      const c = parseInt(e.target.getAttribute('data-col'));
      let text = e.target.innerText.trim();

      // Coba ubah kembali ke angka jika format rupiah/nomor
      const cleanNum = text.replace(/[^0-9\-]/g, '');
      if (cleanNum && !isNaN(Number(cleanNum)) && /^[Rp\s\d.,\-]+$/.test(text)) {
        state.currentTableData.rows[r][c] = Number(cleanNum);
        e.target.innerText = Number(cleanNum) >= 1000 ? formatRupiah(Number(cleanNum)) : cleanNum;
      } else {
        state.currentTableData.rows[r][c] = text;
      }

      // Update total otomatis
      updateCurrentTableTotal();
    });
  });
}

function renderTableFooter(tableData) {
  const colCount = tableData.columns.length;
  el.tableFoot.innerHTML = `
    <tr>
      <td colspan="${Math.max(colCount - 1, 1)}" class="text-right p-3 font-bold text-slate-800 bg-slate-100">
        TOTAL KESELURUHAN (GRAND TOTAL):
      </td>
      <td class="p-3 text-right font-black text-emerald-700 bg-emerald-50 text-base">
        ${formatRupiah(tableData.total || 0)}
      </td>
      <td class="bg-slate-100"></td>
    </tr>
  `;
}

function updateCurrentTableTotal() {
  if (!state.currentTableData) return;
  const lastColIdx = state.currentTableData.columns.length - 1;
  const newTotal = calculateTotalFromRows(state.currentTableData.rows, lastColIdx);
  state.currentTableData.total = newTotal;
  renderTableFooter(state.currentTableData);
}

function addNewTableRow() {
  if (!state.currentTableData) return;
  const newRow = state.currentTableData.columns.map((_, idx) => {
    if (idx === 0) return state.currentTableData.rows.length + 1;
    return '';
  });
  state.currentTableData.rows.push(newRow);
  renderLiveTable(state.currentTableData, false);
  updateCurrentTableTotal();
}

window.deleteTableRow = function(rIdx) {
  if (!state.currentTableData) return;
  state.currentTableData.rows.splice(rIdx, 1);
  state.currentTableData.rows.forEach((r, idx) => {
    if (typeof r[0] === 'number') r[0] = idx + 1;
  });
  renderLiveTable(state.currentTableData, false);
  updateCurrentTableTotal();
};

// ============================================================
// 6. Download Excel (.xlsx) Anti-Gagal (Backend + Client-Side Fallback)
// ============================================================
async function downloadTableAsExcel(tableData, options = {}) {
  const isGrouped = state.groupColIdx >= 0;
  const groupMode = options.groupMode || 'single-sheet';
  
  // Bersihkan nama file agar aman dari karakter terlarang di Windows
  const rawTitle = tableData.document_title || 'Rekap';
  const cleanTitle = rawTitle.replace(/[/\\?%*:|"<>]/g, '_').trim();
  const dateStr = (tableData.date || 'data').replace(/[/\\?%*:|"<>]/g, '_').trim();
  const modeSuffix = groupMode === 'multi-sheet' ? '_MultiTab_PerMobil' : (isGrouped ? '_Terkelompok' : '');
  const safeFilename = `${cleanTitle}_${dateStr}${modeSuffix}.xlsx`;

  showToast('Sedang membuat file Excel (.xlsx)...');

  // Coba 1: Request ke Backend /api/export-excel
  try {
    const res = await fetch('/api/export-excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        tableData: tableData,
        filename: safeFilename,
        isGrouped: isGrouped,
        groupColIdx: state.groupColIdx,
        groupMode: groupMode
      })
    });

    if (res.ok) {
      const blob = await res.blob();
      triggerDownload(blob, safeFilename);
      showToast('File Excel berhasil diunduh ke komputer!');
      return;
    }
  } catch (backendErr) {
    console.warn('[EXCEL] Backend export gagal, beralih ke generator browser...', backendErr);
  }

  // Coba 2: Fallback Generator Langsung di Browser (Client-Side ExcelJS)
  try {
    showToast('Mengunduh langsung via generator internal browser...', 'info');
    await exportExcelClientSide(tableData, safeFilename, isGrouped, state.groupColIdx, groupMode);
    showToast('File Excel berhasil diunduh ke komputer!');
  } catch (clientErr) {
    console.error('[EXCEL FATAL]', clientErr);
    showToast('Gagal mengunduh file Excel: ' + clientErr.message, 'error');
  }
}

/**
 * Generator Excel mandiri langsung di browser (Client-Side Fallback jika server terputus)
 */
async function exportExcelClientSide(tableData, filename, isGrouped, groupColIdx, groupMode) {
  if (!window.ExcelJS) {
    throw new Error('Library ExcelJS belum termuat. Periksa koneksi internet Anda.');
  }

  const workbook = new window.ExcelJS.Workbook();
  workbook.creator = 'ExcelBot AI Browser';
  workbook.created = new Date();

  const columns = Array.isArray(tableData.columns) ? tableData.columns : [];
  const rows = Array.isArray(tableData.rows) ? tableData.rows : [];

  if (isGrouped && groupMode === 'multi-sheet') {
    // Mode Multi-Tab: Tiap mobil dibuatkan Sheet tersendiri
    const groupsMap = new Map();
    rows.forEach(r => {
      const rawVal = r[groupColIdx];
      const key = (rawVal !== null && rawVal !== undefined && String(rawVal).trim() !== '') ? String(rawVal).trim() : 'Lainnya';
      if (!groupsMap.has(key)) groupsMap.set(key, []);
      groupsMap.get(key).push(r);
    });

    // Tab Ringkasan
    const summaryWs = workbook.addWorksheet('RINGKASAN MOBIL');
    summaryWs.addRow(['No', 'Nama Kelompok / Mobil', 'Jumlah Data', 'Total Biaya / Saldo']);
    let no = 1;
    for (const [gName, gRows] of groupsMap.entries()) {
      let gSum = 0;
      gRows.forEach(r => {
        const val = r[r.length - 1];
        const num = typeof val === 'number' ? val : Number(String(val || '').replace(/[^0-9\-]/g, ''));
        if (!isNaN(num)) gSum += num;
      });
      summaryWs.addRow([no++, gName, `${gRows.length} data`, gSum]);
    }

    // Tab Masing-Masing Mobil
    let tabCount = 1;
    for (const [gName, gRows] of groupsMap.entries()) {
      const cleanTab = gName.replace(/[\\/*?:[\]]/g, '').slice(0, 25);
      const ws = workbook.addWorksheet(`${tabCount++}. ${cleanTab}`.slice(0, 31));
      ws.addRow(columns);
      gRows.forEach(r => ws.addRow(r));
    }
  } else {
    // Mode Standar / Single-Sheet
    const ws = workbook.addWorksheet('Data Rekap');
    ws.addRow(columns);
    rows.forEach(r => ws.addRow(r));
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  triggerDownload(blob, filename);
}

function triggerDownload(blob, filename) {
  const downloadUrl = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.style.display = 'none';
  a.href = downloadUrl;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    window.URL.revokeObjectURL(downloadUrl);
  }, 1500);
}

// ============================================================
// 7. Riwayat & Histori File (localStorage)
// ============================================================
function saveToHistory(tableData) {
  state.history.unshift(tableData);
  if (state.history.length > 50) {
    state.history = state.history.slice(0, 50);
  }
  localStorage.setItem('excelbot_history', JSON.stringify(state.history));
  updateHistoryUI();
}

function updateHistoryUI() {
  const count = state.history.length;

  if (el.heroHistoryCount) {
    el.heroHistoryCount.textContent = `${count} File Tersimpan`;
  }

  if (count > 0) {
    el.historyBadge.classList.remove('hidden');
    el.historyBadge.textContent = count;
  } else {
    el.historyBadge.classList.add('hidden');
  }

  if (count === 0) {
    el.historyEmptyState.classList.remove('hidden');
    el.historyItemsTable.classList.add('hidden');
    el.btnMergeSelectedFromHistory.disabled = true;
    el.btnMergeSelectedFromHistory.textContent = 'Gabungkan File Terpilih (0)';
    return;
  }

  el.historyEmptyState.classList.add('hidden');
  el.historyItemsTable.classList.remove('hidden');

  const selectedCount = state.selectedHistoryIds.size;
  el.btnMergeSelectedFromHistory.disabled = selectedCount < 2;
  el.btnMergeSelectedFromHistory.innerHTML = `
    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7h12m0 0l-4-4m4 4l-4 4m0 6H4m0 0l4 4m-4-4l4-4"></path></svg>
    <span>Gabungkan File Terpilih (${selectedCount})</span>
  `;

  el.historyTableBody.innerHTML = '';
  state.history.forEach(item => {
    const isChecked = state.selectedHistoryIds.has(item.id);
    const tr = document.createElement('tr');
    tr.className = `hover:bg-slate-50 transition ${isChecked ? 'bg-blue-50/50' : ''}`;
    tr.innerHTML = `
      <td class="p-3.5 text-center">
        <input
          type="checkbox"
          ${isChecked ? 'checked' : ''}
          onchange="toggleSelectHistory('${item.id}', this.checked)"
          class="rounded text-blue-600 cursor-pointer"
        />
      </td>
      <td class="p-3.5 text-xs text-slate-500 font-mono">${item.createdAt || item.date}</td>
      <td class="p-3.5 font-bold text-slate-800 text-sm">${escapeHtml(item.document_title)}</td>
      <td class="p-3.5 text-xs text-slate-600 font-medium">
        <span class="bg-slate-100 text-slate-700 px-2 py-1 rounded-md border border-slate-200">
          ${escapeHtml(item.document_type || 'Tabel')}
        </span>
      </td>
      <td class="p-3.5 text-center text-xs font-semibold text-slate-600">${item.rows ? item.rows.length : 0} baris</td>
      <td class="p-3.5 text-right font-bold text-emerald-700 text-sm">${formatRupiah(item.total || 0)}</td>
      <td class="p-3.5 text-center">
        <div class="flex items-center justify-center space-x-1.5">
          <button type="button" onclick="loadHistoryItem('${item.id}')" class="bg-slate-100 hover:bg-blue-100 hover:text-blue-700 text-slate-700 p-1.5 rounded-lg text-xs font-medium transition" title="Buka & Edit di Tabel">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"></path><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"></path></svg>
          </button>
          <button type="button" onclick="downloadHistoryItem('${item.id}')" class="bg-emerald-50 hover:bg-emerald-600 hover:text-white text-emerald-700 p-1.5 rounded-lg text-xs font-medium transition" title="Unduh Excel .xlsx">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
          </button>
          <button type="button" onclick="deleteHistoryItem('${item.id}')" class="hover:bg-rose-100 text-slate-400 hover:text-rose-600 p-1.5 rounded-lg text-xs transition" title="Hapus dari riwayat">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
          </button>
        </div>
      </td>
    `;
    el.historyTableBody.appendChild(tr);
  });
}

window.toggleSelectHistory = function(id, isChecked) {
  if (isChecked) {
    state.selectedHistoryIds.add(id);
  } else {
    state.selectedHistoryIds.delete(id);
  }
  updateHistoryUI();
};

window.loadHistoryItem = function(id) {
  const found = state.history.find(h => h.id === id);
  if (found) {
    state.currentTableData = JSON.parse(JSON.stringify(found));
    state.groupColIdx = -1;
    switchTab('upload');
    renderLiveTable(state.currentTableData);
    showToast(`Dokumen "${found.document_title}" dimuat ke tabel`);
  }
};

window.downloadHistoryItem = function(id) {
  const found = state.history.find(h => h.id === id);
  if (found) {
    downloadTableAsExcel(found);
  }
};

window.deleteHistoryItem = function(id) {
  state.history = state.history.filter(h => h.id !== id);
  state.selectedHistoryIds.delete(id);
  localStorage.setItem('excelbot_history', JSON.stringify(state.history));
  updateHistoryUI();
  showToast('File berhasil dihapus dari riwayat');
};

// ============================================================
// 8. Merge Files Feature
// ============================================================
function updateMergeTabUI() {
  const selectedDatasets = state.history.filter(h => state.selectedHistoryIds.has(h.id));
  const count = selectedDatasets.length;

  el.mergeSelectedCount.textContent = `${count} File Dipilih`;
  el.btnExecuteMerge.disabled = count < 2;

  if (count === 0) {
    el.mergeSelectedList.innerHTML = `
      <div class="text-xs text-slate-400 italic py-8 text-center">
        Belum ada file yang dipilih.<br/>
        Buka tab <strong>"2. Riwayat & Histori File"</strong> lalu centang 2 file atau lebih untuk digabung.
      </div>
    `;
    return;
  }

  el.mergeSelectedList.innerHTML = '';
  selectedDatasets.forEach((item, idx) => {
    const d = document.createElement('div');
    d.className = 'bg-white p-3 rounded-xl border border-slate-200 shadow-sm flex items-center justify-between text-xs';
    d.innerHTML = `
      <div class="flex items-center space-x-2">
        <span class="w-5 h-5 rounded-full bg-blue-100 text-blue-700 font-bold flex items-center justify-center text-[10px]">
          ${idx + 1}
        </span>
        <div>
          <div class="font-bold text-slate-800">${escapeHtml(item.document_title)}</div>
          <div class="text-[10px] text-slate-400">${item.rows.length} baris | ${formatRupiah(item.total || 0)}</div>
        </div>
      </div>
      <button type="button" onclick="unselectMergeItem('${item.id}')" class="text-slate-400 hover:text-rose-600">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
      </button>
    `;
    el.mergeSelectedList.appendChild(d);
  });
}

window.unselectMergeItem = function(id) {
  state.selectedHistoryIds.delete(id);
  updateHistoryUI();
  updateMergeTabUI();
};

async function executeMerge() {
  const selectedDatasets = state.history.filter(h => state.selectedHistoryIds.has(h.id));
  if (selectedDatasets.length < 2) {
    showToast('Harap pilih minimal 2 file dari riwayat untuk digabung.', 'warning');
    return;
  }

  const mergeMode = document.querySelector('input[name="mergeMode"]:checked').value;
  const masterTitle = el.mergeTitleInput.value.trim() || 'MASTER REKAP GABUNGAN EXCEL';
  const cleanFilename = `${masterTitle.replace(/[/\\?%*:|"<>]/g, '_')}.xlsx`;

  try {
    showToast('Sedang menggabungkan file Excel master...');
    el.btnExecuteMerge.disabled = true;

    const res = await fetch('/api/merge-excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        datasets: selectedDatasets,
        mode: mergeMode,
        title: masterTitle,
        filename: cleanFilename
      })
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || 'Gagal menggabungkan file');
    }

    const blob = await res.blob();
    triggerDownload(blob, cleanFilename);

    showToast('File Master Gabungan berhasil diunduh!');
  } catch (err) {
    showToast(err.message, 'error');
  } finally {
    el.btnExecuteMerge.disabled = false;
  }
}

// ============================================================
// 9. Helpers & Utilities
// ============================================================
function calculateTotalFromRows(rows, totalColIdx) {
  let sum = 0;
  rows.forEach(r => {
    let val = r[totalColIdx];
    if (typeof val === 'string') {
      const clean = val.replace(/[^0-9\-]/g, '');
      if (clean && !isNaN(Number(clean))) val = Number(clean);
    }
    if (typeof val === 'number' && !isNaN(val)) {
      sum += val;
    }
  });
  return sum;
}

function formatRupiah(num) {
  if (typeof num !== 'number' || isNaN(num)) return 'Rp 0';
  return 'Rp ' + num.toLocaleString('id-ID');
}

function formatBytes(bytes) {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

function getPresetLabel(preset) {
  switch (preset) {
    case 'nota': return 'Nota Belanja';
    case 'mobil': return 'Data Mobil & Armada';
    case 'kas': return 'Buku Kas';
    case 'stok': return 'Stok Barang';
    case 'absensi': return 'Absensi';
    default: return 'Tabel Umum';
  }
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.innerText = text;
  return div.innerHTML;
}

function showToast(message, type = 'success') {
  el.toastMessage.textContent = message;
  el.toast.classList.remove('hidden');

  setTimeout(() => {
    el.toast.classList.add('hidden');
  }, 4000);
}

// Jalankan inisialisasi saat window dimuat
window.addEventListener('DOMContentLoaded', init);
