/**
 * Rekap Hasil Kerja - Asisten Input Follow Up Shipment
 * Logic & Grouping Engine
 */

(function () {
  'use strict';

  // State Management
  const state = {
    rawRows: [],
    batches: [],
    filteredBatches: [],
    groupingMode: 'kendala_only', // 'kendala_only' (Default SOP) | 'toko_kendala' | 'toko_grouping'
    completedKeys: new Set(),
    activeShift: localStorage.getItem('rekap_active_shift') || '1',
    agentName: localStorage.getItem('rekap_agent_name') || 'SYAIFUL',
    activeFileName: 'Hasil_Kerja.xlsx',
    searchQuery: '',
    statusFilter: 'all', // 'all' | 'pending' | 'completed'
    storeFilter: '',
    simulatorIndex: 0,
    theme: 'light'
  };

  // DOM Elements Cache
  const DOM = {
    html: document.documentElement,
    toastContainer: document.getElementById('toast-container'),
    btnThemeToggle: document.getElementById('btn-theme-toggle'),
    btnShift1: document.getElementById('btn-shift-1'),
    btnShift2: document.getElementById('btn-shift-2'),
    globalAgentName: document.getElementById('global-agent-name'),
    btnOpenSimulator: document.getElementById('btn-open-simulator'),
    btnOpenRecap: document.getElementById('btn-open-recap'),
    btnStartWizard: document.getElementById('btn-start-wizard'),

    // Upload & File
    fileInput: document.getElementById('file-input'),
    btnBrowseFile: document.getElementById('btn-browse-file'),
    btnReloadSample: document.getElementById('btn-reload-sample'),
    activeFilename: document.getElementById('active-filename'),
    activeFileMeta: document.getElementById('active-file-meta'),

    // Paste Spreadsheet Modal
    btnOpenPasteModal: document.getElementById('btn-open-paste-modal'),
    modalPaste: document.getElementById('modal-paste-backdrop'),
    pasteTextarea: document.getElementById('paste-input-textarea'),
    pasteDetectedInfo: document.getElementById('paste-detected-info'),
    btnClearPaste: document.getElementById('btn-clear-paste'),
    btnCancelPaste: document.getElementById('btn-cancel-paste'),
    btnProcessPaste: document.getElementById('btn-process-paste'),
    btnClosePaste: document.getElementById('btn-close-paste'),

    // Stats
    statTotalShipments: document.getElementById('stat-total-shipments'),
    statTotalBatches: document.getElementById('stat-total-batches'),
    statTotalStores: document.getElementById('stat-total-stores'),
    statProgressPercent: document.getElementById('stat-progress-percent'),
    progressBarFill: document.getElementById('progress-bar-fill'),
    statProgressCounter: document.getElementById('stat-progress-counter'),
    btnResetProgress: document.getElementById('btn-reset-progress'),

    // Toolbar & Filters
    groupingModeOptions: document.getElementById('grouping-mode-options'),
    searchInput: document.getElementById('search-input'),
    btnClearSearch: document.getElementById('btn-clear-search'),
    statusFilterTabs: document.getElementById('status-filter-tabs'),
    countAll: document.getElementById('count-all'),
    countPending: document.getElementById('count-pending'),
    countCompleted: document.getElementById('count-completed'),
    filterStore: document.getElementById('filter-store'),
    btnMarkAll: document.getElementById('btn-mark-all'),

    // Batch Grid
    batchGrid: document.getElementById('batch-grid'),
    batchCountInfo: document.getElementById('batch-count-info'),
    emptyState: document.getElementById('empty-state'),
    btnResetFilters: document.getElementById('btn-reset-filters'),

    // Rekap Semua Catatan Kendala
    rekapKendalaSection: document.getElementById('rekap-kendala-section'),
    rekapKendalaText: document.getElementById('rekap-kendala-text'),
    rekapKendalaCount: document.getElementById('rekap-kendala-count'),
    btnCopyAllKendala: document.getElementById('btn-copy-all-kendala'),

    // Floating Progress Dock (Melayang Sederhana & Elegan)
    floatingProgressDock: document.getElementById('floating-progress-dock'),
    floatingProgressFill: document.getElementById('dock-progress-fill'),
    floatingPercentVal: document.getElementById('dock-percent-val'),
    floatingCounterText: document.getElementById('dock-counter-text'),
    dockPercentPill: document.getElementById('dock-percent-pill'),

    // Simulator Modal
    modalSimulator: document.getElementById('modal-simulator-backdrop'),
    simModalTitle: document.getElementById('simulator-modal-title'),
    simBatchIndicator: document.getElementById('simulator-batch-indicator'),
    simTokoBadge: document.getElementById('sim-toko-badge'),
    simShipmentCount: document.getElementById('sim-shipment-count'),
    simBtnPrev: document.getElementById('sim-btn-prev'),
    simBtnNext: document.getElementById('sim-btn-next'),
    simBtnClose: document.getElementById('sim-btn-close'),
    simBtnCopyAll: document.getElementById('sim-btn-copy-all'),

    // Simulator Form Inputs
    simInputShipment: document.getElementById('sim-input-shipment'),
    simFuSudah: document.getElementById('sim-fu-sudah'),
    simFuBelum: document.getElementById('sim-fu-belum'),
    simInputPersonil: document.getElementById('sim-input-personil'),
    simInputPejabat: document.getElementById('sim-input-pejabat'),
    simMotorBagus: document.getElementById('sim-motor-bagus'),
    simMotorRusak: document.getElementById('sim-motor-rusak'),
    btnCopyMotorVal: document.getElementById('btn-copy-motor-val'),
    simSelectHasilFu: document.getElementById('sim-select-hasil-fu'),
    simSelectGrouping: document.getElementById('sim-select-grouping'),
    simSelectAgent: document.getElementById('sim-select-agent'),
    simInputCatatan: document.getElementById('sim-input-catatan'),
    simBtnSubmit: document.getElementById('sim-btn-submit'),
    simBtnCompleteNext: document.getElementById('sim-btn-complete-next'),

    // Recap Modal
    modalRecap: document.getElementById('modal-recap-backdrop'),
    recapTextOutput: document.getElementById('recap-text-output'),
    recapBtnClose: document.getElementById('recap-btn-close'),
    recapBtnDismiss: document.getElementById('recap-btn-dismiss'),
    recapBtnCopy: document.getElementById('recap-btn-copy')
  };

  /* ==========================================================================
     Web Audio Feedback (Pleasant Micro-interaction)
     ========================================================================== */
  function playAudioBeep(type = 'copy') {
    try {
      const AudioContext = window.AudioContext || window.webkitAudioContext;
      if (!AudioContext) return;
      const ctx = new AudioContext();
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.connect(gain);
      gain.connect(ctx.destination);

      if (type === 'copy') {
        osc.frequency.setValueAtTime(650, ctx.currentTime);
        osc.frequency.exponentialRampToValueAtTime(880, ctx.currentTime + 0.08);
        gain.gain.setValueAtTime(0.08, ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.1);
        osc.start(ctx.currentTime);
        osc.stop(ctx.currentTime + 0.1);
      } else if (type === 'success') {
        osc.frequency.setValueAtTime(523.25, ctx.currentTime); // C5
        osc.frequency.setValueAtTime(659.25, ctx.currentTime + 0.08); // E5
        osc.frequency.setValueAtTime(783.99, ctx.currentTime + 0.16); // G5
        gain.gain.setValueAtTime(0.09, ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.28);
        osc.start(ctx.currentTime);
        osc.stop(ctx.currentTime + 0.28);
      }
    } catch (e) {
      // Audio context not allowed or failed silently
    }
  }

  /* ==========================================================================
     Toast Notifications
     ========================================================================== */
  function showToast(message, type = 'success', duration = 2800) {
    if (!DOM.toastContainer) return;
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    let icon = '✅';
    if (type === 'info') icon = 'ℹ️';
    if (type === 'warning') icon = '⚠️';

    toast.innerHTML = `
      <div class="toast-icon">${icon}</div>
      <div class="toast-message">${message}</div>
    `;

    DOM.toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.classList.add('hiding');
      setTimeout(() => toast.remove(), 260);
    }, duration);
  }

  /* ==========================================================================
     Clipboard Copy Helper
     ========================================================================== */
  async function copyText(text, btnElement = null, toastMessage = null) {
    if (text === null || text === undefined) {
      showToast('Tidak ada teks untuk disalin!', 'warning');
      return false;
    }

    // Special handling for empty string (e.g. empty personil toko)
    if (text === '') {
      if (navigator.clipboard) {
        try { await navigator.clipboard.writeText(''); } catch (e) {}
      }
      playAudioBeep('copy');
      if (btnElement) {
        const originalHtml = btnElement.innerHTML;
        btnElement.classList.add('copied');
        btnElement.innerHTML = `<span>(Kosong) ✅</span>`;
        setTimeout(() => {
          btnElement.classList.remove('copied');
          btnElement.innerHTML = originalHtml;
        }, 1800);
      }
      showToast(toastMessage || 'Personil kosong (biarkan kosong pada form)', 'info');
      return true;
    }

    let copied = false;
    if (navigator.clipboard) {
      try {
        await navigator.clipboard.writeText(text);
        copied = true;
      } catch (err) {
        // Fallback below
      }
    }

    if (!copied) {
      try {
        const textarea = document.createElement('textarea');
        textarea.value = text;
        textarea.setAttribute('readonly', '');
        textarea.style.position = 'fixed';
        textarea.style.left = '-9999px';
        textarea.style.top = '0';
        document.body.appendChild(textarea);
        textarea.select();
        copied = document.execCommand('copy');
        document.body.removeChild(textarea);
      } catch (e) {
        console.warn('Clipboard fallback error:', e);
      }
    }

    playAudioBeep('copy');

    if (btnElement) {
      const originalHtml = btnElement.innerHTML;
      btnElement.classList.add('copied');
      btnElement.innerHTML = `
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
          <polyline points="20 6 9 17 4 12"></polyline>
        </svg>
        <span>Tersalin! ✅</span>
      `;
      setTimeout(() => {
        btnElement.classList.remove('copied');
        btnElement.innerHTML = originalHtml;
      }, 1800);
    }

    if (toastMessage) {
      showToast(toastMessage, 'success');
    }

    return true;
  }

  /* ==========================================================================
     Theme & Shift Management
     ========================================================================== */
  function applyTheme(theme) {
    state.theme = theme;
    DOM.html.setAttribute('data-theme', theme);
    localStorage.setItem('rekap_theme', theme);
  }

  function setShift(shift) {
    state.activeShift = shift;
    localStorage.setItem('rekap_active_shift', shift);

    if (DOM.btnShift1) DOM.btnShift1.classList.toggle('active', shift === '1');
    if (DOM.btnShift2) DOM.btnShift2.classList.toggle('active', shift === '2');

    // Load completed keys for this shift
    loadCompletedKeys();
    renderBatches();
    updateStats();
  }

  function loadCompletedKeys() {
    const saved = localStorage.getItem(`rekap_completed_keys_shift_${state.activeShift}`);
    if (saved) {
      try {
        state.completedKeys = new Set(JSON.parse(saved));
      } catch (e) {
        state.completedKeys = new Set();
      }
    } else {
      state.completedKeys = new Set();
    }
  }

  function saveCompletedKeys() {
    localStorage.setItem(
      `rekap_completed_keys_shift_${state.activeShift}`,
      JSON.stringify(Array.from(state.completedKeys))
    );
  }

  /* ==========================================================================
     Grouping Engine
     ========================================================================== */
  function cleanKendalaString(str) {
    if (!str) return '';
    return String(str)
      .replace(/\u00a0/g, ' ')
      .replace(/[\u2010-\u2015]/g, '-')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function normalizeRow(row) {
    // Helper to find value by possible column name matches
    const getVal = (...keys) => {
      for (const k of keys) {
        if (row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '') {
          return String(row[k]).trim();
        }
      }
      return '';
    };

    const shipment = getVal('Shipment', 'No Shipment', 'Nomor Shipment', 'shipment', 'No_Shipment');
    const namaToko = getVal('Nama Toko', 'Toko', 'Nama_Toko', 'nama_toko') || 'TOKO TANPA NAMA';
    const kodeToko = getVal('Kode Toko', 'Kode_Toko', 'kode_toko');
    const cabang = getVal('Cabang', 'DC', 'DC Cabang');
    const region = getVal('Region');
    const namaAC = getVal('Nama AC', 'Nama_AC', 'AC', 'Nama AC Toko');
    const rawKendala = getVal('Kendala', 'Catatan Kendala', 'kendala');
    const rawGrouping = getVal('Grouping', 'Hasil Follow Up', 'grouping');
    const kendala = cleanKendalaString(rawKendala);
    const grouping = cleanKendalaString(rawGrouping);
    const kondisiMotor = getVal('Kondisi Motor', 'Status Motor', 'Motor');
    const rawPersonil = getVal('Jumlah Personil', 'Personil', 'Jumlah Personel');
    // Extract numbers only for personil, or leave empty '' if not provided
    let jumlahPersonil = '';
    if (rawPersonil) {
      const match = String(rawPersonil).match(/\d+/);
      if (match) {
        jumlahPersonil = match[0];
      }
    }
    const namaPejabat = getVal('Nama Pejabat', 'Pejabat Toko', 'Nama Pejabat Toko', 'Pejabat');
    const pic = getVal('PIC', 'Agent', 'Nama Agent') || state.agentName;
    const status = getVal('Status', 'Status Kerja');
    const konsumen = getVal('Nama Konsumen', 'Konsumen');
    const sla = getVal('SLA');

    // Motor mapping: 'Baik'/'baik' -> 'Bagus', 'Rusak'/'1 Rusak' -> 'Rusak'
    let motorStatus = 'Bagus';
    if (kondisiMotor) {
      const lowerMotor = kondisiMotor.toLowerCase();
      if (lowerMotor.includes('rusak')) {
        motorStatus = 'Rusak';
      } else if (lowerMotor.includes('baik') || lowerMotor.includes('bagus')) {
        motorStatus = 'Bagus';
      }
    }

    return {
      shipment,
      namaToko,
      kodeToko,
      cabang,
      region,
      namaAC,
      kendala,
      grouping,
      kondisiMotor,
      motorStatus,
      jumlahPersonil,
      namaPejabat,
      pic,
      status,
      konsumen,
      sla
    };
  }

  function calculateBatches() {
    if (!state.rawRows || state.rawRows.length === 0) {
      state.batches = [];
      state.filteredBatches = [];
      return;
    }

    const groupsMap = new Map();

    state.rawRows.forEach((raw) => {
      const item = normalizeRow(raw);
      if (!item.shipment) return; // Skip rows without shipment ID

      let groupKey = '';
      const cleanK = cleanKendalaString(item.kendala) || cleanKendalaString(item.grouping) || 'Tanpa Kendala';
      const cleanG = cleanKendalaString(item.grouping) || cleanKendalaString(item.kendala) || 'Tanpa Grouping';

      if (state.groupingMode === 'kendala_only') {
        // Strict Kendala-only grouping (combines identical issues across all stores)
        groupKey = `KENDALA__${cleanK.toLowerCase()}`;
      } else if (state.groupingMode === 'toko_kendala') {
        groupKey = `${item.namaToko.toLowerCase()}__SPLIT__${cleanK.toLowerCase()}`;
      } else if (state.groupingMode === 'toko_grouping') {
        groupKey = `${item.namaToko.toLowerCase()}__SPLIT__${cleanG.toLowerCase()}`;
      }

      if (!groupsMap.has(groupKey)) {
        groupsMap.set(groupKey, {
          key: groupKey,
          namaToko: item.namaToko,
          kodeToko: item.kodeToko,
          cabang: item.cabang,
          kendala: cleanK,
          grouping: cleanG,
          motorStatus: item.motorStatus,
          kondisiMotor: item.kondisiMotor,
          jumlahPersonil: item.jumlahPersonil || '',
          namaPejabat: item.namaPejabat,
          namaAC: item.namaAC,
          pic: item.pic || state.agentName,
          items: []
        });
      }

      const grp = groupsMap.get(groupKey);
      grp.items.push(item);

      // Prefer non-empty values for metadata
      if (!grp.kodeToko && item.kodeToko) grp.kodeToko = item.kodeToko;
      if (!grp.cabang && item.cabang) grp.cabang = item.cabang;
      if (!grp.namaPejabat && item.namaPejabat) grp.namaPejabat = item.namaPejabat;
      if (!grp.namaAC && item.namaAC) grp.namaAC = item.namaAC;
      if (item.jumlahPersonil) grp.jumlahPersonil = item.jumlahPersonil;
      if (item.motorStatus === 'Rusak') grp.motorStatus = 'Rusak';
      if (!grp.kendala && item.kendala) grp.kendala = item.kendala;
      if (!grp.grouping && item.grouping) grp.grouping = item.grouping;
    });

    // Form final batch list with comma-separated shipments
    const batchList = [];
    let batchIndex = 1;

    groupsMap.forEach((grp) => {
      // Distinct shipment IDs joined by comma
      const shipmentIds = grp.items.map(it => it.shipment);
      const uniqueShipments = Array.from(new Set(shipmentIds));
      const shipmentMergedString = uniqueShipments.join(',');

      // Store name display
      const uniqueStores = Array.from(new Set(grp.items.map(it => it.namaToko).filter(Boolean)));
      let displayStoreName = grp.namaToko;
      if (state.groupingMode === 'kendala_only') {
        if (uniqueStores.length === 1) {
          displayStoreName = uniqueStores[0];
        } else {
          displayStoreName = uniqueStores.join(', ');
        }
      }

      // Resolve Pejabat Toko (User Rule 2: if empty/missing, use Nama AC from one of the rows in the batch)
      let resolvedPejabat = '';
      let isPejabatFromAC = false;
      for (const it of grp.items) {
        if (it.namaPejabat && it.namaPejabat.trim() && it.namaPejabat.trim() !== '-') {
          resolvedPejabat = it.namaPejabat.trim();
          break;
        }
      }
      if (!resolvedPejabat) {
        for (const it of grp.items) {
          if (it.namaAC && it.namaAC.trim() && it.namaAC.trim() !== '-') {
            resolvedPejabat = it.namaAC.trim();
            isPejabatFromAC = true;
            break;
          }
        }
      }
      if (!resolvedPejabat) resolvedPejabat = '-';

      // Grouping Kendala default: "1. Handle Toko"
      const groupingKendala = '1. Handle Toko';

      // Hasil Follow Up default from grouping or kendala
      let hasilFollowUp = grp.grouping || 'Antrian Pengiriman < 5 Order';
      for (const it of grp.items) {
        if (it.grouping && it.grouping.trim() && it.grouping.trim() !== 'Tanpa Grouping') {
          hasilFollowUp = it.grouping.trim();
          break;
        }
      }

      // Catatan Tambahan default from Kendala
      let catatanTambahan = grp.kendala || '';
      for (const it of grp.items) {
        if (it.kendala && it.kendala.trim() && it.kendala.trim() !== 'Tanpa Kendala') {
          catatanTambahan = it.kendala.trim();
          break;
        }
      }

      batchList.push({
        id: batchIndex,
        key: grp.key,
        namaToko: displayStoreName,
        uniqueStores: uniqueStores,
        storeCount: uniqueStores.length,
        kodeToko: grp.kodeToko,
        cabang: grp.cabang,
        shipments: uniqueShipments,
        shipmentMergedString: shipmentMergedString,
        shipmentCount: uniqueShipments.length,
        jumlahPersonil: grp.jumlahPersonil || '',
        namaPejabat: resolvedPejabat,
        isPejabatFromAC: isPejabatFromAC,
        motorStatus: grp.motorStatus || 'Bagus',
        hasilFollowUp: hasilFollowUp,
        groupingKendala: groupingKendala,
        pic: grp.pic || state.agentName,
        catatanTambahan: catatanTambahan,
        rawItems: grp.items,
        isCompleted: state.completedKeys.has(grp.key)
      });

      batchIndex++;
    });

    state.batches = batchList;
    populateStoreFilterOptions();
    applyFilters();
  }

  function populateStoreFilterOptions() {
    const storeSet = new Set();
    state.batches.forEach(b => {
      if (b.rawItems && b.rawItems.length > 0) {
        b.rawItems.forEach(it => {
          if (it.namaToko) storeSet.add(it.namaToko);
        });
      } else if (b.namaToko) {
        storeSet.add(b.namaToko);
      }
    });
    const stores = Array.from(storeSet).sort();
    const currentVal = DOM.filterStore.value;
    
    let html = '<option value="">Semua Toko</option>';
    stores.forEach(s => {
      html += `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`;
    });
    DOM.filterStore.innerHTML = html;
    DOM.filterStore.value = currentVal;
  }

  /* ==========================================================================
     Filtering & Search
     ========================================================================== */
  function applyFilters() {
    const q = state.searchQuery.toLowerCase().trim();
    const status = state.statusFilter;
    const store = state.storeFilter;

    let pendingCount = 0;
    let completedCount = 0;

    state.batches.forEach(b => {
      b.isCompleted = state.completedKeys.has(b.key);
      if (b.isCompleted) completedCount++;
      else pendingCount++;
    });

    DOM.countAll.textContent = state.batches.length;
    DOM.countPending.textContent = pendingCount;
    DOM.countCompleted.textContent = completedCount;

    state.filteredBatches = state.batches.filter((b) => {
      // Status filter
      if (status === 'pending' && b.isCompleted) return false;
      if (status === 'completed' && !b.isCompleted) return false;

      // Store filter
      if (store) {
        const storeLower = store.toLowerCase();
        const matchesTitle = b.namaToko && b.namaToko.toLowerCase().includes(storeLower);
        const matchesSubItem = b.rawItems && b.rawItems.some(it => it.namaToko && it.namaToko.toLowerCase() === storeLower);
        if (!matchesTitle && !matchesSubItem) return false;
      }

      // Search query filter (matches shipment, store, kendala, pejabat, grouping)
      if (q) {
        const inShipment = b.shipmentMergedString.toLowerCase().includes(q);
        const inToko = b.namaToko.toLowerCase().includes(q);
        const inKendala = b.catatanTambahan.toLowerCase().includes(q);
        const inPejabat = b.namaPejabat.toLowerCase().includes(q);
        const inHasilFu = b.hasilFollowUp.toLowerCase().includes(q);
        if (!inShipment && !inToko && !inKendala && !inPejabat && !inHasilFu) {
          return false;
        }
      }

      return true;
    });

    renderBatches();
    updateStats();
    updateRekapKendala();
  }

  /* ==========================================================================
     UI Rendering
     ========================================================================== */
  function updateRekapKendala() {
    if (!DOM.rekapKendalaText) return;

    // Collect all non-empty catatan tambahan (kendala) from active filtered batches
    const list = [];
    state.filteredBatches.forEach((b) => {
      const val = (b.catatanTambahan || '').trim();
      if (val && val !== '-') {
        list.push(val);
      }
    });

    if (list.length === 0) {
      DOM.rekapKendalaText.textContent = 'Tidak ada catatan tambahan kendala pada batch saat ini.';
      DOM.rekapKendalaText.classList.add('is-empty');
      if (DOM.rekapKendalaCount) DOM.rekapKendalaCount.textContent = '(0 Kendala)';
      return;
    }

    DOM.rekapKendalaText.classList.remove('is-empty');
    const joinedText = list.join(', ');
    DOM.rekapKendalaText.textContent = joinedText;
    if (DOM.rekapKendalaCount) {
      DOM.rekapKendalaCount.textContent = `(${list.length} Catatan Kendala dari ${state.filteredBatches.length} Batch)`;
    }
  }

  function triggerDockPulseAnimation() {
    if (!DOM.floatingProgressDock) return;
    DOM.floatingProgressDock.classList.remove('just-updated');
    if (DOM.floatingPercentVal) DOM.floatingPercentVal.classList.remove('just-updated');
    void DOM.floatingProgressDock.offsetWidth; // force reflow
    DOM.floatingProgressDock.classList.add('just-updated');
    if (DOM.floatingPercentVal) DOM.floatingPercentVal.classList.add('just-updated');
    setTimeout(() => {
      if (DOM.floatingProgressDock) DOM.floatingProgressDock.classList.remove('just-updated');
      if (DOM.floatingPercentVal) DOM.floatingPercentVal.classList.remove('just-updated');
    }, 600);
  }

  function updateStats() {
    // CRITICAL: Synchronize each batch's isCompleted directly from state.completedKeys
    state.batches.forEach(b => {
      b.isCompleted = state.completedKeys.has(b.key);
    });

    const totalShipments = state.rawRows.filter(r => r.Shipment).length;
    const totalBatches = state.batches.length;
    const uniqueStores = new Set(state.rawRows.map(r => r['Nama Toko']).filter(Boolean)).size;

    const completedBatchesCount = state.batches.filter(b => b.isCompleted).length;
    const percent = totalBatches > 0 ? Math.round((completedBatchesCount / totalBatches) * 100) : 0;
    const remainingCount = Math.max(0, totalBatches - completedBatchesCount);

    // Filter count pills
    if (DOM.countAll) DOM.countAll.textContent = totalBatches;
    if (DOM.countPending) DOM.countPending.textContent = remainingCount;
    if (DOM.countCompleted) DOM.countCompleted.textContent = completedBatchesCount;

    // Top Stats Strip
    if (DOM.statTotalShipments) DOM.statTotalShipments.textContent = totalShipments;
    if (DOM.statTotalBatches) DOM.statTotalBatches.textContent = totalBatches;
    if (DOM.statTotalStores) DOM.statTotalStores.textContent = uniqueStores;

    if (DOM.statProgressPercent) DOM.statProgressPercent.textContent = `${percent}%`;
    if (DOM.progressBarFill) DOM.progressBarFill.style.width = `${percent}%`;
    if (DOM.statProgressCounter) DOM.statProgressCounter.textContent = `${completedBatchesCount} dari ${totalBatches} Batch Selesai`;
    if (DOM.batchCountInfo) DOM.batchCountInfo.textContent = `Menampilkan ${state.filteredBatches.length} dari ${totalBatches} batch laporan`;

    // Floating Bottom Progress Dock (Minimalis, Bersih, Sederhana & Beranimasi)
    if (DOM.floatingPercentVal) {
      DOM.floatingPercentVal.textContent = `${percent}%`;
    }
    if (DOM.floatingCounterText) {
      if (percent === 100 && totalBatches > 0) {
        DOM.floatingCounterText.textContent = `🎉 ${completedBatchesCount} / ${totalBatches} Selesai!`;
      } else {
        DOM.floatingCounterText.textContent = `${completedBatchesCount} / ${totalBatches} Selesai`;
      }
    }
    if (DOM.floatingProgressFill) {
      DOM.floatingProgressFill.style.width = `${percent}%`;

      // Smooth color transitions based on progress
      let barColor = '#2563eb';
      let shadowColor = 'rgba(37, 99, 235, 0.45)';

      if (percent === 0) {
        barColor = '#94a3b8';
        shadowColor = 'none';
      } else if (percent <= 25) {
        barColor = '#2563eb'; // Deep Blue
        shadowColor = 'rgba(37, 99, 235, 0.45)';
      } else if (percent <= 50) {
        barColor = '#0284c7'; // Vivid Cyan
        shadowColor = 'rgba(2, 132, 199, 0.45)';
      } else if (percent <= 75) {
        barColor = '#0d9488'; // Ocean Teal
        shadowColor = 'rgba(13, 148, 136, 0.45)';
      } else if (percent < 100) {
        barColor = '#16a34a'; // Vibrant Green
        shadowColor = 'rgba(22, 163, 74, 0.5)';
      } else {
        barColor = '#15803d'; // Victory Green
        shadowColor = 'rgba(21, 128, 61, 0.65)';
      }

      DOM.floatingProgressFill.style.backgroundColor = barColor;
      DOM.floatingProgressFill.style.boxShadow = `0 0 14px ${shadowColor}`;
    }

    if (DOM.floatingProgressDock) {
      if (percent === 100 && totalBatches > 0) {
        DOM.floatingProgressDock.classList.add('is-all-completed');
      } else {
        DOM.floatingProgressDock.classList.remove('is-all-completed');
      }
    }
  }

  function renderBatches() {
    if (state.filteredBatches.length === 0) {
      DOM.batchGrid.innerHTML = '';
      DOM.emptyState.classList.remove('visually-hidden');
      updateRekapKendala();
      return;
    }

    DOM.emptyState.classList.add('visually-hidden');

    let html = '';
    state.filteredBatches.forEach((batch) => {
      const isCompleted = batch.isCompleted;
      const completedClass = isCompleted ? 'is-completed' : '';
      const checkedAttr = isCompleted ? 'checked' : '';
      const storeCodeStr = batch.kodeToko ? ` (${escapeHtml(batch.kodeToko)})` : '';
      const dcStr = batch.cabang ? `<span class="batch-dc-name">📍 ${escapeHtml(batch.cabang)}</span>` : '';

      html += `
        <article class="batch-card ${completedClass}" id="batch-card-${batch.id}" data-batch-id="${batch.id}" data-key="${escapeHtml(batch.key)}">
          
          <!-- Top Row: Store Info & Batch # on Left, Copy All & Selesai Button on Right -->
          <div class="batch-top-row">
            <div class="batch-top-left">
              <span class="batch-number">#${batch.id}</span>
              <div class="batch-toko-name">
                <span>🏪 ${escapeHtml(batch.namaToko)}${storeCodeStr}</span>
                ${dcStr}
              </div>
              <span class="count-pill">📦 ${batch.shipmentCount} Shipment</span>
            </div>

            <div class="batch-top-right">
              <button type="button" class="btn btn-outline btn-xs btn-copy-summary" data-batch-id="${batch.id}" title="Salin seluruh data form sekaligus (Tab-separated)">
                📋 Copy Semua Field
              </button>

              <!-- TOMBOL SELESAI BESAR DI SEBELAH KANAN DENGAN EFEK CENTANG -->
              <label class="btn-toggle-selesai ${isCompleted ? 'is-completed' : ''}" title="Klik untuk menandai batch ini ${isCompleted ? 'belum selesai' : 'sudah selesai'} diinput">
                <input type="checkbox" class="batch-checkbox" data-key="${escapeHtml(batch.key)}" ${checkedAttr}>
                <span class="selesai-checkbox-custom">
                  <svg class="check-svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round">
                    <polyline points="20 6 9 17 4 12"></polyline>
                  </svg>
                </span>
                <span class="selesai-label-text">${isCompleted ? 'Selesai ✅' : 'Tandai Selesai'}</span>
              </label>
            </div>
          </div>

          <!-- Middle Body: Sequential Step Items (1 to 6 Top to Bottom) -->
          <div class="batch-body">
            
            <div class="batch-steps-list">

              <!-- 1. Nomor Shipment -->
              <div class="step-item step-item-shipment">
                <div class="step-item-header">
                  <div class="step-header-left">
                    <span class="step-badge">1</span>
                    <span class="step-label">Nomor Shipment (Dipisah Koma)</span>
                  </div>
                  <button type="button" class="btn-copy-direct" data-copy="${escapeHtml(batch.shipmentMergedString)}" title="Salin nomor shipment dipisah koma">
                    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
                      <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
                      <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
                    </svg>
                    <span>📋 Salin Shipment</span>
                  </button>
                </div>
                <div class="step-shipment-text batch-shipment-text" data-copy="${escapeHtml(batch.shipmentMergedString)}" title="Klik teks untuk langsung menyalin shipment">
                  ${escapeHtml(batch.shipmentMergedString)}
                </div>
              </div>

              <!-- 2. Jumlah Personil Toko -->
              <div class="step-item">
                <div class="step-header-left">
                  <span class="step-badge">2</span>
                  <div class="step-content-box">
                    <span class="step-label">Jumlah Personil Toko</span>
                    <div class="step-val-text ${batch.jumlahPersonil ? 'number-val' : 'empty-val'}">
                      ${batch.jumlahPersonil ? `${escapeHtml(batch.jumlahPersonil)} <span style="font-size:0.75rem; font-weight:500; color:var(--text-muted);">(nomor saja)</span>` : '- (Kosong / Tidak diisi)'}
                    </div>
                  </div>
                </div>
                <button type="button" class="btn-copy-inline" data-copy="${escapeHtml(batch.jumlahPersonil || '')}" title="Salin Jumlah Personil Toko">
                  Salin
                </button>
              </div>

              <!-- 3. Nama Pejabat Toko -->
              <div class="step-item">
                <div class="step-header-left">
                  <span class="step-badge">3</span>
                  <div class="step-content-box">
                    <span class="step-label">Nama Pejabat Toko ${batch.isPejabatFromAC ? '<span style="color:var(--primary); font-weight:600; text-transform:none;">(dari Nama AC)</span>' : ''}</span>
                    <div class="step-val-text ${batch.namaPejabat && batch.namaPejabat !== '-' ? 'bold-val' : 'empty-val'}">
                      ${escapeHtml((batch.namaPejabat && batch.namaPejabat !== '-') ? batch.namaPejabat : '- (Kosong)')}
                    </div>
                  </div>
                </div>
                <button type="button" class="btn-copy-inline" data-copy="${escapeHtml((batch.namaPejabat && batch.namaPejabat !== '-') ? batch.namaPejabat : '')}" title="Salin Nama Pejabat Toko">
                  Salin
                </button>
              </div>

              <!-- 4. Status Equipment Motor? -->
              <div class="step-item">
                <div class="step-header-left">
                  <span class="step-badge">4</span>
                  <div class="step-content-box">
                    <span class="step-label">Status Equipment Motor?</span>
                    <div>
                      <span class="motor-status-pill ${batch.motorStatus === 'Rusak' ? 'is-rusak' : 'is-bagus'}">
                        ${batch.motorStatus === 'Rusak' ? '⚠️ Rusak' : '✅ Bagus'}
                      </span>
                    </div>
                  </div>
                </div>
                <button type="button" class="btn-copy-inline" data-copy="${escapeHtml(batch.motorStatus)}" title="Salin Status Motor (${batch.motorStatus})">
                  Salin
                </button>
              </div>

              <!-- 5. Hasil Follow Up (Grouping) -->
              <div class="step-item step-fu">
                <div class="step-header-left">
                  <span class="step-badge">5</span>
                  <div class="step-content-box">
                    <span class="step-label">Hasil Follow Up (Grouping)</span>
                    <div class="step-val-text fu-val">
                      ${escapeHtml(batch.hasilFollowUp)}
                    </div>
                  </div>
                </div>
                <button type="button" class="btn-copy-inline" data-copy="${escapeHtml(batch.hasilFollowUp)}" title="Salin Hasil Follow Up">
                  Salin
                </button>
              </div>

              <!-- 6. Catatan Tambahan (Kendala) -->
              <div class="step-item step-kendala">
                <div class="step-header-left">
                  <span class="step-badge">6</span>
                  <div class="step-content-box">
                    <span class="step-label">Catatan Tambahan (Kendala)</span>
                    <div class="step-val-text kendala-val">
                      ${escapeHtml(batch.catatanTambahan || '- (Tanpa catatan kendala)')}
                    </div>
                  </div>
                </div>
                <button type="button" class="btn-copy-inline" data-copy="${escapeHtml(batch.catatanTambahan || '')}" title="Salin Catatan Tambahan (Kendala)">
                  Salin
                </button>
              </div>

            </div>

            <!-- Bottom Row Actions -->
            <div class="batch-bottom-row">
              <div class="batch-bottom-left">
                <button type="button" class="btn btn-primary btn-xs btn-open-sim-batch" data-batch-id="${batch.id}">
                  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <rect x="2" y="3" width="20" height="14" rx="2" ry="2"></rect>
                    <line x1="8" y1="21" x2="16" y2="21"></line>
                    <line x1="12" y1="17" x2="12" y2="21"></line>
                  </svg>
                  <span>Form Simulator</span>
                </button>
              </div>
              <div class="batch-bottom-right">
                <span class="chip-counter">Grouping: ${escapeHtml(batch.groupingKendala)} | PIC: ${escapeHtml(batch.pic)}</span>
              </div>
            </div>

          </div>

        </article>
      `;
    });

    DOM.batchGrid.innerHTML = html;
  }

  /* ==========================================================================
     Simulator Modal Logic
     ========================================================================== */
  function openSimulator(batchIndex = 0) {
    if (state.batches.length === 0) {
      showToast('Belum ada data untuk disimulasikan!', 'warning');
      return;
    }

    state.simulatorIndex = Math.max(0, Math.min(batchIndex, state.batches.length - 1));
    populateSimulatorValues();
    DOM.modalSimulator.classList.remove('visually-hidden');
    document.body.style.overflow = 'hidden';
  }

  function closeSimulator() {
    DOM.modalSimulator.classList.add('visually-hidden');
    document.body.style.overflow = '';
  }

  function populateSimulatorValues() {
    const batch = state.batches[state.simulatorIndex];
    if (!batch) return;

    // Header updates
    DOM.simBatchIndicator.textContent = `Batch ${state.simulatorIndex + 1} dari ${state.batches.length}`;
    DOM.simTokoBadge.textContent = `${batch.namaToko}${batch.kodeToko ? ' (' + batch.kodeToko + ')' : ''}`;
    DOM.simShipmentCount.textContent = `${batch.shipmentCount} Shipment Digabung`;

    // Navigation buttons
    DOM.simBtnPrev.disabled = state.simulatorIndex === 0;
    DOM.simBtnNext.disabled = state.simulatorIndex === state.batches.length - 1;

    // Populate form fields matching user's screenshot
    DOM.simInputShipment.value = batch.shipmentMergedString;
    DOM.simFuSudah.checked = true; // Default sudah follow up
    DOM.simFuBelum.checked = false;

    DOM.simInputPersonil.value = batch.jumlahPersonil || '';
    DOM.simInputPejabat.value = (batch.namaPejabat && batch.namaPejabat !== '-') ? batch.namaPejabat : '';

    // Status motor
    if (batch.motorStatus === 'Rusak') {
      DOM.simMotorRusak.checked = true;
      DOM.simMotorBagus.checked = false;
    } else {
      DOM.simMotorBagus.checked = true;
      DOM.simMotorRusak.checked = false;
    }

    // Hasil Follow Up
    // Check if option exists in select, if not add it dynamically
    let found = false;
    for (let i = 0; i < DOM.simSelectHasilFu.options.length; i++) {
      if (DOM.simSelectHasilFu.options[i].value === batch.hasilFollowUp) {
        DOM.simSelectHasilFu.selectedIndex = i;
        found = true;
        break;
      }
    }
    if (!found && batch.hasilFollowUp) {
      const opt = document.createElement('option');
      opt.value = batch.hasilFollowUp;
      opt.textContent = batch.hasilFollowUp;
      DOM.simSelectHasilFu.appendChild(opt);
      DOM.simSelectHasilFu.value = batch.hasilFollowUp;
    }

    DOM.simSelectGrouping.value = batch.groupingKendala || '1. Handle Toko';
    DOM.simSelectAgent.value = batch.pic || state.agentName;
    DOM.simInputCatatan.value = batch.catatanTambahan || '';
  }

  function submitSimulatorForm() {
    const batch = state.batches[state.simulatorIndex];
    if (!batch) return;

    // Mark as completed
    state.completedKeys.add(batch.key);
    saveCompletedKeys();
    batch.isCompleted = true;

    playAudioBeep('success');
    showToast(`Status Batch #${batch.id} (${batch.namaToko}) Berhasil Diupdate! ✅`, 'success');

    // Update main UI in background
    applyFilters();

    // Auto advance to next pending batch if available
    let nextIndex = -1;
    for (let i = state.simulatorIndex + 1; i < state.batches.length; i++) {
      if (!state.completedKeys.has(state.batches[i].key)) {
        nextIndex = i;
        break;
      }
    }
    if (nextIndex === -1) {
      // Check from beginning
      for (let i = 0; i < state.simulatorIndex; i++) {
        if (!state.completedKeys.has(state.batches[i].key)) {
          nextIndex = i;
          break;
        }
      }
    }

    if (nextIndex !== -1) {
      setTimeout(() => {
        state.simulatorIndex = nextIndex;
        populateSimulatorValues();
      }, 500);
    } else {
      showToast('🎉 Semua batch telah selesai diinput!', 'success', 3500);
      setTimeout(() => {
        closeSimulator();
      }, 800);
    }
  }

  /* ==========================================================================
     Recap WhatsApp Generator
     ========================================================================== */
  function openRecapModal() {
    if (state.batches.length === 0) {
      showToast('Tidak ada data batch untuk dirangkum!', 'warning');
      return;
    }

    const today = new Date().toLocaleDateString('id-ID', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });

    const shiftText = state.activeShift === '1' ? 'Shift 1 (Pagi)' : 'Shift 2 (Sore)';
    const completedCount = state.batches.filter(b => b.isCompleted).length;

    let text = `*LAPORAN HASIL KERJA - FOLLOW UP SHIPMENT*\n`;
    text += `📅 Tanggal: ${today}\n`;
    text += `⏰ Shift: ${shiftText}\n`;
    text += `👤 Agent: ${state.agentName}\n`;
    text += `📊 Total: ${state.rawRows.length} Shipment | ${state.batches.length} Batch (${completedCount} Selesai Diinput)\n\n`;
    text += `-----------------------------------------\n`;

    state.batches.forEach((b, idx) => {
      const personilStr = b.jumlahPersonil ? b.jumlahPersonil : '(Kosong)';
      const pejabatStr = (b.namaPejabat && b.namaPejabat !== '-') ? b.namaPejabat : '(Kosong)';
      text += `*${idx + 1}. ${b.namaToko}* (${b.shipmentCount} Shipment)\n`;
      text += `1. Shipment: ${b.shipmentMergedString}\n`;
      text += `2. Personil Toko: ${personilStr}\n`;
      text += `3. Pejabat Toko: ${pejabatStr}\n`;
      text += `4. Status Motor: ${b.motorStatus}\n`;
      text += `5. Hasil Follow Up: ${b.hasilFollowUp}\n`;
      text += `6. Catatan Tambahan: ${b.catatanTambahan || '-'}\n`;
      text += `• Status Input: ${b.isCompleted ? '✅ Selesai' : '⏳ Belum'}\n\n`;
    });

    text += `-----------------------------------------\n`;
    text += `_Laporan otomatis digenerate via Asisten Rekap Hasil Kerja_`;

    DOM.recapTextOutput.value = text;
    DOM.modalRecap.classList.remove('visually-hidden');
    document.body.style.overflow = 'hidden';
  }

  function closeRecapModal() {
    DOM.modalRecap.classList.add('visually-hidden');
    document.body.style.overflow = '';
  }

  /* ==========================================================================
     Excel File Processing (SheetJS)
     ========================================================================== */
  function handleExcelFile(file) {
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

        if (!json || json.length === 0) {
          showToast('File Excel kosong atau format tidak sesuai!', 'warning');
          return;
        }

        state.rawRows = json;
        state.activeFileName = file.name;
        DOM.activeFilename.textContent = file.name;
        DOM.activeFileMeta.textContent = `(${json.length} baris data berhasil dimuat)`;

        calculateBatches();
        showToast(`Berhasil memuat ${json.length} baris dari ${file.name}`, 'success');
      } catch (err) {
        console.error('Error parsing excel:', err);
        showToast('Gagal membaca file Excel. Pastikan file valid .xlsx / .xls!', 'warning');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function loadDefaultSampleData() {
    if (window.DEFAULT_DATA && Array.isArray(window.DEFAULT_DATA)) {
      state.rawRows = window.DEFAULT_DATA;
      state.activeFileName = 'Hasil_Kerja.xlsx';
      if (DOM.activeFilename) DOM.activeFilename.textContent = 'Hasil_Kerja.xlsx';
      if (DOM.activeFileMeta) DOM.activeFileMeta.textContent = `(${window.DEFAULT_DATA.length} baris)`;
      calculateBatches();
      showToast('Data contoh Hasil_Kerja.xlsx berhasil dimuat!', 'success');
    } else {
      showToast('File default tidak ditemukan. Silakan upload file Excel Anda.', 'warning');
    }
  }

  /* ==========================================================================
     Spreadsheet Paste Processing (Google Sheets / Excel)
     ========================================================================== */
  function parseSpreadsheetData(rawText) {
    if (!rawText || !rawText.trim()) return [];

    const lines = rawText.split(/\r?\n/).map(l => l.trimEnd()).filter(l => l.trim().length > 0);
    if (lines.length === 0) return [];

    // Delimiter detection: Tab (\t) is standard clipboard format from Google Sheets & Excel
    const sampleLine = lines[0];
    let delimiter = '\t';
    if (!sampleLine.includes('\t')) {
      if (sampleLine.includes(';') && (sampleLine.match(/;/g) || []).length > (sampleLine.match(/,/g) || []).length) {
        delimiter = ';';
      } else if (sampleLine.includes(',')) {
        delimiter = ',';
      }
    }

    const rows = lines.map(line => {
      if (delimiter === '\t') {
        return line.split('\t').map(c => c.trim().replace(/^"(.*)"$/, '$1'));
      }
      const regex = new RegExp(`(?:^|${delimiter})(?:"([^"]*)"|([^${delimiter}]*))`, 'g');
      const cells = [];
      let match;
      while ((match = regex.exec(line)) !== null) {
        cells.push((match[1] || match[2] || '').trim());
      }
      return cells;
    });

    if (rows.length === 0) return [];

    const firstRow = rows[0];
    
    // Check if first row has actual data (e.g. Shipment code like S-261020-...)
    const containsShipmentData = firstRow.some(cell => /^S-\d{4,}/i.test(cell.trim()));

    const headerKeywords = ['shipment', 'toko', 'kendala', 'pejabat', 'personil', 'motor', 'follow', 'ac', 'cabang', 'dc', 'kode toko', 'nama toko'];
    const hasHeaders = !containsShipmentData && firstRow.some(cell => {
      const lower = String(cell).trim().toLowerCase();
      return headerKeywords.some(kw => lower === kw || lower.includes(kw));
    });

    let headers = [];
    let dataRows = [];

    const defaultCols = [
      "Jam", "No AC", "Nama AC", "Kode Toko", "Nama Toko", "Cabang", "Region", 
      "Shipment", "Nama Konsumen", "SLA", "Nama AM", "No AM", "POD AC", "APV AM", 
      "PIC", "Status", "Kondisi Motor", "Jumlah Personil", "Nama Pejabat", "Kendala", "Grouping", "GLI"
    ];

    if (hasHeaders) {
      headers = firstRow.map((h, i) => h.trim() || `Kolom_${i + 1}`);
      dataRows = rows.slice(1);
    } else {
      // If columns match the default width (18-25 columns), use default column mapping
      if (firstRow.length >= 18 && firstRow.length <= 25) {
        headers = defaultCols.slice(0, firstRow.length);
        dataRows = rows;
      } else {
        // Dynamic heuristic: identify key columns by value patterns
        let shipmentColIdx = -1;
        let tokoColIdx = -1;
        let kendalaColIdx = -1;

        for (let col = 0; col < firstRow.length; col++) {
          const sampleValues = rows.slice(0, 5).map(r => r[col] || '');
          if (sampleValues.some(v => /^S-\d+/i.test(v) || /AG[A-Z0-9]+/i.test(v))) {
            shipmentColIdx = col;
          } else if (sampleValues.some(v => /pembatalan|batal|rusak|jaringan|titip|retur|ganti/i.test(v))) {
            kendalaColIdx = col;
          } else if (sampleValues.some(v => /DC |ALFAMART|INDOMARET|KAPT|TOKO/i.test(v))) {
            tokoColIdx = col;
          }
        }

        headers = firstRow.map((_, i) => {
          if (i === shipmentColIdx) return 'Shipment';
          if (i === tokoColIdx) return 'Nama Toko';
          if (i === kendalaColIdx) return 'Kendala';
          return `Kolom_${i + 1}`;
        });
        dataRows = rows;
      }
    }

    const result = [];
    dataRows.forEach(cells => {
      if (cells.length === 0 || cells.every(c => !c)) return;
      const obj = {};
      headers.forEach((h, idx) => {
        obj[h] = cells[idx] !== undefined ? cells[idx] : '';
      });
      result.push(obj);
    });

    return result;
  }

  function openPasteModal() {
    if (!DOM.modalPaste) return;
    DOM.modalPaste.classList.remove('visually-hidden');
    document.body.style.overflow = 'hidden';
    if (DOM.pasteTextarea) {
      DOM.pasteTextarea.focus();
      updatePasteDetectionInfo();
    }
  }

  function closePasteModal() {
    if (!DOM.modalPaste) return;
    DOM.modalPaste.classList.add('visually-hidden');
    document.body.style.overflow = '';
  }

  function updatePasteDetectionInfo() {
    if (!DOM.pasteTextarea || !DOM.pasteDetectedInfo) return;
    const text = DOM.pasteTextarea.value.trim();
    if (!text) {
      DOM.pasteDetectedInfo.textContent = 'Belum ada data yang ditempel';
      return;
    }

    const parsed = parseSpreadsheetData(text);
    if (parsed.length > 0) {
      DOM.pasteDetectedInfo.textContent = `✅ Terdeteksi ${parsed.length} baris data spreadsheet siap diolah`;
    } else {
      DOM.pasteDetectedInfo.textContent = '⚠️ Format data belum terdeteksi. Pastikan data disalin dari spreadsheet.';
    }
  }

  function processPastedSpreadsheet() {
    const text = DOM.pasteTextarea ? DOM.pasteTextarea.value.trim() : '';
    if (!text) {
      showToast('Silakan tempel (paste) data dari spreadsheet terlebih dahulu!', 'warning');
      if (DOM.pasteTextarea) DOM.pasteTextarea.focus();
      return;
    }

    const parsed = parseSpreadsheetData(text);
    if (!parsed || parsed.length === 0) {
      showToast('Gagal memproses data. Pastikan baris data memiliki nomor shipment!', 'warning');
      return;
    }

    state.rawRows = parsed;
    const sourceName = 'Data Salinan Spreadsheet';
    state.activeFileName = sourceName;
    if (DOM.activeFilename) DOM.activeFilename.textContent = sourceName;
    if (DOM.activeFileMeta) DOM.activeFileMeta.textContent = `(${parsed.length} baris)`;

    state.completedKeys.clear();
    saveCompletedKeys();
    calculateBatches();
    closePasteModal();

    playAudioBeep('success');
    showToast(`Berhasil memuat ${parsed.length} baris data dari spreadsheet!`, 'success');
  }

  /* ==========================================================================
     Export to Excel
     ========================================================================== */
  function exportGroupedExcel() {
    if (state.batches.length === 0) {
      showToast('Tidak ada data batch untuk diekspor!', 'warning');
      return;
    }

    try {
      const exportData = state.batches.map((b, i) => ({
        'No Batch': i + 1,
        'Nama Toko': b.namaToko,
        'Kode Toko': b.kodeToko,
        'Cabang': b.cabang,
        'Jumlah Shipment': b.shipmentCount,
        'Nomor Shipment (Koma)': b.shipmentMergedString,
        'Jumlah Personil': b.jumlahPersonil,
        'Nama Pejabat Toko': b.namaPejabat,
        'Status Motor': b.motorStatus,
        'Hasil Follow Up': b.hasilFollowUp,
        'Grouping Kendala': b.groupingKendala,
        'Agent Follow Up': b.pic,
        'Catatan Tambahan (Kendala)': b.catatanTambahan,
        'Status Input Web': b.isCompleted ? 'Selesai' : 'Belum Diinput'
      }));

      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Hasil_Gabungan_Batch');

      const fileName = `Rekap_Grouped_Shipment_Shift${state.activeShift}.xlsx`;
      XLSX.writeFile(wb, fileName);

      showToast(`File ${fileName} berhasil diunduh!`, 'success');
    } catch (err) {
      console.error('Export error: ', err);
      showToast('Gagal mengekspor file Excel', 'warning');
    }
  }

  /* ==========================================================================
     Helper Utilities
     ========================================================================== */
  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  /* ==========================================================================
     Event Listeners Binding
     ========================================================================== */
  function initEventListeners() {
    // Theme toggle
    DOM.btnThemeToggle.addEventListener('click', () => {
      const newTheme = state.theme === 'dark' ? 'light' : 'dark';
      applyTheme(newTheme);
      showToast(`Tema ${newTheme === 'dark' ? 'Gelap' : 'Terang'} diaktifkan`, 'info');
    });

    // Shift selector (guarded)
    if (DOM.btnShift1) DOM.btnShift1.addEventListener('click', () => setShift('1'));
    if (DOM.btnShift2) DOM.btnShift2.addEventListener('click', () => setShift('2'));

    // Agent name change (guarded)
    if (DOM.globalAgentName) {
      DOM.globalAgentName.value = state.agentName;
      DOM.globalAgentName.addEventListener('change', (e) => {
        state.agentName = e.target.value.trim() || 'SYAIFUL';
        localStorage.setItem('rekap_agent_name', state.agentName);
        calculateBatches();
        showToast(`Nama agent disimpan: ${state.agentName}`, 'info');
      });
    }

    // Grouping mode change (guarded)
    if (DOM.groupingModeOptions) {
      const radioTabs = DOM.groupingModeOptions.querySelectorAll('.radio-tab');
      radioTabs.forEach(tab => {
        tab.addEventListener('click', () => {
          radioTabs.forEach(t => t.classList.remove('active'));
          tab.classList.add('active');
          const input = tab.querySelector('input[type="radio"]');
          if (input) {
            input.checked = true;
            state.groupingMode = input.value;
            calculateBatches();
            showToast(`Mode pengelompokan diubah ke: ${input.value.replace('_', ' ')}`, 'info');
          }
        });
      });
    }

    // Search input
    DOM.searchInput.addEventListener('input', (e) => {
      state.searchQuery = e.target.value;
      DOM.btnClearSearch.classList.toggle('visually-hidden', !state.searchQuery);
      applyFilters();
    });

    DOM.btnClearSearch.addEventListener('click', () => {
      DOM.searchInput.value = '';
      state.searchQuery = '';
      DOM.btnClearSearch.classList.add('visually-hidden');
      applyFilters();
    });

    // Status filter tabs
    DOM.statusFilterTabs.addEventListener('click', (e) => {
      const btn = e.target.closest('.btn-tab');
      if (!btn) return;
      DOM.statusFilterTabs.querySelectorAll('.btn-tab').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.statusFilter = btn.dataset.status;
      applyFilters();
    });

    // Store filter dropdown
    DOM.filterStore.addEventListener('change', (e) => {
      state.storeFilter = e.target.value;
      applyFilters();
    });

    // Reset filters button in empty state
    DOM.btnResetFilters.addEventListener('click', () => {
      DOM.searchInput.value = '';
      state.searchQuery = '';
      DOM.btnClearSearch.classList.add('visually-hidden');
      DOM.filterStore.value = '';
      state.storeFilter = '';
      state.statusFilter = 'all';
      DOM.statusFilterTabs.querySelectorAll('.btn-tab').forEach(b => {
        b.classList.toggle('active', b.dataset.status === 'all');
      });
      applyFilters();
    });

    // Mark all currently visible batches completed
    DOM.btnMarkAll.addEventListener('click', () => {
      if (state.filteredBatches.length === 0) return;
      state.filteredBatches.forEach(b => state.completedKeys.add(b.key));
      saveCompletedKeys();
      applyFilters();
      playAudioBeep('success');
      showToast(`Ditandai selesai untuk ${state.filteredBatches.length} batch!`, 'success');
    });

    // Reset progress
    DOM.btnResetProgress.addEventListener('click', () => {
      if (confirm('Yakin ingin mereset checklist progres input untuk shift ini?')) {
        state.completedKeys.clear();
        saveCompletedKeys();
        applyFilters();
        showToast('Progres checklist berhasil direset', 'info');
      }
    });

    // Upload & Data Source
    if (DOM.btnBrowseFile) DOM.btnBrowseFile.addEventListener('click', () => DOM.fileInput && DOM.fileInput.click());
    if (DOM.fileInput) {
      DOM.fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files[0]) {
          handleExcelFile(e.target.files[0]);
        }
      });
    }

    if (DOM.btnReloadSample) DOM.btnReloadSample.addEventListener('click', () => loadDefaultSampleData());

    // Paste from Spreadsheet Modal Listeners
    if (DOM.btnOpenPasteModal) DOM.btnOpenPasteModal.addEventListener('click', openPasteModal);
    if (DOM.btnClosePaste) DOM.btnClosePaste.addEventListener('click', closePasteModal);
    if (DOM.btnCancelPaste) DOM.btnCancelPaste.addEventListener('click', closePasteModal);
    if (DOM.btnClearPaste) {
      DOM.btnClearPaste.addEventListener('click', () => {
        if (DOM.pasteTextarea) {
          DOM.pasteTextarea.value = '';
          updatePasteDetectionInfo();
          DOM.pasteTextarea.focus();
        }
      });
    }
    if (DOM.pasteTextarea) {
      DOM.pasteTextarea.addEventListener('input', updatePasteDetectionInfo);
    }
    if (DOM.btnProcessPaste) {
      DOM.btnProcessPaste.addEventListener('click', processPastedSpreadsheet);
    }

    if (DOM.dropZone) {
      DOM.dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        DOM.dropZone.classList.add('dragover');
      });

      DOM.dropZone.addEventListener('dragleave', () => {
        DOM.dropZone.classList.remove('dragover');
      });

      DOM.dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        DOM.dropZone.classList.remove('dragover');
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
          handleExcelFile(e.dataTransfer.files[0]);
        }
      });
    }

    if (DOM.btnExportExcel) {
      DOM.btnExportExcel.addEventListener('click', exportGroupedExcel);
    }

    // Delegate batch grid actions: copy, checkbox toggle, open simulator
    DOM.batchGrid.addEventListener('click', (e) => {
      // 1. Direct Shipment Copy Button
      const btnCopyDirect = e.target.closest('.btn-copy-direct, .btn-copy-shipment');
      if (btnCopyDirect) {
        const text = btnCopyDirect.dataset.copy;
        copyText(text, btnCopyDirect, 'Nomor shipment (koma) tersalin!');
        return;
      }

      // 2. Direct click on shipment text itself
      const shipmentTextElem = e.target.closest('.batch-shipment-text');
      if (shipmentTextElem) {
        const text = shipmentTextElem.dataset.copy;
        const directBtn = shipmentTextElem.parentElement ? shipmentTextElem.parentElement.querySelector('.btn-copy-direct') : null;
        copyText(text, directBtn, 'Nomor shipment (koma) tersalin!');
        return;
      }

      // 3. Inline Field Copy Button
      const btnCopyInline = e.target.closest('.btn-copy-inline, .btn-copy-tiny');
      if (btnCopyInline) {
        const text = btnCopyInline.dataset.copy;
        const msg = (text !== undefined && text !== null && text !== '') ? `Tersalin: "${text}"` : 'Personil kosong (biarkan kosong pada form) ✅';
        copyText(text, btnCopyInline, msg);
        return;
      }

      // 4. Open Simulator Button on Card
      const btnOpenSim = e.target.closest('.btn-open-sim-batch');
      if (btnOpenSim) {
        const batchId = parseInt(btnOpenSim.dataset.batchId, 10);
        const idx = state.batches.findIndex(b => b.id === batchId);
        if (idx !== -1) openSimulator(idx);
        return;
      }

      // 5. Copy all fields of batch (Exact Sequence 1 to 6)
      const btnSummary = e.target.closest('.btn-copy-summary');
      if (btnSummary) {
        const batchId = parseInt(btnSummary.dataset.batchId, 10);
        const b = state.batches.find(it => it.id === batchId);
        if (b) {
          const personil = b.jumlahPersonil || '';
          const pejabat = (b.namaPejabat && b.namaPejabat !== '-') ? b.namaPejabat : '';
          const tabSeparated = [
            b.shipmentMergedString,
            personil,
            pejabat,
            b.motorStatus,
            b.hasilFollowUp,
            b.catatanTambahan || ''
          ].join('\t');
          copyText(tabSeparated, btnSummary, 'Semua field (Urutan 1-6) berhasil disalin!');
        }
        return;
      }
    });

    // Delegate Checkbox Change for Large Selesai Button with Visual Effect
    DOM.batchGrid.addEventListener('change', (e) => {
      const checkbox = e.target.closest('.batch-checkbox');
      if (checkbox) {
        const key = checkbox.dataset.key;
        const card = checkbox.closest('.batch-card');
        const toggleBtn = checkbox.closest('.btn-toggle-selesai');
        const batchId = card ? card.dataset.batchId : '';
        const bObj = state.batches.find(b => b.key === key || String(b.id) === String(batchId));

        if (checkbox.checked) {
          state.completedKeys.add(key);
          if (bObj) bObj.isCompleted = true;
          playAudioBeep('success');
          showToast(`Batch #${batchId} Ditandai Selesai! ✅`, 'success');
          if (card) card.classList.add('is-completed');
          if (toggleBtn) {
            toggleBtn.classList.add('is-completed');
            const lbl = toggleBtn.querySelector('.selesai-label-text');
            if (lbl) lbl.textContent = 'Selesai ✅';
          }
        } else {
          state.completedKeys.delete(key);
          if (bObj) bObj.isCompleted = false;
          playAudioBeep('copy');
          showToast(`Batch #${batchId} Ditandai Belum Selesai`, 'info');
          if (card) card.classList.remove('is-completed');
          if (toggleBtn) {
            toggleBtn.classList.remove('is-completed');
            const lbl = toggleBtn.querySelector('.selesai-label-text');
            if (lbl) lbl.textContent = 'Tandai Selesai';
          }
        }

        saveCompletedKeys();
        updateStats();
        triggerDockPulseAnimation();

        // If filtering by pending or completed status, update list after animation
        if (state.statusFilter !== 'all') {
          setTimeout(() => applyFilters(), 360);
        }
      }
    });

    // Rekap Semua Catatan Tambahan (Kendala) Copy Listeners
    if (DOM.btnCopyAllKendala) {
      DOM.btnCopyAllKendala.addEventListener('click', () => {
        const text = DOM.rekapKendalaText ? DOM.rekapKendalaText.textContent.trim() : '';
        if (text && !DOM.rekapKendalaText.classList.contains('is-empty')) {
          copyText(text, DOM.btnCopyAllKendala, 'Semua Catatan Kendala (koma) berhasil disalin!');
        } else {
          showToast('Tidak ada catatan kendala untuk disalin!', 'warning');
        }
      });
    }

    if (DOM.rekapKendalaText) {
      DOM.rekapKendalaText.addEventListener('click', () => {
        const text = DOM.rekapKendalaText.textContent.trim();
        if (text && !DOM.rekapKendalaText.classList.contains('is-empty')) {
          copyText(text, DOM.btnCopyAllKendala, 'Semua Catatan Kendala (koma) berhasil disalin!');
        }
      });
    }

    // Open Simulator via Header & Wizard buttons (guarded)
    if (DOM.btnOpenSimulator) DOM.btnOpenSimulator.addEventListener('click', () => openSimulator(0));
    if (DOM.btnStartWizard) {
      DOM.btnStartWizard.addEventListener('click', () => {
        let firstPending = state.batches.findIndex(b => !b.isCompleted);
        if (firstPending === -1) firstPending = 0;
        openSimulator(firstPending);
      });
    }

    // Simulator Navigation
    if (DOM.simBtnPrev) {
      DOM.simBtnPrev.addEventListener('click', () => {
        if (state.simulatorIndex > 0) {
          state.simulatorIndex--;
          populateSimulatorValues();
        }
      });
    }

    if (DOM.simBtnNext) {
      DOM.simBtnNext.addEventListener('click', () => {
        if (state.simulatorIndex < state.batches.length - 1) {
          state.simulatorIndex++;
          populateSimulatorValues();
        }
      });
    }

    if (DOM.simBtnClose) DOM.simBtnClose.addEventListener('click', closeSimulator);

    // Simulator copy all field button (Exact Sequence 1 to 6)
    if (DOM.simBtnCopyAll) {
      DOM.simBtnCopyAll.addEventListener('click', () => {
        const batch = state.batches[state.simulatorIndex];
        if (!batch) return;
        const personil = DOM.simInputPersonil.value.trim();
        const pejabat = DOM.simInputPejabat.value.trim();
        const motor = DOM.simMotorBagus.checked ? 'Bagus' : 'Rusak';
        const allText = [
          batch.shipmentMergedString,
          personil,
          pejabat,
          motor,
          DOM.simSelectHasilFu.value,
          DOM.simInputCatatan.value.trim()
        ].join('\t');
        copyText(allText, DOM.simBtnCopyAll, 'Semua field form (Urutan 1-6) tersalin (Tab-separated)!');
      });
    }

    // Simulator field copy buttons
    if (DOM.modalSimulator) {
      DOM.modalSimulator.addEventListener('click', (e) => {
        const copyBtn = e.target.closest('.btn-field-copy');
        if (!copyBtn) return;

        const targetId = copyBtn.dataset.copyTarget;
        if (targetId) {
          const inputElem = document.getElementById(targetId);
          if (inputElem) {
            copyText(inputElem.value, copyBtn, `Tersalin: "${inputElem.value}"`);
          }
        } else if (copyBtn.id === 'btn-copy-motor-val') {
          const motorVal = DOM.simMotorBagus.checked ? 'Bagus' : 'Rusak';
          copyText(motorVal, copyBtn, `Tersalin Status Motor: ${motorVal}`);
        }
      });
    }

    // Simulator Submissions
    if (DOM.simBtnSubmit) DOM.simBtnSubmit.addEventListener('click', submitSimulatorForm);
    if (DOM.simBtnCompleteNext) DOM.simBtnCompleteNext.addEventListener('click', submitSimulatorForm);

    // Keyboard navigation & Escape key inside modals
    window.addEventListener('keydown', (e) => {
      if (DOM.modalSimulator && !DOM.modalSimulator.classList.contains('visually-hidden')) {
        if (e.key === 'Escape') {
          closeSimulator();
        } else if (e.key === 'ArrowLeft' && !e.target.matches('input, textarea')) {
          if (state.simulatorIndex > 0) {
            state.simulatorIndex--;
            populateSimulatorValues();
          }
        } else if (e.key === 'ArrowRight' && !e.target.matches('input, textarea')) {
          if (state.simulatorIndex < state.batches.length - 1) {
            state.simulatorIndex++;
            populateSimulatorValues();
          }
        }
      } else if (DOM.modalRecap && !DOM.modalRecap.classList.contains('visually-hidden')) {
        if (e.key === 'Escape') closeRecapModal();
      } else if (DOM.modalPaste && !DOM.modalPaste.classList.contains('visually-hidden')) {
        if (e.key === 'Escape') closePasteModal();
      }
    });

    // Recap Modal
    if (DOM.btnOpenRecap) DOM.btnOpenRecap.addEventListener('click', openRecapModal);
    if (DOM.recapBtnClose) DOM.recapBtnClose.addEventListener('click', closeRecapModal);
    if (DOM.recapBtnDismiss) DOM.recapBtnDismiss.addEventListener('click', closeRecapModal);
    if (DOM.recapBtnCopy && DOM.recapTextOutput) {
      DOM.recapBtnCopy.addEventListener('click', () => {
        copyText(DOM.recapTextOutput.value, DOM.recapBtnCopy, 'Rangkuman WhatsApp berhasil disalin!');
      });
    }

    // Close modals on backdrop click
    if (DOM.modalSimulator) {
      DOM.modalSimulator.addEventListener('click', (e) => {
        if (e.target === DOM.modalSimulator) closeSimulator();
      });
    }
    if (DOM.modalRecap) {
      DOM.modalRecap.addEventListener('click', (e) => {
        if (e.target === DOM.modalRecap) closeRecapModal();
      });
    }
    if (DOM.modalPaste) {
      DOM.modalPaste.addEventListener('click', (e) => {
        if (e.target === DOM.modalPaste) closePasteModal();
      });
    }
  }

  /* ==========================================================================
     Application Initialization
     ========================================================================== */
  function init() {
    applyTheme(state.theme);
    loadCompletedKeys();
    initEventListeners();
    loadDefaultSampleData();
  }

  // Start when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
