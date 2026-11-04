/* Aplikasi Rekap Asistensi Kemiskinan PPPM
   Frontend: HTML/CSS/JS. Backend: Supabase Auth + Postgres + RLS. */
(() => {
  const cfg = window.APP_CONFIG || {};
  const configured = cfg.SUPABASE_URL && cfg.SUPABASE_ANON_KEY &&
    !cfg.SUPABASE_URL.includes("PASTE_") && !cfg.SUPABASE_ANON_KEY.includes("PASTE_");
  const $ = (id) => document.getElementById(id);
  const state = { supabase: null, user: null, profile: null, records: [], filtered: [], importBatches: [], importTotals: [], workbook: null, importRows: [], importTotalsPreview: [], currentFileName: '' };
  const money = n => new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(Number(n || 0));
  const number = n => new Intl.NumberFormat('id-ID', { maximumFractionDigits: 2 }).format(Number(n || 0));
  const clean = v => String(v ?? '').replace(/\s+/g, ' ').trim();
  const numeric = v => {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    let s = String(v).trim().replace(/[Rp\s]/gi, '');
    if (!s) return null;
    // Format Indonesia: 1.234.567,89; format Excel umum: 1,234,567
    if (s.includes(',') && s.includes('.')) {
      if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
      else s = s.replace(/,/g, '');
    } else if (s.includes(',')) {
      const parts = s.split(',');
      s = parts.length === 2 && parts[1].length <= 2 ? parts[0].replace(/\./g, '') + '.' + parts[1] : s.replace(/,/g, '');
    }
    const parsed = Number(s);
    return Number.isFinite(parsed) ? parsed : null;
  };
  const escapeHtml = v => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const showMessage = (id, msg, kind = '') => { const el = $(id); if (!el) return; el.textContent = msg; el.className = 'message ' + kind; };
  const toast = (msg, error = false) => { const el = $('toast'); el.textContent = msg; el.className = 'toast' + (error ? ' error' : ''); el.classList.remove('hidden'); setTimeout(() => el.classList.add('hidden'), 3500); };
  const setBusy = (button, busy, busyText = 'Memproses...') => { if (!button) return; if (busy) { button.dataset.oldText = button.textContent; button.textContent = busyText; button.disabled = true; } else { button.textContent = button.dataset.oldText || button.textContent; button.disabled = false; } };

  function init() {
    if (!configured || !window.supabase) {
      $('setupNotice').textContent = 'Konfigurasi Supabase belum diisi. Ikuti README.md: buat project Supabase, jalankan schema.sql, lalu isi config.js.';
      $('setupNotice').classList.remove('hidden');
    } else {
      state.supabase = window.supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY);
      state.supabase.auth.getSession().then(({ data }) => {
        if (data.session) { state.user = data.session.user; afterLogin(); }
      });
      state.supabase.auth.onAuthStateChange((_event, session) => {
        if (session && !state.user) { state.user = session.user; setTimeout(afterLogin, 0); }
        else if (!session) { state.user = null; state.profile = null; $('appView').classList.add('hidden'); $('authView').classList.remove('hidden'); }
      });
    }
    bindEvents();
  }

  function bindEvents() {
    $('authForm').addEventListener('submit', login);
    $('signupBtn').addEventListener('click', signup);
    $('logoutBtn').addEventListener('click', async () => { if (state.supabase) await state.supabase.auth.signOut(); state.user = null; state.profile = null; $('appView').classList.add('hidden'); $('authView').classList.remove('hidden'); });
    document.querySelectorAll('.nav-item').forEach(b => b.addEventListener('click', () => showPage(b.dataset.page)));
    document.querySelectorAll('[data-goto]').forEach(b => b.addEventListener('click', () => showPage(b.dataset.goto)));
    $('addRecordBtn').addEventListener('click', () => openRecordModal());
    $('closeModalBtn').addEventListener('click', closeRecordModal);
    $('cancelModalBtn').addEventListener('click', closeRecordModal);
    $('recordModal').addEventListener('click', e => { if (e.target === $('recordModal')) closeRecordModal(); });
    $('recordForm').addEventListener('submit', saveRecord);
    ['searchInput','filterYear','filterImpact','filterAgency','filterBatch'].forEach(id => $(id).addEventListener(id === 'searchInput' ? 'input' : 'change', renderDataTable));
    $('clearFilters').addEventListener('click', () => { $('searchInput').value=''; $('filterYear').value=''; $('filterImpact').value=''; $('filterAgency').value=''; $('filterBatch').value=''; renderDataTable(); });
    $('exportExcelBtn').addEventListener('click', () => exportExcel(state.filtered, 'rekap-data-pppm.xlsx'));
    $('exportPdfBtn').addEventListener('click', () => exportPdf(state.filtered, 'Laporan Rekap Data Asistensi Kemiskinan PPPM'));
    $('reportExcelBtn').addEventListener('click', exportReportExcel);
    $('reportPdfBtn').addEventListener('click', exportReportPdf);
    $('refreshDashboard').addEventListener('click', renderDashboard);
    $('refreshReport').addEventListener('click', () => { renderReport(); toast('Rekap berhasil diperbarui.'); });
    ['dashboardYear','dashboardAgency'].forEach(id => $(id).addEventListener('change', renderDashboard));
    ['reportYear','reportAgency'].forEach(id => $(id).addEventListener('change', renderReport));
    $('excelFile').addEventListener('change', handleExcelFile);
    $('previewImportBtn').addEventListener('click', previewImport);
    $('runImportBtn').addEventListener('click', runImport);
    $('sheetSelect').addEventListener('change', () => {
      const selected = $('sheetSelect').value;
      state.importRows = [];
      state.importTotalsPreview = [];
      $('previewPanel').classList.add('hidden');
      $('runImportBtn').disabled = true;
      if (state.workbook) $('previewImportBtn').disabled = false;
    });
  }

  async function login(e) {
    e.preventDefault();
    if (!state.supabase) return showMessage('authMessage','Konfigurasi Supabase belum selesai. Buka README.md dan selesaikan bagian persiapan.','error');
    const btn = $('loginBtn'); setBusy(btn, true, 'Memeriksa...');
    const { error } = await state.supabase.auth.signInWithPassword({ email: $('email').value.trim(), password: $('password').value });
    setBusy(btn, false);
    if (error) showMessage('authMessage', error.message, 'error');
    else showMessage('authMessage','Berhasil masuk.','success');
  }

  async function signup() {
    if (!state.supabase) return showMessage('authMessage','Konfigurasi Supabase belum selesai.','error');
    const email = $('email').value.trim(), password = $('password').value;
    if (!email || password.length < 6) return showMessage('authMessage','Isi email dan kata sandi minimal 6 karakter terlebih dahulu.','error');
    const btn = $('signupBtn'); setBusy(btn, true, 'Mendaftarkan...');
    const { data, error } = await state.supabase.auth.signUp({ email, password });
    setBusy(btn, false);
    if (error) return showMessage('authMessage', error.message, 'error');
    if (data.session) showMessage('authMessage','Akun berhasil dibuat. Kamu masuk sebagai anggota.','success');
    else showMessage('authMessage','Pendaftaran diterima. Periksa email untuk konfirmasi, lalu masuk.','success');
  }

  async function afterLogin() {
    if (!state.supabase || !state.user) return;
    $('authView').classList.add('hidden'); $('appView').classList.remove('hidden');
    $('userEmail').textContent = state.user.email || 'Pengguna';
    $('userInitial').textContent = (state.user.email || 'U').charAt(0).toUpperCase();
    const { data: profile, error } = await state.supabase.from('profiles').select('id,email,role,created_at').eq('id', state.user.id).maybeSingle();
    if (error) { toast('Gagal membaca profil. Pastikan schema.sql sudah dijalankan.', true); return; }
    state.profile = profile || { role: 'member', email: state.user.email };
    $('userRole').textContent = state.profile.role === 'admin' ? 'Admin' : 'Anggota';
    document.querySelectorAll('.admin-only').forEach(el => el.classList.toggle('hidden', state.profile.role !== 'admin'));
    await loadRecords();
    showPage('dashboard');
  }

  async function loadRecords() {
    if (!state.supabase || !state.user) return;

    const [programResult, batchResult, totalResult] = await Promise.all([
      state.supabase.from('program_kemiskinan').select('*')
        .order('tahun', { ascending: true })
        .order('perangkat_daerah', { ascending: true }).limit(10000),
      state.supabase.from('import_batches').select('*')
        .order('created_at', { ascending: false }).limit(1000),
      state.supabase.from('import_agency_totals').select('*').order('tahun', { ascending: true }).limit(10000)
    ]);

    if (programResult.error) {
      toast('Gagal memuat data: ' + programResult.error.message, true);
      state.records = [];
    } else state.records = programResult.data || [];

    if (batchResult.error) {
      toast('Gagal memuat daftar kelompok Excel: ' + batchResult.error.message + '. Pastikan SQL tambahan sudah dijalankan.', true);
      state.importBatches = [];
    } else state.importBatches = batchResult.data || [];

    if (totalResult.error) {
      toast('Gagal memuat total per perangkat daerah: ' + totalResult.error.message + '. Jalankan SQL migrasi terbaru.', true);
      state.importTotals = [];
    } else state.importTotals = totalResult.data || [];
    const batchNames = new Map(state.importBatches.map(b => [b.id, b.file_name]));
    state.records = state.records.map(r => ({
      ...r,
      source_file: r.import_batch_id ? (batchNames.get(r.import_batch_id) || 'File Excel') : 'Data lama / input manual'
    }));

    populateAgencyFilters();
    populateRecordAgencyOptions();
    populateBatchFilters();
    renderDashboard(); renderDataTable(); renderReport();
  }

  function showPage(page) {
    document.querySelectorAll('.page').forEach(p => p.classList.add('hidden'));
    const target = $('page-' + page); if (target) target.classList.remove('hidden');
    document.querySelectorAll('.nav-item').forEach(b => b.classList.toggle('active', b.dataset.page === page));
    const titles = {dashboard:'Dashboard',data:'Data Program',import:'Impor Excel',reports:'Rekap & Laporan',users:'Pengguna'};
    $('pageTitle').textContent = titles[page] || 'Dashboard';
    if (page === 'users') loadUsers();
    if (page === 'dashboard') renderDashboard();
    if (page === 'data') renderDataTable();
    if (page === 'reports') renderReport();
  }

  function populateRecordAgencyOptions(extraNames = []) {
    const currentValue = $('recordAgency')?.value || '';
    const names = [...new Set([
      ...state.records.map(r => r.perangkat_daerah),
      ...state.importRows.map(r => r.perangkat_daerah),
      ...extraNames
    ].map(clean).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'id'));

    const list = $('agencyOptions');
    if (list) {
      list.innerHTML = names.map(name =>
        `<option value="${escapeHtml(name)}"></option>`
      ).join('');
    }

    // Jangan menghapus nama yang sedang diketik saat daftar diperbarui.
    if ($('recordAgency') && currentValue) $('recordAgency').value = currentValue;
  }

  function populateAgencyFilters() {
    const names = [...new Set(state.records.map(r => r.perangkat_daerah).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'id'));
    ['dashboardAgency','filterAgency','reportAgency'].forEach(id => {
      const el = $(id), old = el.value;
      el.innerHTML = '<option value="">Semua perangkat daerah</option>' + names.map(n=>`<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join('');
      if (names.includes(old)) el.value = old;
    });
  }

  function populateBatchFilters() {
    const el = $('filterBatch');
    if (!el) return;
    const old = el.value;
    el.innerHTML = '<option value="">Semua kelompok data</option>' +
      '<option value="__old__">Data lama / input manual</option>' +
      state.importBatches.map(b => `<option value="${escapeHtml(b.id)}">${escapeHtml(b.file_name)} · ${new Date(b.created_at).toLocaleDateString('id-ID')}</option>`).join('');
    if ([...el.options].some(o => o.value === old)) el.value = old;
  }
  function byFilters(records, year, agency) { return records.filter(r => (!year || String(r.tahun) === String(year)) && (!agency || r.perangkat_daerah === agency)); }
  function sum(rows, key) { return rows.reduce((a,r)=>a + (Number(r[key]) || 0), 0); }
  function renderDashboard() {
    const rows = byFilters(state.records, $('dashboardYear').value, $('dashboardAgency').value);
    const budget = sum(rows,'anggaran'), realization = sum(rows,'realisasi_anggaran');
    $('statPrograms').textContent = number(rows.length);
    $('statTarget').textContent = number(sum(rows,'target'));
    $('statBudget').textContent = money(budget);
    $('statRealization').textContent = money(realization);
    $('realizationPct').textContent = (budget ? (realization/budget*100).toLocaleString('id-ID',{maximumFractionDigits:1}) : '0') + '% dari anggaran';
    const years = [...new Set(rows.map(r=>Number(r.tahun)))].sort();
    const yearTotals = years.map(y=>({year:y,total:sum(rows.filter(r=>Number(r.tahun)===y),'anggaran')}));
    const max = Math.max(1,...yearTotals.map(x=>x.total));
    $('yearChart').className = yearTotals.length ? 'bar-chart' : 'bar-chart empty-state';
    $('yearChart').innerHTML = yearTotals.length ? yearTotals.map(x=>`<div class="bar-row"><strong>${x.year}</strong><div class="bar-track"><div class="bar-fill" style="width:${Math.max(1,x.total/max*100)}%"></div></div><span class="bar-value">${money(x.total)}</span></div>`).join('') : 'Belum ada data untuk ditampilkan.';
    const impacts = ['Langsung','Tidak Langsung'];
    const impactCounts = impacts.map(label=>({label,count:rows.filter(r=>(r.dampak||'').toLowerCase()===label.toLowerCase()).length}));
    $('impactSummary').className = rows.length ? 'summary-list' : 'summary-list empty-state';
    $('impactSummary').innerHTML = rows.length ? impactCounts.map(x=>`<div class="summary-item"><span><i class="bullet"></i>${x.label}</span><strong>${number(x.count)}</strong></div>`).join('') : 'Belum ada data untuk ditampilkan.';
    const recent = [...rows].sort((a,b)=>new Date(b.updated_at||b.created_at||0)-new Date(a.updated_at||a.created_at||0)).slice(0,5);
    $('recentRows').innerHTML = recent.length ? recent.map(r=>`<tr><td>${r.tahun}</td><td>${escapeHtml(r.perangkat_daerah)}</td><td>${escapeHtml(r.nomenklatur)}</td><td>${number(r.target)}</td><td>${money(r.anggaran)}</td></tr>`).join('') : '<tr><td colspan="5" class="empty-cell">Belum ada data. Impor file Excel untuk memulai.</td></tr>';
  }

  function renderDataTable() {
    const q = $('searchInput').value.trim().toLocaleLowerCase('id');
    const year = $('filterYear').value, impact = $('filterImpact').value, agency = $('filterAgency').value, batch = $('filterBatch').value;
    state.filtered = state.records.filter(r => {
      const hay = [r.perangkat_daerah,r.nomenklatur,r.kinerja,r.indikator,r.strategi,r.sasaran,r.lokasi].join(' ').toLocaleLowerCase('id');
      return (!q || hay.includes(q)) && (!year || String(r.tahun)===year) && (!impact || r.dampak===impact) && (!agency || r.perangkat_daerah===agency) && (!batch || (batch === '__old__' ? !r.import_batch_id : r.import_batch_id === batch));
    });
    $('tableCount').textContent = `${number(state.filtered.length)} dari ${number(state.records.length)} data`;
    $('dataRows').innerHTML = state.filtered.length ? state.filtered.map(r=>`<tr>
      <td>${r.tahun}</td><td>${escapeHtml(r.perangkat_daerah)}</td><td>${escapeHtml(r.nomenklatur)}</td><td>${escapeHtml(r.indikator)}</td>
      <td>${number(r.target)}</td><td>${r.capaian===null||r.capaian===undefined?'—':number(r.capaian)}</td><td>${money(r.anggaran)}</td><td>${r.realisasi_anggaran===null||r.realisasi_anggaran===undefined?'—':money(r.realisasi_anggaran)}</td><td>${escapeHtml(r.source_file || 'Data lama / input manual')}</td>
      <td class="action-cell"><button class="action-btn edit" data-edit="${r.id}">Edit</button><button class="action-btn delete" data-delete="${r.id}">Hapus</button></td></tr>`).join('') : '<tr><td colspan="10" class="empty-cell">Data tidak ditemukan.</td></tr>';
    const visibleBatchIds = new Set(state.filtered.map(r => r.import_batch_id).filter(Boolean));
    const totalFilter = state.importTotals.filter(t =>
      (!year || String(t.tahun) === year) && (!agency || t.perangkat_daerah === agency) &&
      (!batch || (batch === '__old__' ? false : t.import_batch_id === batch))
    );
    const totalsBody = $('agencyTotalsRows');
    if (totalsBody) totalsBody.innerHTML = totalFilter.length ? totalFilter.map(t => `<tr>
      <td>${t.tahun}</td><td>${escapeHtml(t.perangkat_daerah)}</td><td>${escapeHtml(t.sumber_sheet || '')}</td>
      <td>${t.total_target == null ? '—' : number(t.total_target)}</td><td>${t.total_capaian == null ? '—' : number(t.total_capaian)}</td>
      <td>${t.total_anggaran == null ? '—' : money(t.total_anggaran)}</td><td>${t.total_realisasi == null ? '—' : money(t.total_realisasi)}</td>
    </tr>`).join('') : '<tr><td colspan="7" class="empty-cell">Belum ada baris TOTAL pada file Excel yang diimpor atau filter tidak cocok.</td></tr>';
    if ($('agencyTotalsCount')) $('agencyTotalsCount').textContent = `${number(totalFilter.length)} total ditemukan`;
    $('dataRows').querySelectorAll('[data-edit]').forEach(b=>b.addEventListener('click',()=>openRecordModal(state.records.find(r=>r.id===b.dataset.edit))));
    $('dataRows').querySelectorAll('[data-delete]').forEach(b=>b.addEventListener('click',()=>deleteRecord(b.dataset.delete)));
  }

  function openRecordModal(record = null) {
    populateRecordAgencyOptions();
    $('recordForm').reset(); $('recordId').value = record?.id || '';
    $('modalTitle').textContent = record ? 'Edit data program' : 'Tambah data program';
    const map = {recordYear:'tahun',recordAgency:'perangkat_daerah',recordNomenclature:'nomenklatur',recordPerformance:'kinerja',recordIndicator:'indikator',recordUnit:'satuan',recordImpact:'dampak',recordStrategy:'strategi',recordTarget:'target',recordAchievement:'capaian',recordBudget:'anggaran',recordRealization:'realisasi_anggaran',recordTargetGroup:'sasaran',recordLocation:'lokasi'};
    for (const [id,key] of Object.entries(map)) $(id).value = record?.[key] ?? (key==='tahun' ? '2026' : '');
    showMessage('recordMessage','');
    $('recordModal').classList.remove('hidden');
  }
  function closeRecordModal() { $('recordModal').classList.add('hidden'); }
  async function saveRecord(e) {
    e.preventDefault(); if (!state.supabase) return;
    const id = $('recordId').value;
    const row = {
      tahun:Number($('recordYear').value), perangkat_daerah:$('recordAgency').value.trim(),
      nomenklatur:$('recordNomenclature').value.trim(), kinerja:$('recordPerformance').value.trim()||null,
      indikator:$('recordIndicator').value.trim()||null, satuan:$('recordUnit').value.trim()||null,
      dampak:$('recordImpact').value||null, strategi:$('recordStrategy').value.trim()||null,
      target:numeric($('recordTarget').value), capaian:numeric($('recordAchievement').value),
      anggaran:numeric($('recordBudget').value) ?? 0, realisasi_anggaran:numeric($('recordRealization').value),
      sasaran:$('recordTargetGroup').value.trim()||null, lokasi:$('recordLocation').value.trim()||null,
      updated_by:state.user.id
    };
    if (!row.perangkat_daerah || !row.nomenklatur) return showMessage('recordMessage','Tahun, perangkat daerah, dan nomenklatur wajib diisi.','error');
    const btn = $('saveRecordBtn'); setBusy(btn,true,'Menyimpan...');
    let result;
    if (id) result = await state.supabase.from('program_kemiskinan').update(row).eq('id',id);
    else result = await state.supabase.from('program_kemiskinan').insert({...row,created_by:state.user.id});
    setBusy(btn,false);
    if (result.error) return showMessage('recordMessage','Gagal menyimpan: '+result.error.message,'error');
    closeRecordModal(); toast(id?'Data berhasil diperbarui.':'Data berhasil ditambahkan.'); await loadRecords();
  }
  async function deleteRecord(id) {
    const r = state.records.find(x=>x.id===id);
    if (!r || !confirm(`Hapus data "${r.nomenklatur}"? Tindakan ini tidak dapat dibatalkan.`)) return;
    const {error} = await state.supabase.from('program_kemiskinan').delete().eq('id',id);
    if (error) toast('Gagal menghapus: '+error.message,true); else {toast('Data berhasil dihapus.'); await loadRecords();}
  }

  function parseWorkbookSheet(sheetName, yearOverride = null) {
    const ws = state.workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
    const year = Number(yearOverride || (String(sheetName).includes('2027') ? 2027 : 2026));
    const is2027 = year === 2027;
    let currentAgency = '';
    const records = [];
    const totals = [];
    for (let i = 3; i < rows.length; i++) {
      const row = rows[i] || [];
      const first = clean(row[0]);
      if (!first) continue;
      const hasOtherData = row.slice(1, is2027 ? 10 : 12).some(v => v !== null && v !== undefined && clean(v) !== '');
      if (!hasOtherData) { currentAgency = first; continue; }
      if (/^NOMENKLATUR URUSAN/i.test(first)) continue;

      // Baris TOTAL disimpan terpisah, tidak dimasukkan sebagai program.
      if (/^TOTAL\s*$/i.test(first)) {
        totals.push({
          tahun: year,
          perangkat_daerah: currentAgency || 'Belum ditentukan',
          sumber_sheet: sheetName,
          total_target: is2027 ? null : numeric(row[6]),
          total_capaian: is2027 ? null : numeric(row[7]),
          total_anggaran: numeric(row[is2027 ? 7 : 8]),
          total_realisasi: is2027 ? null : numeric(row[9])
        });
        continue;
      }

      const target = numeric(row[6]);
      const budget = numeric(row[is2027 ? 7 : 8]);
      const achievement = is2027 ? null : numeric(row[7]);
      const realization = is2027 ? null : numeric(row[9]);
      const targetGroup = is2027 ? row[8] : row[10];
      const location = is2027 ? row[9] : row[11];
      records.push({
        tahun: year, perangkat_daerah: currentAgency || 'Belum ditentukan',
        nomenklatur: first, kinerja: clean(row[1]) || null, indikator: clean(row[2]) || null,
        satuan: clean(row[3]) || null, strategi: clean(row[4]) || null, dampak: clean(row[5]) || null,
        target, capaian: achievement, anggaran: budget ?? 0, realisasi_anggaran: realization,
        sasaran: targetGroup === null ? null : clean(targetGroup) || null,
        lokasi: location === null ? null : clean(location) || null
      });
    }
    return { records: records.filter(r => r.nomenklatur && r.nomenklatur !== 'NOMENKLATUR URUSAN PROVINSI'), totals };
  }

  function selectedImportSheets() {
    const selected = $('sheetSelect').value;
    return selected === '__ALL__' ? state.workbook.SheetNames : [selected];
  }

  async function handleExcelFile(e) {
    const file = e.target.files?.[0]; if (!file) return;
    state.currentFileName = file.name;
    showMessage('importMessage','Membaca workbook...');
    try {
      const buffer = await file.arrayBuffer();
      state.workbook = XLSX.read(buffer,{type:'array',cellDates:true});
      $('sheetSelect').innerHTML = '<option value="__ALL__">Semua sheet</option>' +
        state.workbook.SheetNames.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
      $('sheetSelect').disabled = false;
      const firstYearSheet = state.workbook.SheetNames.find(s=>s.includes('2026')) || state.workbook.SheetNames[0];
      $('sheetSelect').value = '__ALL__';
      $('fileDetails').textContent = `${file.name} · ${state.workbook.SheetNames.length} sheet ditemukan`;
      $('fileDetails').classList.remove('hidden');
      $('previewImportBtn').disabled = false;
      state.importRows = [];
      state.importTotalsPreview = [];
      $('previewPanel').classList.add('hidden'); $('runImportBtn').disabled = true;
      showMessage('importMessage','File berhasil dibaca. Pilih satu sheet atau Semua sheet, lalu klik Pratinjau data.','success');
    } catch(err) { showMessage('importMessage','File tidak dapat dibaca: '+err.message,'error'); }
  }

  function previewImport() {
    if (!state.workbook) return;
    const sheets = selectedImportSheets();
    const parsedSheets = sheets.map(sheet => parseWorkbookSheet(sheet));
    state.importRows = parsedSheets.flatMap(result => result.records);
    state.importTotalsPreview = parsedSheets.flatMap(result => result.totals);
    // Hindari baris ganda yang muncul lebih dari sekali di workbook yang sama.
    const seen = new Set();
    state.importRows = state.importRows.filter(row => {
      const key = duplicateKey(row);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
    populateRecordAgencyOptions();
    $('previewPanel').classList.remove('hidden');
    $('previewSummary').textContent = `${state.importRows.length} baris program dan ${state.importTotalsPreview.length} baris total terdeteksi dari ${sheets.length} sheet (${sheets.join(', ')}). Baris TOTAL dipisahkan dari tabel program.`;
    $('previewRows').innerHTML = state.importRows.slice(0,10).map(r=>`<tr><td>${escapeHtml(r.perangkat_daerah)}</td><td>${escapeHtml(r.nomenklatur)}</td><td>${r.tahun}</td><td>${number(r.target)}</td><td>${money(r.anggaran)}</td><td>${escapeHtml(r.dampak)}</td></tr>`).join('') || '<tr><td colspan="6" class="empty-cell">Tidak ada baris data yang terdeteksi. Periksa struktur sheet.</td></tr>';
    $('runImportBtn').disabled = !state.importRows.length;
    showMessage('importMessage',state.importRows.length ? 'Periksa pratinjau. Jika sudah benar, lanjutkan impor.' : 'Tidak ada data yang bisa diimpor.','');
  }

  function duplicateKey(r) { return [r.tahun,clean(r.perangkat_daerah).toLowerCase(),clean(r.nomenklatur).toLowerCase(),clean(r.indikator).toLowerCase()].join('|'); }
  async function runImport() {
    if (!state.importRows.length || !state.supabase) return;
    const btn = $('runImportBtn'); setBusy(btn,true,'Mengimpor...');
    let batchId = null;
    let inserted = 0;
    try {
      let rows = [...state.importRows];
      if ($('duplicateMode').value === 'skip') {
        const existingKeys = new Set(state.records.map(duplicateKey));
        rows = rows.filter(r => {
          const key = duplicateKey(r);
          if (existingKeys.has(key)) return false;
          existingKeys.add(key);
          return true;
        });
      }
      if (!rows.length && !state.importTotalsPreview.length) {
        showMessage('importMessage','Semua data program sudah ada dan tidak ditemukan baris TOTAL baru. Tidak ada kelompok baru yang dibuat.','');
        return;
      }

      // Setiap file Excel yang berhasil diimpor mendapat ID kelompok baru.
      const { data: batchInfo, error: batchError } = await state.supabase
        .from('import_batches')
        .insert({
          file_name: state.currentFileName || 'File Excel tanpa nama',
          selected_sheets: selectedImportSheets(),
          total_rows: rows.length + state.importTotalsPreview.length,
          imported_rows: 0,
          created_by: state.user.id
        })
        .select('id')
        .single();
      if (batchError) throw new Error('Gagal membuat kelompok impor: ' + batchError.message);
      batchId = batchInfo.id;

      // Semua baris dari file ini menyimpan ID kelompok yang sama.
      for (let i=0; i<rows.length; i+=250) {
        const chunk = rows.slice(i,i+250).map(r => ({
          ...r,
          import_batch_id: batchId,
          created_by: state.user.id,
          updated_by: state.user.id
        }));
        const { error } = await state.supabase.from('program_kemiskinan').insert(chunk);
        if (error) throw new Error(error.message);
        inserted += chunk.length;
      }

      if (state.importTotalsPreview.length) {
        const totalRows = state.importTotalsPreview.map(t => ({ ...t, import_batch_id: batchId, created_by: state.user.id }));
        for (let i = 0; i < totalRows.length; i += 250) {
          const { error: totalError } = await state.supabase.from('import_agency_totals').insert(totalRows.slice(i, i + 250));
          if (totalError) throw new Error('Data program masuk, tetapi penyimpanan tabel total gagal: ' + totalError.message);
        }
      }

      const { error: updateBatchError } = await state.supabase
        .from('import_batches').update({ imported_rows: inserted }).eq('id', batchId);
      if (updateBatchError) throw new Error('Data masuk, tetapi jumlah kelompok belum terbarui: ' + updateBatchError.message);

      showMessage('importMessage',`${inserted} baris program dan ${state.importTotalsPreview.length} baris TOTAL berhasil disimpan sebagai kelompok baru: ${state.currentFileName}.`,'success');
      toast(`${inserted} data program dan ${state.importTotalsPreview.length} total berhasil disimpan.`);
      await loadRecords();
      $('filterBatch').value = batchId;
      renderDataTable();
    } catch(err) {
      if (batchId) {
        await state.supabase.from('import_batches').update({ imported_rows: inserted }).eq('id', batchId);
      }
      showMessage('importMessage',`Impor tidak selesai: ${err.message}. ${inserted ? inserted + ' baris sudah masuk ke kelompok ini; periksa kelompok tersebut sebelum mengulang.' : 'Belum ada baris yang berhasil disimpan.'}`,'error');
      if (batchId) await loadRecords();
    } finally { setBusy(btn,false); }
  }

  // Kolom ekspor dipusatkan agar Excel dan PDF memakai urutan data yang konsisten.
  const exportColumns = [
    { key: 'tahun', label: 'Tahun', width: 9 },
    { key: 'perangkat_daerah', label: 'Perangkat Daerah', width: 28 },
    { key: 'nomenklatur', label: 'Nomenklatur Urusan/Program', width: 42 },
    { key: 'kinerja', label: 'Kinerja', width: 32 },
    { key: 'indikator', label: 'Indikator', width: 34 },
    { key: 'satuan', label: 'Satuan', width: 14 },
    { key: 'strategi', label: 'Strategi', width: 34 },
    { key: 'dampak', label: 'Dampak', width: 18 },
    { key: 'target', label: 'Target', width: 14 },
    { key: 'capaian', label: 'Capaian', width: 14 },
    { key: 'anggaran', label: 'Anggaran (Rp)', width: 20 },
    { key: 'realisasi_anggaran', label: 'Realisasi Anggaran (Rp)', width: 22 },
    { key: 'sasaran', label: 'Sasaran', width: 28 },
    { key: 'lokasi', label: 'Lokasi', width: 28 },
    { key: 'source_file', label: 'Kelompok / Sumber Data', width: 32 }
  ];

  const exportValue = (value) =>
    value === null || value === undefined || value === '' ? '' : value;

  function makeExcelSheet(rows, sheetTitle = 'Data Asistensi Kemiskinan PPPM') {
    const headers = exportColumns.map(c => c.label);
    const data = rows.map(r => exportColumns.map(c => exportValue(r[c.key])));
    const aoa = [[sheetTitle], [`Tanggal ekspor: ${new Date().toLocaleString('id-ID')}`], headers, ...data];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = exportColumns.map(c => ({ wch: c.width }));
    ws['!rows'] = [{ hpt: 28 }, { hpt: 20 }, { hpt: 32 }, ...rows.map(() => ({ hpt: 34 }))];
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: exportColumns.length - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: exportColumns.length - 1 } }
    ];
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({
      s: { r: 2, c: 0 }, e: { r: rows.length + 2, c: exportColumns.length - 1 }
    }) };
    ws['!freeze'] = { xSplit: 0, ySplit: 3 };

    const border = {
      top: { style: 'thin', color: { rgb: 'D9DEEA' } },
      bottom: { style: 'thin', color: { rgb: 'D9DEEA' } },
      left: { style: 'thin', color: { rgb: 'D9DEEA' } },
      right: { style: 'thin', color: { rgb: 'D9DEEA' } }
    };
    for (let c = 0; c < exportColumns.length; c++) {
      const titleCell = ws[XLSX.utils.encode_cell({ r: 0, c })];
      if (titleCell) titleCell.s = {
        font: { name: 'Aptos Display', sz: 15, bold: true, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: '5143B4' } }, alignment: { vertical: 'center', horizontal: 'left' }
      };
      const subCell = ws[XLSX.utils.encode_cell({ r: 1, c })];
      if (subCell) subCell.s = {
        font: { name: 'Aptos', sz: 10, italic: true, color: { rgb: '667085' } },
        fill: { fgColor: { rgb: 'F4F3FC' } }, alignment: { vertical: 'center' }
      };
      const head = ws[XLSX.utils.encode_cell({ r: 2, c })];
      if (head) head.s = {
        font: { name: 'Aptos', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: '6658C9' } },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border
      };
    }
    rows.forEach((row, index) => {
      exportColumns.forEach((column, colIndex) => {
        const cellRef = XLSX.utils.encode_cell({ r: index + 3, c: colIndex });
        const cell = ws[cellRef];
        if (!cell) return;
        cell.s = {
          font: { name: 'Aptos', sz: 10, color: { rgb: '344054' } },
          fill: { fgColor: { rgb: index % 2 ? 'F7F8FC' : 'FFFFFF' } },
          alignment: { vertical: 'top', wrapText: true },
          border
        };
        if (['target', 'capaian'].includes(column.key)) {
          cell.z = '#,##0.##';
          cell.s.alignment.horizontal = 'right';
        }
        if (['anggaran', 'realisasi_anggaran'].includes(column.key)) {
          cell.z = '"Rp" #,##0';
          cell.s.alignment.horizontal = 'right';
        }
      });
    });
    return ws;
  }

  function totalsForRows(rows) {
    const years = new Set((rows || []).map(r => Number(r.tahun)));
    const agencies = new Set((rows || []).map(r => r.perangkat_daerah));
    const selectedBatch = $('filterBatch')?.value || '';
    const selectedYear = $('filterYear')?.value || '';
    const selectedAgency = $('filterAgency')?.value || '';
    return state.importTotals.filter(t => {
      if (selectedBatch && selectedBatch !== '__old__' && t.import_batch_id !== selectedBatch) return false;
      if (selectedBatch === '__old__') return false;
      if (selectedYear && String(t.tahun) !== selectedYear) return false;
      if (selectedAgency && t.perangkat_daerah !== selectedAgency) return false;
      if (years.size && !selectedYear && !years.has(Number(t.tahun))) return false;
      if (agencies.size && !selectedAgency && !agencies.has(t.perangkat_daerah)) return false;
      return true;
    });
  }

  function totalsForReportRows() {
    const selectedYear = $('reportYear')?.value || '';
    const selectedAgency = $('reportAgency')?.value || '';
    return state.importTotals.filter(t =>
      (!selectedYear || String(t.tahun) === String(selectedYear)) &&
      (!selectedAgency || t.perangkat_daerah === selectedAgency)
    );
  }

  const totalExportColumns = [
    { key: 'tahun', label: 'Tahun', width: 10 },
    { key: 'perangkat_daerah', label: 'Perangkat Daerah / Bidang / Badan', width: 48 },
    { key: 'sumber_sheet', label: 'Sheet Sumber', width: 18 },
    { key: 'total_target', label: 'Total Target', width: 16 },
    { key: 'total_capaian', label: 'Total Capaian', width: 16 },
    { key: 'total_anggaran', label: 'Total Anggaran (Rp)', width: 24 },
    { key: 'total_realisasi', label: 'Total Realisasi Anggaran (Rp)', width: 28 }
  ];

  function makeTotalsSheet(totals, title = 'TOTAL PER PERANGKAT DAERAH') {
    const aoa = [[title], [`Tanggal ekspor: ${new Date().toLocaleString('id-ID')}`], totalExportColumns.map(c => c.label),
      ...totals.map(t => totalExportColumns.map(c => t[c.key] ?? ''))];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = totalExportColumns.map(c => ({ wch: c.width }));
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: totalExportColumns.length - 1 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: totalExportColumns.length - 1 } }];
    ws['!rows'] = [{ hpt: 28 }, { hpt: 20 }, { hpt: 32 }, ...totals.map(() => ({ hpt: 26 }))];
    for (let c = 0; c < totalExportColumns.length; c++) {
      const h = ws[XLSX.utils.encode_cell({ r: 2, c })];
      if (h) h.s = { font: { name: 'Aptos', sz: 10, bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '6658C9' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } };
    }
    totals.forEach((t, i) => totalExportColumns.forEach((col, c) => {
      const cell = ws[XLSX.utils.encode_cell({ r: i + 3, c })];
      if (!cell) return;
      cell.s = { font: { name: 'Aptos', sz: 10, color: { rgb: '344054' } }, fill: { fgColor: { rgb: i % 2 ? 'F7F8FC' : 'FFFFFF' } }, alignment: { vertical: 'center', wrapText: true } };
      if (['total_target', 'total_capaian'].includes(col.key)) cell.z = '#,##0.##';
      if (['total_anggaran', 'total_realisasi'].includes(col.key)) cell.z = '"Rp" #,##0';
    }));
    return ws;
  }

  function exportExcel(rows, filename) {
    const totals = totalsForRows(rows || []);
    if (!rows?.length && !totals.length) return toast('Tidak ada data untuk diekspor.', true);
    const wb = XLSX.utils.book_new();
    if (rows?.length) XLSX.utils.book_append_sheet(wb, makeExcelSheet(rows), 'Data PPPM');
    if (totals.length) XLSX.utils.book_append_sheet(wb, makeTotalsSheet(totals), 'Total per Dinas');
    XLSX.writeFile(wb, filename, { compression: true });
  }

  function pdfBase(title, rows, columns, filename, options = {}, totals = []) {
    if (!rows?.length && !totals?.length) return toast('Tidak ada data untuk diekspor.', true);
    if (!window.jspdf?.jsPDF) return toast('Library PDF belum berhasil dimuat. Periksa koneksi internet.', true);

    const { jsPDF } = window.jspdf;
    // A3 landscape memberi ruang lebih lebar untuk seluruh kolom dalam satu tabel.
    const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a3' });
    const pageWidth = doc.internal.pageSize.getWidth();

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.text(title, 34, 34);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.text(`Dicetak: ${new Date().toLocaleString('id-ID')}`, 34, 49);
    doc.text(`Jumlah data: ${rows.length}`, pageWidth - 34, 49, { align: 'right' });

    const tableColumns = columns || exportColumns;
    const body = (rows || []).map(r => tableColumns.map(c => {
      const value = r[c.key];
      if (value === null || value === undefined || value === '') return '—';
      if (['anggaran', 'realisasi_anggaran', 'budget', 'realization'].includes(c.key)) return money(value);
      if (['target', 'capaian', 'count'].includes(c.key)) return number(value);
      return String(value);
    }));

    if (body.length) doc.autoTable({
      startY: 62,
      head: [tableColumns.map(c => c.label)],
      body,
      theme: 'grid',
      margin: { top: 62, right: 28, bottom: 34, left: 28 },
      styles: {
        font: 'helvetica',
        fontSize: options.fontSize || 6.5,
        cellPadding: 3,
        overflow: 'linebreak',
        valign: 'middle',
        lineColor: [218, 222, 232],
        lineWidth: 0.35,
        textColor: [45, 52, 68]
      },
      headStyles: {
        fillColor: [79, 70, 160],
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        fontSize: options.headerFontSize || 7,
        halign: 'center',
        valign: 'middle'
      },
      alternateRowStyles: { fillColor: [246, 247, 251] },
      columnStyles: options.columnStyles || {},
      showHead: 'everyPage',
      rowPageBreak: 'avoid',
      // Tidak memakai horizontalPageBreak: semua kolom berada dalam satu tabel utuh.
      didDrawPage: () => {
        const page = doc.internal.getNumberOfPages();
        const pageHeight = doc.internal.pageSize.getHeight();
        doc.setFontSize(7);
        doc.setTextColor(110);
        doc.text(`Halaman ${page}`, pageWidth - 30, pageHeight - 14, { align: 'right' });
      }
    });
    if (totals && totals.length) {
      const totalBody = totals.map(t => totalExportColumns.map(c => {
        const v = t[c.key];
        if (v === null || v === undefined || v === '') return '—';
        if (['total_anggaran', 'total_realisasi'].includes(c.key)) return money(v);
        if (['total_target', 'total_capaian'].includes(c.key)) return number(v);
        return String(v);
      }));
      let y = (doc.lastAutoTable?.finalY || 62) + 24;
      if (y > doc.internal.pageSize.getHeight() - 80) { doc.addPage(); y = 45; }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(35);
      doc.text('TABEL TOTAL PER PERANGKAT DAERAH / BIDANG / BADAN', 28, y);
      doc.autoTable({ startY: y + 10, head: [totalExportColumns.map(c => c.label)], body: totalBody, theme: 'grid',
        styles: { font: 'helvetica', fontSize: 7, cellPadding: 4, overflow: 'linebreak', valign: 'middle' },
        headStyles: { fillColor: [79, 70, 160], textColor: [255, 255, 255], fontStyle: 'bold', halign: 'center' },
        alternateRowStyles: { fillColor: [246, 247, 251] }, margin: { left: 28, right: 28, bottom: 34 }, rowPageBreak: 'avoid' });
    }
    doc.save(filename);
  }

  function exportPdf(rows, title) {
    pdfBase(title, rows, exportColumns, 'laporan-rekap-pppm.pdf', { fontSize: 6, headerFontSize: 6.5 }, totalsForRows(rows));
  }

  function reportRows() {
    return byFilters(state.records, $('reportYear').value, $('reportAgency').value);
  }

  function renderReport() {
    const rows = reportRows();
    const budget = sum(rows, 'anggaran');
    const realization = sum(rows, 'realisasi_anggaran');
    $('reportCount').textContent = number(rows.length);
    $('reportTarget').textContent = number(sum(rows, 'target'));
    $('reportBudget').textContent = money(budget);
    $('reportRealization').textContent = money(realization);

    const selectedYear = $('reportYear').value || 'Semua tahun';
    const selectedAgency = $('reportAgency').value || 'Semua perangkat daerah';
    $('reportStatus').textContent =
      `Rekap diperbarui: ${selectedYear} · ${selectedAgency} · ${number(rows.length)} data ditemukan.`;

    const grouped = new Map();
    rows.forEach(r => {
      const key = r.perangkat_daerah || 'Belum ditentukan';
      if (!grouped.has(key)) grouped.set(key, []);
      grouped.get(key).push(r);
    });
    const arr = [...grouped.entries()]
      .sort((a, b) => a[0].localeCompare(b[0], 'id'))
      .map(([name, rs]) => ({
        name, count: rs.length, target: sum(rs, 'target'),
        budget: sum(rs, 'anggaran'), realization: sum(rs, 'realisasi_anggaran')
      }));

    $('reportRows').innerHTML = arr.length
      ? arr.map(r => `<tr><td>${escapeHtml(r.name)}</td><td>${number(r.count)}</td><td>${number(r.target)}</td><td>${money(r.budget)}</td><td>${money(r.realization)}</td><td>${r.budget ? (r.realization / r.budget * 100).toLocaleString('id-ID', { maximumFractionDigits: 1 }) : '0'}%</td></tr>`).join('')
      : '<tr><td colspan="6" class="empty-cell">Belum ada data untuk direkap. Periksa filter tahun/perangkat daerah atau pastikan data sudah diimpor.</td></tr>';
    const importedTotals = state.importTotals.filter(t =>
      (!selectedYear || selectedYear === 'Semua tahun' || String(t.tahun) === String(selectedYear)) &&
      (!selectedAgency || selectedAgency === 'Semua perangkat daerah' || t.perangkat_daerah === selectedAgency)
    );
    const reportTotalsBody = $('reportImportedTotalsRows');
    if (reportTotalsBody) reportTotalsBody.innerHTML = importedTotals.length ? importedTotals.map(t => `<tr>
      <td>${t.tahun}</td><td>${escapeHtml(t.perangkat_daerah)}</td><td>${escapeHtml(t.sumber_sheet || '')}</td>
      <td>${t.total_target == null ? '—' : number(t.total_target)}</td><td>${t.total_capaian == null ? '—' : number(t.total_capaian)}</td>
      <td>${t.total_anggaran == null ? '—' : money(t.total_anggaran)}</td><td>${t.total_realisasi == null ? '—' : money(t.total_realisasi)}</td>
    </tr>`).join('') : '<tr><td colspan="7" class="empty-cell">Belum ada total dari file Excel untuk filter ini.</td></tr>';
  }

  function getReportSummary() {
    const grouped = new Map();
    reportRows().forEach(r => {
      const name = r.perangkat_daerah || 'Belum ditentukan';
      if (!grouped.has(name)) grouped.set(name, []);
      grouped.get(name).push(r);
    });
    return [...grouped.entries()]
      .sort((a, b) => a[0].localeCompare(b[0], 'id'))
      .map(([name, rs]) => ({
        name,
        count: rs.length,
        target: sum(rs, 'target'),
        budget: sum(rs, 'anggaran'),
        realization: sum(rs, 'realisasi_anggaran')
      }));
  }

  function exportReportExcel() {
    const rows = reportRows();
    if (!rows.length) return toast('Tidak ada data untuk diekspor.', true);

    const summary = getReportSummary().map(r => ({
      'Perangkat Daerah': r.name,
      'Jumlah Program': r.count,
      'Total Target': r.target,
      'Total Anggaran (Rp)': r.budget,
      'Realisasi Anggaran (Rp)': r.realization,
      'Realisasi (%)': r.budget ? r.realization / r.budget * 100 : 0
    }));

    const wb = XLSX.utils.book_new();
    const summaryRows = [
      ['REKAP ASISTENSI KEMISKINAN PPPM'],
      [`Tanggal ekspor: ${new Date().toLocaleString('id-ID')}`],
      Object.keys(summary[0]),
      ...summary.map(r => Object.values(r))
    ];
    const summaryWs = XLSX.utils.aoa_to_sheet(summaryRows);
    summaryWs['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 5 } }
    ];
    summaryWs['!rows'] = [{ hpt: 28 }, { hpt: 20 }, { hpt: 30 }, ...summary.map(() => ({ hpt: 26 }))];
    const summaryBorder = { top:{style:'thin',color:{rgb:'D9DEEA'}},bottom:{style:'thin',color:{rgb:'D9DEEA'}},left:{style:'thin',color:{rgb:'D9DEEA'}},right:{style:'thin',color:{rgb:'D9DEEA'}} };
    for (let c=0;c<6;c++) {
      for (const r of [0,1,2]) {
        const cell=summaryWs[XLSX.utils.encode_cell({r,c})];
        if (!cell) continue;
        cell.s = r===0
          ? {font:{name:'Aptos Display',sz:15,bold:true,color:{rgb:'FFFFFF'}},fill:{fgColor:{rgb:'5143B4'}},alignment:{vertical:'center'}}
          : r===1
          ? {font:{name:'Aptos',sz:10,italic:true,color:{rgb:'667085'}},fill:{fgColor:{rgb:'F4F3FC'}}}
          : {font:{name:'Aptos',sz:10,bold:true,color:{rgb:'FFFFFF'}},fill:{fgColor:{rgb:'6658C9'}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:summaryBorder};
      }
    }
    summary.forEach((row,i)=>Object.values(row).forEach((value,c)=>{
      const cell=summaryWs[XLSX.utils.encode_cell({r:i+3,c})];
      if(!cell)return;
      cell.s={font:{name:'Aptos',sz:10,color:{rgb:'344054'}},fill:{fgColor:{rgb:i%2?'F7F8FC':'FFFFFF'}},alignment:{vertical:'center'},border:summaryBorder};
      if(c>=3) cell.z=c===5?'0.0"%"':'"Rp" #,##0';
      if(c===1||c===2) cell.z='#,##0.##';
    }));
    summaryWs['!cols'] = [
      { wch: 32 }, { wch: 16 }, { wch: 16 }, { wch: 23 }, { wch: 25 }, { wch: 17 }
    ];
    summaryWs['!autofilter'] = { ref: XLSX.utils.encode_range({
      s: { r: 2, c: 0 }, e: { r: summary.length + 2, c: 5 }
    }) };
    summaryWs['!freeze'] = { xSplit: 0, ySplit: 3 };

    const detailWs = makeExcelSheet(rows, 'DETAIL DATA ASISTENSI KEMISKINAN PPPM');
    XLSX.utils.book_append_sheet(wb, summaryWs, 'Rekap per Dinas');
    XLSX.utils.book_append_sheet(wb, detailWs, 'Detail Data');
    const importedTotals = totalsForReportRows();
    if (importedTotals.length) XLSX.utils.book_append_sheet(wb, makeTotalsSheet(importedTotals), 'Total per Dinas');
    XLSX.writeFile(wb, 'rekap-laporan-pppm.xlsx', { compression: true });
  }

  function exportReportPdf() {
    const rows = reportRows();
    if (!rows.length) return toast('Tidak ada data untuk diekspor.', true);

    // PDF laporan juga berupa satu tabel detail lengkap, bukan tabel ringkasan terpisah.
    pdfBase(
      'Rekap Asistensi Kemiskinan PPPM',
      rows,
      exportColumns,
      'rekap-laporan-pppm.pdf',
      { fontSize: 6, headerFontSize: 6.5 },
      totalsForReportRows()
    );
  }

  async function loadUsers() {
    if (state.profile?.role !== 'admin') return;
    const {data,error}=await state.supabase.from('profiles').select('id,email,role,created_at').order('created_at',{ascending:true});
    if(error){$('userRows').innerHTML=`<tr><td colspan="4" class="empty-cell">${escapeHtml(error.message)}</td></tr>`;return;}
    $('userRows').innerHTML=(data||[]).map(u=>`<tr><td>${escapeHtml(u.email)}</td><td><span class="role-pill">${escapeHtml(u.role)}</span></td><td>${u.created_at?new Date(u.created_at).toLocaleDateString('id-ID'):'—'}</td><td>${u.id===state.user.id?'<span class="muted">Akun aktif</span>':`<button class="action-btn" data-role="${u.id}" data-next="${u.role==='admin'?'member':'admin'}">Jadikan ${u.role==='admin'?'anggota':'admin'}</button>`}</td></tr>`).join('')||'<tr><td colspan="4" class="empty-cell">Belum ada profil pengguna.</td></tr>';
    $('userRows').querySelectorAll('[data-role]').forEach(b=>b.addEventListener('click',async()=>{
      const next=b.dataset.next;
      if(!confirm(`Ubah peran pengguna ini menjadi ${next}?`))return;
      const {error}=await state.supabase.from('profiles').update({role:next}).eq('id',b.dataset.role);
      if(error)showMessage('usersMessage','Gagal mengubah peran: '+error.message,'error');
      else{showMessage('usersMessage','Peran pengguna berhasil diperbarui.','success');await loadUsers();}
    }));
  }

  init();
})();
