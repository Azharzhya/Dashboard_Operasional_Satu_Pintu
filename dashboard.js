/* =====================================================================
   Dashboard Operasional -- dashboard.js
   Urutan:  1 Konfigurasi · 2 Alat bantu · 3 API & sesi · 4 Tampilan umum
            5 Masuk · 6 Dashboard · 7 Rincian · 8 Ringkasan
            9 Baca Excel · 10 Unggah · 11 Unduh Excel · 12 Mulai
   Aturan: isi dari Sheet hanya dipasang lewat textContent / DOM, tidak pernah sebagai HTML.
   ===================================================================== */
(() => {
'use strict';

/* ---------------------------------------------------------------------
   1. Konfigurasi
   --------------------------------------------------------------------- */
const API_URL = 'https://script.google.com/macros/s/AKfycbz3mpy9zQunopiYQglyn2ct9SgPbOviPQCR0Z1uGE_MjcDnic3sTcsx7bcNyITnR76LcA/exec';
const SESSION_KEY = 'dash_sesi_v1';
const CACHE_DB = 'dash_cache_v3';   // v3: susunan kolom balasan berubah (sisa & angsuran dihitung di halaman, per kategori)
const CACHE_DB_OLD = 'dash_cache_v2';   // salinan versi lama: dihapus saat halaman dibuka
const API_TIMEOUT_MS = 60000;
const AUTO_SYNC_MS = 2 * 60 * 1000; // pemeriksaan ringan perubahan Wilayah Kerja/data setiap 2 menit saat aplikasi aktif
const AUTO_SYNC_START_JITTER_MS = 20 * 1000; // sebar waktu mulai antarpengguna agar request tidak menumpuk bersamaan
const AUTO_SYNC_MAX_PER_LOGIN = 3; // setelah 3 sinkron otomatis, percobaan otomatis ke-4 memaksa login ulang
let autoSyncTimer = 0;
let autoSyncStartTimer = 0;
const PAGE_SIZES = [25, 50, 100, 200];
const SLOW_MS = 400;                 // jeda sebelum rangka "memuat" ditampilkan, supaya data cache tidak berkedip

// Indeks kolom balasan webDashboard (lihat Dashboard.gs):
// 0 cabang, 1 loan, 2 client id, 3 nama, 4 produk, 5 kode/staf, 6 total balance, 7 cicilan,
// 8-14 simpanan (Wajib, Sukarela, Pensiun, Hari Raya, Khusus, Qurban, Sipadan; sudah dikurangi 2.000),
// 15 kum. simpanan (tanpa Wajib), 16 total simpanan (kum. + Wajib), 17 catatan, 18 Total Balance + Penyesuaian Margin, 19 kode kategori.
// 20 = kunci pencarian, dibuat di sini (tidak dikirim server).
// Sisa pelunasan, sisa tanpa margin, dan jumlah angsuran tidak dikirim: nilainya bergantung pada pilihan "Include: Simpanan Wajib" tiap kategori,
// jadi dihitung di sini untuk kategori yang sedang dibuka (rumus sama dengan dashCalc_ di Dashboard.gs).
const F = { BRANCH: 0, LOAN: 1, CLIENT: 2, NAME: 3, PRODUCT: 4, STAF: 5, BALANCE: 6, CICILAN: 7, SIMP: 8,
  SUKARELA: 9, KUM: 15, TOTAL: 16, NOTE: 17, TBADJ: 18, FLAGS: 19, KEY: 20 };
const SIMPANAN = ['Wajib', 'Sukarela', 'Pensiun', 'Hari Raya', 'Khusus', 'Qurban', 'Sipadan']; // = indeks 8..14

// Simp. Wajib ikut dihitung di kategori k? Pilihan dari Setting "Include: Simpanan Wajib"; server mengirim kodenya di d.wajib (contoh "1t").
const inc = k => !!(state.data && state.data.wajib.indexOf(k) > -1);
const wajibOf = (r, k) => inc(k) ? r[F.SIMP] : 0;       // Simp. Wajib yang dihitung di kategori k

// Catatan "Simpanan < 2.000 (...)": pada Pelunasan dari Simpanan (k = 's') yang tidak menghitung Simp. Wajib, namanya dibuang dari catatan.
function noteText(n, k) {
  if (k !== 's' || inc('s') || !n) return n;
  const m = /\((.*)\)/.exec(n);
  if (!m) return n;
  const items = m[1].split(', ').filter(x => x !== 'Wajib');
  return items.length ? 'Simpanan < 2.000 (' + items.join(', ') + ')' : '';
}

// Kolom tampilan. num = rata kanan & diurutkan sebagai angka. get(r, k): k = kode kategori yang sedang ditampilkan.
const col = (label, idx, num) => ({ label, get: r => r[idx], num: !!num });
const COLS = {
  branch: col('Branch Name', F.BRANCH), loan: col('Loan No', F.LOAN), client: col('Client ID', F.CLIENT),
  name: col('Client Name', F.NAME), product: col('Product Name', F.PRODUCT), staf: col('Kode/Staf', F.STAF),
  balance: col('Total Balance', F.BALANCE, true), balanceAdj: col('Total Balance + Margin', F.TBADJ, true), cicilan: col('Cicilan', F.CICILAN, true),
  sukarela: col('Simp. Sukarela', F.SUKARELA, true), wajib: col('Simp. Wajib', F.SIMP, true),
  kum: col('Kum. Simpanan', F.KUM, true), total: col('Total Simpanan', F.TOTAL, true),
  angsuran: { label: 'Jumlah Angsuran', get: (r, k) => r[F.CICILAN] > 0 ? Math.floor((r[F.SUKARELA] + wajibOf(r, k)) / r[F.CICILAN]) : 0, num: true, noSum: true },
  sisa: { label: 'Sisa Pelunasan', get: (r, k) => r[F.TBADJ] - r[F.KUM] - wajibOf(r, k), num: true },
  sisaTm: { label: 'Total Balance - Simpanan', get: (r, k) => r[F.BALANCE] - r[F.KUM] - wajibOf(r, k), num: true },
  note: { label: 'Catatan', get: (r, k) => noteText(r[F.NOTE], k), num: false },
  ya: { label: 'Keterangan', get: () => 'YA', mid: true },    // hanya di file Excel (tetap seperti versi lama)
};
const BASE = ['branch', 'client', 'loan', 'name', 'product', 'staf', 'balance'];

// k = kode kategori dari server. xl = kolom tambahan khusus Excel. rule = rumus yang tampil di Setting "Include: Simpanan Wajib".
// Judul kolom hitungan ada di extraOf(): bergantung pada apakah Simp. Wajib ikut dihitung di kategori itu.
const minus = k => inc(k) ? 'Kum. Simpanan dan Simp. Wajib' : 'Kum. Simpanan';
const RULE_PEL = 'Sisa = Total Balance + Margin - Kum. Simpanan - Simp. Wajib';
const CATS = [
  { k: 'a', title: 'Angsuran By Sukarela', sheet: 'Angsuran By Sukarela', xl: [], noNote: true, rule: 'Jumlah angsuran = (Simp. Sukarela + Simp. Wajib) / Cicilan',
    hint: () => (inc('a') ? 'Simp. Sukarela + Simp. Wajib' : 'Hanya Simp. Sukarela') + ': dibagi Cicilan (Pokok + Margin), dibulatkan ke bawah, minimal 1 angsuran.' },
  { k: '1', title: 'Pelunasan < 100.000', sheet: 'Pelunasan 1-100rb', xl: ['ya'], rule: RULE_PEL,
    hint: () => 'Total Balance + Margin dikurangi ' + minus('1') + ': 1 sampai 100.000.' },
  { k: '2', title: 'Pelunasan < 200.000', sheet: 'Pelunasan 100-200rb', xl: ['ya'], rule: RULE_PEL,
    hint: () => 'Total Balance + Margin dikurangi ' + minus('2') + ': lebih dari 100.000 sampai 200.000.' },
  { k: '5', title: 'Pelunasan < 500.000', sheet: 'Pelunasan 200-500rb', xl: ['ya'], rule: RULE_PEL,
    hint: () => 'Total Balance + Margin dikurangi ' + minus('5') + ': lebih dari 200.000 sampai 500.000.' },
  { k: 's', title: 'Pelunasan dari Simpanan', sheet: 'Pelunasan dari Simpanan', xl: ['ya'], rule: RULE_PEL,
    hint: () => minus('s') + ' sudah menutup Total Balance + Margin (sisa 0 atau minus).' },
  { k: 't', title: 'Pelunasan Tanpa Margin', sheet: 'Pelunasan Tanpa Margin', xl: ['ya'], rule: 'Sisa = Total Balance - Kum. Simpanan - Simp. Wajib (tanpa Margin)',
    hint: () => (inc('t') ? 'Total Simpanan' : 'Kum. Simpanan (tanpa Wajib)') + ' sudah menutup Total Balance (tanpa Penyesuaian Margin).' },
];
const catOf = k => CATS.find(c => c.k === k);
// Kolom hitungan tiap kategori. Simp. Wajib hanya muncul sebagai kolom bila ikut dihitung di kategori itu.
function extraOf(c) {
  const w = inc(c.k);
  if (c.k === 'a') return ['cicilan', 'sukarela'].concat(w ? ['wajib'] : [], ['angsuran']);
  if (c.k === 't') return [w ? 'total' : 'kum', 'sisaTm'];
  return ['balanceAdj', 'kum'].concat(w ? ['wajib'] : [], ['sisa']);
}
const screenCols = c => BASE.concat(extraOf(c), c.noNote ? [] : ['note']).map(n => COLS[n]);
const excelCols = c => BASE.concat(extraOf(c), c.xl, c.noNote ? [] : ['note']).map(n => COLS[n]);

// Kolom file Excel unggahan: [judul di Sheet, kolom Excel (0 = A), tipe: t teks, n angka, d tanggal].
// Urutan WAJIB sama dengan UNGGAH_COLS di UnggahExcel.gs.
const UP_COLS = [['No.', 0, 'n'], ['Branch Name', 1, 't'], ['Loan No', 2, 't'], ['Ctr ID', 3, 't'], ['Client ID', 4, 't'], ['Client Name', 5, 't'],
  ['Product Name', 7, 't'], ['Disbursed Date', 10, 'd'], ['Total Balance', 13, 'n'], ['Simp. Wajib', 21, 'n'], ['Simp. Sukarela', 22, 'n'],
  ['Simp. Pensiun', 23, 'n'], ['Simp. Hari Raya', 24, 'n'], ['Simp. Khusus', 25, 'n'], ['Simp. Qurban', 26, 'n'], ['Simp. Sipadan', 27, 'n'],
  ['Cicilan (Pokok + Margin)', 28, 'n'], ['Minggon', 32, 't'], ['Kode/Staf', 33, 't'], ['Top Up', 34, 't'], ['Trans Pinj Terakhir', 35, 'd'],
  ['Trans Tab Terakhir', 36, 'd']];
const UP_PREVIEW = [1, 2, 5, 6, 7, 8, 18];   // indeks UP_COLS yang tampil di pratinjau
const UP_FIRST_ROW = 5, UP_CHUNK = 1500, UP_MAX_MB = 40;
const XLSX_URL = 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.min.js';
const MONTHS_EN_ID = { jan: 1, feb: 2, mar: 3, apr: 4, mei: 5, may: 5, jun: 6, jul: 7, agu: 8, agt: 8, aug: 8, sep: 9, okt: 10, oct: 10, nov: 11, des: 12, dec: 12 };
const MONTHS_ID = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

/* ---------------------------------------------------------------------
   2. Alat bantu
   --------------------------------------------------------------------- */
const $ = id => document.getElementById(id);
const nf = new Intl.NumberFormat('id-ID');               // dibuat sekali: jauh lebih cepat daripada toLocaleString per sel
const nf1 = new Intl.NumberFormat('id-ID', { maximumFractionDigits: 1 });
const collator = new Intl.Collator('id', { numeric: true, sensitivity: 'base' });
// Normalisasi Client ID untuk deteksi duplikat; nilai kosong tidak dihitung.
const clientIdKey = v => String(v == null ? '' : v).trim().replace(/\s+/g, ' ').toUpperCase();
const fmt = n => nf.format(n);
const rp = n => 'Rp ' + fmt(n);
const pad = n => String(n).padStart(2, '0');
const sleep = ms => new Promise(r => setTimeout(r, ms));
const nextFrame = () => new Promise(r => requestAnimationFrame(() => setTimeout(r, 0)));

// Rupiah ringkas untuk kartu & angka utama. Angka lengkap tetap ada di atribut title.
function rpShort(n) {
  const a = Math.abs(n), s = n < 0 ? '-' : '';
  if (a >= 1e12) return s + 'Rp ' + nf1.format(a / 1e12) + ' triliun';
  if (a >= 1e9) return s + 'Rp ' + nf1.format(a / 1e9) + ' miliar';
  if (a >= 1e6) return s + 'Rp ' + nf1.format(a / 1e6) + ' juta';
  return rp(n);
}
// 'yyyy-MM-dd HH:mm' (dari server) -> '30 Okt 2026, 14.02'
function fmtStamp(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}))?/.exec(s || '');
  return m ? +m[3] + ' ' + MONTHS_ID[+m[2] - 1] + ' ' + m[1] + (m[4] ? ', ' + m[4] + '.' + m[5] : '') : s || '';
}
const today = () => { const d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); };

// Pembuat elemen kecil. Teks selalu lewat node teks, bukan HTML.
function h(tag, props, ...kids) {
  const e = document.createElement(tag);
  for (const k in props || {}) {
    const v = props[k];
    if (v == null || v === false) continue;
    if (k === 'class') e.className = v;
    else if (k === 'text') e.textContent = v;
    else if (k === 'style') e.style.cssText = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else e.setAttribute(k, v === true ? '' : v);
  }
  for (const c of kids.flat()) if (c != null && c !== false) e.append(c);
  return e;
}
// Seperti replaceChildren, tetapi mengabaikan null / false (replaceChildren akan menuliskannya sebagai teks "null").
const put = (el, ...kids) => el.replaceChildren(...kids.flat().filter(c => c != null && c !== false));
const debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };

let toastTimer;
function toast(text, bad) {
  const t = $('toast');
  t.textContent = text; t.className = 'toast' + (bad ? ' bad' : ''); t.hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.hidden = true; }, 3200);
}
function setNote(id, text, cls) { const e = $(id); e.textContent = text || ''; e.className = 'note ' + (cls || ''); }

/* ---------------------------------------------------------------------
   3. API & sesi
   --------------------------------------------------------------------- */
const session = {
  get() { try { return JSON.parse(localStorage.getItem(SESSION_KEY) || 'null'); } catch (e) { return null; } },
  set(v) { try { localStorage.setItem(SESSION_KEY, JSON.stringify(v)); } catch (e) { /* mode privat: abaikan */ } },
  clear() { try { localStorage.removeItem(SESSION_KEY); } catch (e) { /* abaikan */ } },
};

// Salinan data Dashboard di perangkat (IndexedDB) supaya kunjungan berikutnya langsung tampil. Dihapus saat Keluar.
const cache = (() => {
  const open = () => new Promise((ok, no) => {
    const q = indexedDB.open(CACHE_DB, 1);
    q.onupgradeneeded = () => q.result.createObjectStore('kv');
    q.onsuccess = () => ok(q.result); q.onerror = () => no(q.error);
  });
  const run = async (mode, fn) => {
    try {
      const db = await open();
      return await new Promise((ok, no) => {
        const tx = db.transaction('kv', mode), req = fn(tx.objectStore('kv'));
        tx.oncomplete = () => { db.close(); ok(req && req.result); }; tx.onerror = tx.onabort = () => { db.close(); no(tx.error); };
      });
    } catch (e) { return null; }                       // IndexedDB tidak tersedia: aplikasi tetap jalan tanpa cache
  };
  return { get: () => run('readonly', s => s.get('dash')), set: v => run('readwrite', s => s.put(v, 'dash')), clear: () => run('readwrite', s => s.clear()) };
})();

const fail = error => ({ ok: false, error });
const MSG_OLD_SERVER = 'Server belum mendukung pembatasan Wilayah Kerja. Perbarui Dashboard.gs dan Akses.gs di Apps Script, lalu Deploy > Manage deployments > Edit > Version: New version > Deploy.';
const MSG_ROUTING = 'Server belum meneruskan aksi web. Di Kode.gs, ganti baris webLogin menjadi if (/^web[A-Z]/.test(p.action)) return handleAksesRoute(p); lalu Deploy > Manage deployments > Edit > Version: New version > Deploy.';

// POST form-urlencoded = "simple request": tanpa preflight CORS, cocok untuk Apps Script.
// retry: satu kali ulang bila jaringan putus (untuk aksi baca saja; aksi tulis tidak pernah diulang di sini).
async function api(action, data, opts) {
  const body = new URLSearchParams({ action });
  for (const k in data || {}) if (data[k] != null) body.set(k, data[k]);
  for (let attempt = 0; ; attempt++) {
    const ctl = new AbortController(), timer = setTimeout(() => ctl.abort(), (opts && opts.timeoutMs) || API_TIMEOUT_MS);
    try {
      const t0 = performance.now(), res = await fetch(API_URL, { method: 'POST', body, signal: ctl.signal });
      const tHead = performance.now();                  // jawaban pertama server tiba = server selesai menyiapkan data
      if (opts && opts.onHead) opts.onHead(tHead - t0);
      let text = '', got = 0;
      if (opts && opts.onProgress && res.body && res.body.getReader) {
        const rd = res.body.getReader(), dec = new TextDecoder();
        for (;;) { const c = await rd.read(); if (c.done) break; got += c.value.length; text += dec.decode(c.value, { stream: true }); opts.onProgress(got); }
        text += dec.decode();
      } else text = await res.text();
      const tBody = performance.now();
      let json;
      try { json = JSON.parse(text); } catch (e) { return fail('Balasan server tidak terbaca. Coba lagi sebentar lagi.'); }
      if (opts && opts.onDone) opts.onDone({ wait: tHead - t0, download: tBody - tHead, bytes: got || text.length });
      // Kode.gs lama menjawab aksi web yang tidak dikenal dengan pesan login ini
      if (action !== 'webLogin' && json && json.error === 'NIK dan Password wajib diisi.') return fail(MSG_ROUTING);
      return json;
    } catch (e) {
      if (opts && opts.retry && attempt === 0 && e.name !== 'AbortError') { await sleep(800); continue; }
      return fail(e.name === 'AbortError' ? 'Server terlalu lama menjawab. Coba lagi.' : 'Tidak bisa terhubung ke server. Periksa koneksi internet Anda.');
    } finally { clearTimeout(timer); }
  }
}

/* Waktu berjalan + perkiraan kemajuan pemuatan data. Waktu = nyata. Persentase = PERKIRAAN: tahap tunggu server memakai lamanya
   pemuatan sebelumnya di perangkat ini (server tidak mengirim kemajuan saat menghitung), tahap unduh memakai ukuran data terakhir. */
const PERF_KEY = 'dash_perf_v1';
const PERF_DEFAULT = { full: 8000, same: 3000, bytes: 3e6 };
const perf = {
  get() { try { return Object.assign({}, PERF_DEFAULT, JSON.parse(localStorage.getItem(PERF_KEY) || '{}')); } catch (e) { return Object.assign({}, PERF_DEFAULT); } },
  set(v) { try { localStorage.setItem(PERF_KEY, JSON.stringify(v)); } catch (e) { /* abaikan */ } },
};
const clock = ms => { const s = Math.floor(ms / 1000); return pad(Math.floor(s / 60)) + ':' + pad(s % 60); };
const secs = ms => nf1.format(ms / 1000) + ' dtk';

function startProgress(label, expectWait, alive) {
  const p = perf.get(), t0 = performance.now(), st = { phase: 'wait', loaded: 0, wait: 0, download: 0, bytes: 0 };
  let timer = 0;
  const stop = () => { clearInterval(timer); timer = 0; };
  const pct = e => st.phase === 'wait' ? Math.round(60 * (1 - Math.exp(-e / expectWait)))
    : st.phase === 'download' ? 60 + Math.round(30 * Math.min(1, st.loaded / p.bytes)) : st.phase === 'process' ? 95 : 100;
  const paint = () => {
    if (!alive()) return stop();
    const e = performance.now() - t0, n = pct(e);
    const where = st.phase === 'wait' ? (e > expectWait * 1.6 + 3000 ? 'server lebih lama dari biasanya, mohon tunggu' : 'menunggu server')
      : st.phase === 'download' ? 'mengunduh data' : 'memproses data';
    const el = $('dsync'); el.className = 'sync';
    el.replaceChildren(h('span', { class: 'spin' }), h('span', { text: label + ' (~' + n + '%) - ' + clock(e) + ' · ' + where }),
      h('span', { class: 'pbar', 'aria-hidden': 'true' }, h('i', { style: 'width:' + n + '%' })));
  };
  timer = setInterval(paint, 250); paint();
  return {
    head: () => { st.phase = 'download'; },
    bytes: n => { st.loaded = n; },
    done: i => { st.phase = 'process'; st.wait = i.wait; st.download = i.download; st.bytes = i.bytes; },
    process: () => { st.phase = 'process'; },
    stop, elapsed: () => performance.now() - t0,
    summary: () => 'menunggu server ' + secs(st.wait) + ' · unduh ' + secs(st.download) + (st.bytes ? ' (' + nf1.format(st.bytes / 1048576) + ' MB)' : '')
      + ' · proses ' + secs(Math.max(0, performance.now() - t0 - st.wait - st.download)),
  };
}

/* ---------------------------------------------------------------------
   4. Tampilan umum
   --------------------------------------------------------------------- */
const state = {
  user: null, sessionToken: '', loadSeq: 0, loading: false, lastCheck: 0,
  data: null, byCat: {}, branches: new Map(), groups: {}, clientIdCounts: new Map(),
  cat: 'a', q: '', branch: '', staf: '', sort: { label: '', dir: 1 }, page: 0, size: 50, view: [],
  sumBy: 'branch', sumSort: { key: 'n', dir: -1 }, sumQ: '', wjDirty: false,
  wilayah: null, wilayahSelected: '', wilayahDraft: [], wilayahDrafts: Object.create(null), wilayahLoading: false, wilayahLoadedAt: 0,
  rekening: null, rekeningLoading: false, rekeningLoadedAt: 0, dashboardPage: 'move',
  progress: null, progressLoading: false, progressLoadedAt: 0,
  analisaCabang: null, analisaLoading: false, analisaLoadedAt: 0, analisaFilters: { branch: '*', year: '*', month: '*', period: '', compare: '' },
  rbView: 'branch', rbSearch: { branch: '', coordinator: '' },
  rbSort: { branch: { key: 'name', dir: 1 }, coordinator: { key: 'name', dir: 1 } },
};

function showView(name, msg) {
  $('splash').hidden = name !== 'splash';
  $('login').hidden = name !== 'login';
  $('app').hidden = name !== 'app';
  if (name === 'login') { showErr(msg); $(msg ? 'pw' : 'nik').focus(); }
}
function showErr(msg) { $('err').hidden = !msg; $('err').textContent = msg || ''; }

const canSetting = u => !!u && (u.tingkat === 'ADMINISTRATOR' || u.tingkat === 'USER 0');   // menu Setting
const TABS = ['dashboard', 'set', 'up'];
function selectDashboardPage(which) {
  if (which !== 'move' && which !== 'dash' && which !== 'sum' && which !== 'progress' && which !== 'analisa') which = 'move';
  state.dashboardPage = which;
  $('p-move').hidden = which !== 'move';
  $('p-dash').hidden = which !== 'dash';
  $('p-sum').hidden = which !== 'sum';
  $('p-progress').hidden = which !== 'progress';
  $('p-analisa').hidden = which !== 'analisa';
  document.querySelectorAll('[data-dashpage]').forEach(b =>
    b.setAttribute('aria-pressed', String(b.dataset.dashpage === which)));
  if (which === 'move' && state.user && !state.rekeningLoading
      && (!state.rekening || Date.now() - state.rekeningLoadedAt > AUTO_SYNC_MS)) loadRekeningBergerak();
  if (which === 'dash' && state.user && !state.loading
      && (!state.data || Date.now() - state.lastCheck > AUTO_SYNC_MS)) loadDash();
  if (which === 'sum') {
    drawSummary();
    if (state.user && !state.data && !state.loading) loadDash();
  }
  if (which === 'progress' && state.user && !state.progressLoading
      && (!state.progress || Date.now() - state.progressLoadedAt > AUTO_SYNC_MS)) loadProgresHarian();
  if (which === 'analisa' && state.user && !state.analisaLoading
      && (!state.analisaCabang || Date.now() - state.analisaLoadedAt > AUTO_SYNC_MS)) loadAnalisaCabang();
}
function selectTab(which, focus) {
  if (which === 'up' && (!state.user || state.user.tingkat !== 'ADMINISTRATOR')) which = 'dashboard';
  if (which === 'set' && !canSetting(state.user)) which = 'dashboard';
  TABS.forEach(k => {
    const on = k === which;
    $('t-' + k).setAttribute('aria-selected', String(on));
    $('t-' + k).tabIndex = on ? 0 : -1;
    $('p-' + k).hidden = !on;
  });
  if (which === 'dashboard') selectDashboardPage(state.dashboardPage || 'move');
  if (focus) $('t-' + which).focus();
}

function paintUser(user) {
  const name = user.nama || user.nik || '';
  $('uname').textContent = name;
  $('ulevel').textContent = user.tingkat || '';
  $('uavatar').textContent = name.split(/\s+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('').toUpperCase();
  $('t-set').hidden = !canSetting(user);
  if (!canSetting(user) && !$('p-set').hidden) selectTab('dashboard');
  $('t-up').hidden = user.tingkat !== 'ADMINISTRATOR';   // hanya tampilan; penegakan sebenarnya di server
}

// Dialog konfirmasi bawaan browser (<dialog>), pengganti confirm().
function askConfirm(title, text, okLabel) {
  const dlg = $('confirm');
  $('cf-title').textContent = title; $('cf-text').textContent = text; $('cf-ok').textContent = okLabel;
  return new Promise(done => {
    dlg.addEventListener('close', () => done(dlg.returnValue === 'ok'), { once: true });
    dlg.returnValue = 'cancel'; dlg.showModal();
  });
}

/* ---------------------------------------------------------------------
   5. Masuk
   --------------------------------------------------------------------- */
function stopAutoSync() {
  if (autoSyncTimer) clearInterval(autoSyncTimer);
  if (autoSyncStartTimer) clearTimeout(autoSyncStartTimer);
  autoSyncTimer = 0;
  autoSyncStartTimer = 0;
}

// Satu penghitung untuk semua pemicu sinkron otomatis (interval dan tab kembali aktif).
// Disimpan bersama token sesi agar tetap terhitung setelah refresh halaman; login baru meresetnya.
function reserveAutoSync() {
  const current = session.get();
  if (!current || !current.token) {
    logout('Sesi tidak tersedia. Silakan login kembali.', true);
    return false;
  }
  // Administrator dikecualikan dari batas logout otomatis.
  // Sinkronisasi tetap berjalan normal, tetapi penghitung tidak dinaikkan.
  const activeUser = state.user || current.user;
  if (String(activeUser && activeUser.tingkat || '').trim().toUpperCase() === 'ADMINISTRATOR') {
    return true;
  }
  const used = Math.max(0, Math.floor(Number(current.autoSyncCount) || 0));
  if (used >= AUTO_SYNC_MAX_PER_LOGIN) {
    // Jangan kirim request sinkron ke-4. Hapus sesi lokal lalu refresh menuju halaman login.
    logout('Sesi otomatis dihentikan setelah 3 kali sinkronisasi. Silakan login kembali.', true);
    return false;
  }
  current.autoSyncCount = used + 1;
  session.set(current);
  return true;
}

function startAutoSync() {
  stopAutoSync();
  const poll = () => {
    // Polling hanya saat aplikasi aktif dan data awal sudah tersedia.
    // Jika tidak ada perubahan, server menjawab same:true dengan respons ringan.
    if (!state.user || document.hidden) return;
    if (!$('p-dashboard').hidden && state.dashboardPage === 'move') {
      if (!state.rekeningLoading) loadRekeningBergerak({ automatic: true });
    } else if (!$('p-dashboard').hidden && state.dashboardPage === 'progress') {
      if (!state.progressLoading) loadProgresHarian({ automatic: true });
    } else if (!$('p-dashboard').hidden && state.dashboardPage === 'analisa') {
      if (!state.analisaLoading) loadAnalisaCabang({ automatic: true });
    } else if (state.data && !state.loading) loadDash({ automatic: true });
  };
  // Jarak pemeriksaan tetap 2 menit; waktu pemeriksaan pertama disebar acak 0–20 detik
  // setelah interval pertama agar perangkat yang login berdekatan tidak menumpuk request-nya.
  const firstDelay = AUTO_SYNC_MS + Math.floor(Math.random() * AUTO_SYNC_START_JITTER_MS);
  autoSyncStartTimer = setTimeout(() => {
    autoSyncStartTimer = 0;
    poll();
    autoSyncTimer = setInterval(poll, AUTO_SYNC_MS);
  }, firstDelay);
}

function logout(msg, forceRefresh) {
  stopAutoSync();
  state.loadSeq++; state.loading = false; state.user = null; state.sessionToken = ''; state.data = null; state.byCat = {}; state.groups = {}; state.clientIdCounts = new Map(); state.wjDirty = false;
  session.clear();
  const cacheClear = cache.clear();
  state.wilayah = null; state.wilayahDraft = []; state.wilayahDrafts = Object.create(null); state.wilayahSelected = ''; setNote('wjmsg', '');
  state.rekening = null; state.rekeningLoadedAt = 0; state.rekeningLoading = false;
  state.progress = null; state.progressLoadedAt = 0; state.progressLoading = false;
  state.analisaCabang = null; state.analisaLoadedAt = 0; state.analisaLoading = false;
  clearDash(); resetUpload();
  showView('login', msg || '');
  // Hanya logout otomatis karena batas sinkron yang memaksa refresh penuh.
  // Refresh dilakukan setelah cache lokal selesai dibersihkan; tidak ada request logout ke server.
  if (forceRefresh) Promise.resolve(cacheClear).finally(() => window.location.reload());
}

async function submitLogin(ev) {
  ev.preventDefault();
  const nik = $('nik').value.trim(), password = $('pw').value;
  $('nik').classList.toggle('bad', !nik); $('pw').classList.toggle('bad', !password);
  if (!nik || !password) { $(nik ? 'pw' : 'nik').focus(); return showErr('Isi kolom NIK dan Password di bawah ini terlebih dahulu.'); }
  showErr('');
  const go = $('go');
  go.disabled = true;
  go.replaceChildren(h('span', { class: 'spin' }), h('span', { text: 'Memverifikasi…' }));
  const r = await api('webLogin', { nik, password });
  go.disabled = false;
  go.replaceChildren(h('span', { id: 'gotxt', text: 'Masuk' }));
  if (!r.ok) return showErr(r.error || 'Login gagal.');
  session.set({ token: r.token, expires: r.expires, user: r.user, autoSyncCount: 0 });
  $('pw').value = '';
  enterApp(r.user);
}

/* ---------------------------------------------------------------------
   6. Dashboard
   --------------------------------------------------------------------- */
function enterApp(user) {
  state.user = user;
  state.sessionToken = (session.get() || {}).token || '';
  paintUser(user);
  state.cat = 'a'; state.q = ''; state.branch = ''; state.staf = ''; state.page = 0; state.sort = { label: '', dir: 1 };
  resetUpload(); setNote('setmsg', ''); setNote('wjmsg', ''); setNote('rb-setting-msg', ''); state.wjDirty = false;
  state.rekening = null; state.rekeningLoadedAt = 0;
  state.progress = null; state.progressLoadedAt = 0; state.progressLoading = false;
  state.analisaCabang = null; state.analisaLoadedAt = 0; state.analisaLoading = false;
  state.analisaFilters = { branch: '*', year: '*', month: '*', period: '', compare: '' };
  state.dashboardPage = 'move';
  state.rbSearch = { branch: '', coordinator: '' };
  state.rbSort = { branch: { key: 'name', dir: 1 }, coordinator: { key: 'name', dir: 1 } };
  $('rb-branch-q').value = ''; $('rb-coordinator-q').value = '';
  selectRekeningTable('branch');
  selectTab('dashboard'); showView('app');
  startAutoSync();
  loadRekeningBergerak();
  loadDash(); // pertahankan preload Tool PAR dan nilai Setting seperti versi sebelumnya
}

function clearDash() {
  $('dcards').replaceChildren(); $('dkpi').hidden = true; $('dbody').hidden = true; $('dtbl').replaceChildren();
  $('dmeta').textContent = 'Memuat data…'; $('ddl').disabled = true; $('dref').disabled = false; setSync(''); setNote('dmsg', '');
}
function setSync(text, kind) {
  const e = $('dsync');
  e.className = 'sync' + (kind === 'warn' ? ' warn' : '');
  e.replaceChildren(...(kind === 'busy' ? [h('span', { class: 'spin' })] : []), text || '');
}
function skeletonCards() {
  $('dcards').replaceChildren(...CATS.map(() => h('div', { class: 'cat skel', 'aria-hidden': 'true' })));
}

// Memuat data. Cache lokal hanya dipakai setelah server mengonfirmasi bahwa versi dan ruang lingkup akses masih sama.
async function loadDash(options) {
  if (state.loading) return;
  const automatic = !!(options && options.automatic);
  const hadData = !!state.data;
  if (automatic) setSync('Memeriksa perubahan Wilayah Kerja…', 'busy');
  const seq = ++state.loadSeq, tok = (session.get() || {}).token;
  state.loading = true; $('dref').disabled = true; setNote('dmsg', '');

  // Cache lokal tidak ditampilkan sebelum server mengonfirmasi versi ruang lingkup user.
  // Ini mencegah cabang yang baru dicabut dari user tetap terlihat dari IndexedDB.
  let savedDash = null;
  if (!state.data) {
    const slow = setTimeout(() => { if (!state.data && seq === state.loadSeq) skeletonCards(); }, SLOW_MS);
    const saved = await cache.get();
    clearTimeout(slow);
    if (seq !== state.loadSeq) return;
    if (saved && saved.nik === state.user.nik && saved.data && Array.isArray(saved.data.rows)
        && typeof saved.data.wajib === 'string' && saved.data.scopeSchema === 'wilayah-v1') {
      savedDash = saved.data;
    } else skeletonCards();
  }
  // Hitung satu kali per siklus sinkron otomatis, bukan per request internal/retry.
  const countedAutoSync = !!(options && options.autoSyncCounted);
  if (automatic && !countedAutoSync && !reserveAutoSync()) {
    state.loading = false; $('dref').disabled = false;
    return;
  }
  const requestVer = state.data ? state.data.ver : (savedDash ? savedDash.ver : '');
  const pv = perf.get(), prog = startProgress(state.data ? 'Memeriksa pembaruan' : 'Memverifikasi akses dan mengambil data',
    state.data ? pv.same : pv.full, () => seq === state.loadSeq && state.loading);
  let info = null;
  const r = await api('webDashboard', { token: tok, ver: requestVer },
    { retry: true, onHead: prog.head, onProgress: prog.bytes, onDone: i => { info = i; prog.done(i); } });
  prog.stop();
  if (seq !== state.loadSeq) return;
  state.loading = false; $('dref').disabled = false;

  if (r.code === 'SESSION_REQUIRED') return logout(r.error);
  if (!r.ok || r.scopeSchema !== 'wilayah-v1' || (!r.same && typeof r.wajib !== 'string')) {
    // Fail closed: tanpa verifikasi server, jangan pertahankan data cabang pada layar.
    state.data = null; state.byCat = {}; state.groups = {}; state.branches = new Map(); state.clientIdCounts = new Map();
    await cache.clear();
    clearDash();
    return showLoadError(r.ok ? MSG_OLD_SERVER : r.error || 'Data dashboard gagal dimuat.');
  }
  state.lastCheck = Date.now();
  if (r.user) { state.user = r.user; paintUser(r.user); const s = session.get(); if (s) session.set(Object.assign(s, { user: r.user })); }
  if (info) {
    const q = perf.get();
    if (r.same) q.same = Math.max(1500, Math.round(info.wait)); else { q.full = Math.max(3000, Math.round(info.wait)); if (info.bytes) q.bytes = info.bytes; }
    perf.set(q);
  }
  if (r.same && state.data) {
    setSync(automatic
      ? 'Sinkron otomatis aktif · tidak ada perubahan (' + clock(prog.elapsed()) + ').'
      : 'Data sudah yang terbaru (dicek dalam ' + clock(prog.elapsed()) + ').');
    return;
  }
  if (r.same && savedDash && savedDash.ver === r.ver) {
    applyData(savedDash);
    setSync('Data tersimpan sesuai dengan Wilayah Kerja saat ini.');
    return;
  }
  if (r.same) return loadDashFresh(automatic ? { automatic: true, autoSyncCounted: true } : undefined);
  prog.process();
  // Detail lama dapat memuat cabang yang baru saja dicabut. Tutup sebelum data baru dipasang.
  if (hadData && $('detail').open) $('detail').close();
  applyData(r);
  if (automatic && hadData) {
    setSync('Perubahan data / Wilayah Kerja terdeteksi. Sinkronisasi otomatis berhasil.');
    toast('Perubahan Wilayah Kerja atau data berhasil disinkronkan.');
  } else {
    setSync('Data diperbarui dari server dalam ' + clock(prog.elapsed()) + ' (' + prog.summary() + ').');
  }
  const nik = state.user.nik, keep = () => { if (seq === state.loadSeq) cache.set({ nik, data: r }); };
  if (window.requestIdleCallback) window.requestIdleCallback(keep, { timeout: 2000 }); else setTimeout(keep, 200);
}

async function loadDashFresh(options) { state.data = null; state.loading = false; await cache.clear(); return loadDash(options); }

function showLoadError(msg) {
  setNote('dmsg', msg, 'bad');
  $('dmsg').append(' ', h('button', { class: 'lnk', type: 'button', onclick: () => loadDash(), text: 'Coba lagi' }));
}

// Pasang data baru ke seluruh tampilan.
function applyData(d) {
  d.rows.forEach(r => { r[F.KEY] = (r[F.BRANCH] + ' ' + r[F.LOAN] + ' ' + r[F.CLIENT] + ' ' + r[F.NAME] + ' ' + r[F.STAF]).toLowerCase(); });
  // Hitung di seluruh data yang boleh diakses user, bukan hanya kategori/halaman aktif.
  // Dengan demikian ID yang muncul di kategori atau cabang berbeda tetap ditandai.
  state.clientIdCounts = new Map();
  d.rows.forEach(r => {
    const id = clientIdKey(r[F.CLIENT]);
    if (id) state.clientIdCounts.set(id, (state.clientIdCounts.get(id) || 0) + 1);
  });
  state.data = d; state.groups = {};
  state.byCat = {};
  CATS.forEach(c => { state.byCat[c.k] = d.rows.filter(r => r[F.FLAGS].indexOf(c.k) > -1); });
  state.branches = new Map();
  d.rows.forEach(r => {
    let s = state.branches.get(r[F.BRANCH]);
    if (!s) state.branches.set(r[F.BRANCH], s = new Set());
    s.add(r[F.STAF]);
  });
  if (state.branch && !state.branches.has(state.branch)) { state.branch = ''; state.staf = ''; }

  const mg = d.margin || 0;
  const meta = (d.updated ? 'Sumber: ' + d.file + ', diunggah ' + fmtStamp(d.updated) + '. ' : fmt(d.scanned) + ' pinjaman dicek. ') + 'Penyesuaian Margin ' + nf1.format(mg) + '%.';
  $('dmeta').textContent = d.scanned ? meta : '';
  $('mcur').textContent = nf1.format(mg) + '%';
  if (document.activeElement !== $('mval')) $('mval').value = String(mg);
  paintWajib();
  $('ddl').disabled = !d.scanned;
  setNote('dmsg', d.scanned ? '' : (d.scopeMessage || 'Belum ada data. Administrator perlu mengunggah Excel terlebih dahulu.'), d.scopeMessage ? 'bad' : '');
  renderKpis(); renderCards(); fillBranchOptions(); fillSortOptions();
  $('dbody').hidden = !d.scanned;
  if (d.scanned) showCat(state.cat, true);
  if (!$('p-sum').hidden) drawSummary();
}

function renderKpis() {
  const d = state.data, box = $('dkpi');
  box.hidden = !d.scanned;
  if (!d.scanned) return;
  const item = (label, value, sub, title) => h('div', null, h('dt', { text: label }), h('dd', { title }, value, sub ? h('small', { text: sub }) : null));
  const items = [
    item('Pinjaman dicek', fmt(d.scanned)),
    item('Masuk kategori', fmt(d.rows.length), nf1.format(d.rows.length / d.scanned * 100) + '% dari yang dicek'),
  ];
  if (d.stats) {
    items.push(item('Total Balance', rpShort(d.stats.balance), '', rp(d.stats.balance)));
    items.push(item('Total Simpanan', rpShort(d.stats.savings), '', rp(d.stats.savings)));
  }
  box.replaceChildren(...items);
}

function renderCards() {
  const d = state.data;
  $('dcards').replaceChildren(...CATS.map(c => {
    const rows = state.byCat[c.k], bal = rows.reduce((a, r) => a + r[F.BALANCE], 0), share = d.scanned ? rows.length / d.scanned * 100 : 0;
    return h('button', { type: 'button', class: 'cat', 'data-k': c.k, 'aria-pressed': String(c.k === state.cat), onclick: () => showCat(c.k) },
      h('span', { class: 'cat-t', text: c.title }),
      h('b', { class: 'cat-n', text: fmt(rows.length) }),
      h('small', { title: rp(bal), text: 'Total Balance ' + rpShort(bal) }),
      h('span', { class: 'share', 'aria-hidden': 'true' }, h('i', { style: 'width:' + Math.max(share, rows.length ? 1.5 : 0).toFixed(1) + '%' })),
      h('small', { text: nf1.format(share) + '% dari pinjaman yang dicek' }),
      inc(c.k) ? h('em', { class: 'cat-w', text: '+ Simp. Wajib' }) : null);   // penanda di bawah, supaya angka kartu tetap sejajar
  }));
}

function fillBranchOptions() {
  const sel = $('dbranch');
  sel.replaceChildren(h('option', { value: '', text: 'Semua cabang' }),
    ...[...state.branches.keys()].sort(collator.compare).map(b => h('option', { value: b, text: b || '(tanpa cabang)' })));
  sel.value = state.branch; fillStafOptions();
}
function fillStafOptions() {
  const sel = $('dstaf');
  const set = new Set();
  if (state.branch) (state.branches.get(state.branch) || []).forEach(s => set.add(s));
  else state.branches.forEach(s => s.forEach(x => set.add(x)));
  sel.replaceChildren(h('option', { value: '', text: 'Semua staf' }), ...[...set].sort(collator.compare).map(s => h('option', { value: s, text: s || '(tanpa staf)' })));
  if (state.staf && !set.has(state.staf)) state.staf = '';
  sel.value = state.staf;
}
function fillSortOptions() {   // urutan untuk layar kecil, tempat judul kolom tidak tampil
  const sel = $('dsort'), c = catOf(state.cat);
  const opts = [h('option', { value: '', text: 'Urutan awal' })];
  screenCols(c).filter(x => x !== COLS.note).forEach(x => {
    opts.push(h('option', { value: x.label + '|1', text: x.label + (x.num ? ' (terkecil)' : ' (A-Z)') }));
    opts.push(h('option', { value: x.label + '|-1', text: x.label + (x.num ? ' (terbesar)' : ' (Z-A)') }));
  });
  sel.replaceChildren(...opts);
  sel.value = state.sort.label ? state.sort.label + '|' + state.sort.dir : '';
}

function showCat(k, keepFilters) {
  state.cat = k; state.page = 0;
  if (!keepFilters) { state.q = ''; $('dq').value = ''; }
  document.querySelectorAll('#dcards .cat').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.k === k)));
  const c = catOf(k);
  $('dtitle').textContent = c.title; $('dhint').textContent = c.hint();
  if (!screenCols(c).some(x => x.label === state.sort.label)) state.sort = { label: '', dir: 1 };
  fillSortOptions(); drawList();
}

// Saring + urutkan -> state.view
function computeView() {
  let list = state.byCat[state.cat] || [];
  if (state.branch) list = list.filter(r => r[F.BRANCH] === state.branch);
  if (state.staf) list = list.filter(r => r[F.STAF] === state.staf);
  const terms = state.q.toLowerCase().split(/\s+/).filter(Boolean);
  if (terms.length) list = list.filter(r => terms.every(t => r[F.KEY].indexOf(t) > -1));
  const sc = screenCols(catOf(state.cat)).find(x => x.label === state.sort.label);
  if (sc) {
    const dir = state.sort.dir, get = sc.get, k = state.cat;
    list = list.slice().sort(sc.num ? (a, b) => (get(a, k) - get(b, k)) * dir : (a, b) => collator.compare(get(a, k), get(b, k)) * dir);
  }
  state.view = list;
}

function drawList() {
  computeView();
  const c = catOf(state.cat), cols = screenCols(c), total = state.view.length, base = state.byCat[state.cat].length;
  const pages = Math.max(1, Math.ceil(total / state.size));
  state.page = Math.min(state.page, pages - 1);
  const from = state.page * state.size, slice = state.view.slice(from, from + state.size);
  const filtered = !!(state.q.trim() || state.branch || state.staf);

  // Ringkasan baris yang tampil
  let bal = 0; for (let i = 0; i < total; i++) bal += state.view[i][F.BALANCE];
  $('dvsum').replaceChildren(
    h('span', null, h('b', { text: fmt(total) }), ' pinjaman' + (filtered ? ' dari ' + fmt(base) + ' di kategori ini' : '')),
    h('span', null, 'Total Balance ', h('b', { title: rp(bal), text: rpShort(bal) })));
  $('dreset').hidden = !filtered;

  // Tabel
  const hr = h('tr');
  cols.forEach(x => {
    const on = x.label === state.sort.label;
    hr.append(h('th', { scope: 'col', class: x.num ? 'num' : '', 'aria-sort': on ? (state.sort.dir > 0 ? 'ascending' : 'descending') : null },
      x === COLS.note ? x.label : h('button', { type: 'button', class: 'th', 'data-l': x.label, text: x.label })));
  });
  const body = h('tbody');
  slice.forEach((r, i) => {
    const tr = h('tr', { 'data-i': from + i });      // klik di mana saja membuka rincian; pengguna keyboard memakai tombol pada nama
    cols.forEach(x => {
      const v = x.get(r, state.cat), td = h('td', { 'data-label': x.label });
      if (x.num) { td.className = 'num' + (v < 0 ? ' neg' : ''); td.textContent = fmt(v); }
      else if (x === COLS.note) { td.className = 'c-note'; if (v) td.append(h('span', { class: 'warn-chip', title: v, text: v })); else td.classList.add('dim'); }
      else if (x === COLS.name) { td.className = 'c-name'; td.append(h('button', { type: 'button', class: 'rowbtn', title: v.length > 28 ? v : null, text: v })); }
      else if (x === COLS.client) {
        const id = clientIdKey(v), count = id ? (state.clientIdCounts.get(id) || 0) : 0;
        td.textContent = v;
        if (count > 1) {
          td.className = 'c-client-duplicate';
          td.title = 'Duplikat Client ID: ditemukan ' + fmt(count) + ' kali pada data yang dapat Anda akses.';
          td.append(h('span', { class: 'duplicate-chip', text: 'DUPLIKAT' }));
        }
      }
      else {
        td.textContent = v;
        if (x === COLS.product) { td.className = 'c-product'; if (v.length > 28) td.title = v; }
      }
      tr.append(td);
    });
    body.append(tr);
  });
  $('dtbl').replaceChildren(h('caption', { class: 'sr-only', text: c.title }), h('thead', null, hr), body);
  $('dtw').hidden = !total; $('dtw').scrollTop = 0;

  // Kosong / halaman
  $('dempty').hidden = !!total;
  if (!total) {
    $('demptyt').textContent = filtered ? 'Tidak ada baris yang cocok' : 'Tidak ada pinjaman di kategori ini';
    $('demptyp').textContent = filtered ? 'Ubah kata kunci atau filter, atau kembalikan ke semua data.' : 'Pilih kategori lain di atas.';
    $('demptyp').append(...(filtered ? [h('br'), h('button', { class: 'btn ghost', type: 'button', onclick: resetFilters, text: 'Reset filter' })] : []));
  }
  $('dpager').hidden = !total; $('dxl').hidden = !total;
  $('dinfo').textContent = total ? 'Menampilkan ' + fmt(from + 1) + '-' + fmt(from + slice.length) + ' dari ' + fmt(total) : '';
  $('dpg').textContent = 'Halaman ' + fmt(state.page + 1) + ' dari ' + fmt(pages);
  $('dfirst').disabled = $('dprev').disabled = state.page === 0;
  $('dnext').disabled = $('dlast').disabled = state.page >= pages - 1;
}

function resetFilters() {
  state.q = ''; state.branch = ''; state.staf = ''; state.page = 0;
  $('dq').value = ''; $('dbranch').value = ''; fillStafOptions(); drawList();
}
function sortBy(label, dir) {
  const s = state.sort;
  if (dir == null) dir = s.label === label ? -s.dir : (screenCols(catOf(state.cat)).find(x => x.label === label).num ? -1 : 1);
  state.sort = { label, dir }; state.page = 0; $('dsort').value = label + '|' + dir; drawList();
}
function go(page) { state.page = page; drawList(); $('dtw').scrollIntoView({ block: 'nearest' }); }

/* ---------------------------------------------------------------------
   7. Rincian satu pinjaman
   --------------------------------------------------------------------- */
function openDetail(r) {
  const kv = (label, value, cls) => h('div', { class: cls || '' }, h('dt', { text: label }), h('dd', { text: value }));
  const li = (label, value, cls) => h('li', { class: cls || '' }, h('span', { text: label }), h('span', { text: value }));
  const k = state.cat, w = inc(k);                   // w: Simp. Wajib ikut dihitung di kategori yang sedang dibuka
  const onlyA = k === 'a';                           // Angsuran By Sukarela: tanpa rincian simpanan lain, catatan, atau perhitungan pelunasan
  const noWajib = k === 's' && !w;                   // Pelunasan dari Simpanan tanpa Wajib: baris Wajib tidak ditampilkan
  const mg = (state.data && state.data.margin) || 0;
  const val = (name, kk) => COLS[name].get(r, kk || k);   // kk: hitung memakai pilihan Wajib kategori lain
  const cats = onlyA ? [catOf('a')] : CATS.filter(c => r[F.FLAGS].indexOf(c.k) > -1);
  const note = noteText(r[F.NOTE], k);
  put($('detailbody'),
    h('button', { class: 'btn ghost sm dlg-x', type: 'button', onclick: () => $('detail').close(), text: 'Tutup' }),
    h('h2', { id: 'dd-title', text: r[F.NAME] }),
    h('p', { class: 'dlg-sub', text: r[F.BRANCH] + ', Loan No ' + r[F.LOAN] }),
    h('div', { class: 'dlg-chips' }, cats.map(c => h('span', { class: 'chip', text: c.title }))),
    h('dl', { class: 'kv' },
      kv('Product Name', r[F.PRODUCT], 'wide'), kv('Client ID', r[F.CLIENT]), kv('Kode/Staf', r[F.STAF]),
      kv('Total Balance', rp(r[F.BALANCE])), kv('Cicilan (Pokok + Margin)', rp(r[F.CICILAN]))),
    onlyA ? [
      h('h3', { text: 'Perhitungan angsuran dari Simp. Sukarela' + (w ? ' dan Simp. Wajib' : '') }),
      h('ul', { class: 'rows' },
        li('Simp. Sukarela (sudah dikurangi 2.000)', rp(r[F.SUKARELA])),
        w ? li('Simp. Wajib (sudah dikurangi 2.000)', rp(r[F.SIMP])) : null,
        li(w ? 'Jumlah angsuran ((Sukarela + Wajib) / Cicilan)' : 'Jumlah angsuran (Sukarela / Cicilan)', fmt(val('angsuran')) + ' kali', 'sum')),
    ] : [
      h('h3', { text: 'Rincian simpanan (sudah dikurangi 2.000 per jenis)' }),
      h('ul', { class: 'rows' },
        SIMPANAN.map((n, i) => noWajib && i === 0 ? null : li(n, rp(r[F.SIMP + i]))),
        li(noWajib ? 'Kum. Simpanan' : 'Kum. Simpanan (tanpa Wajib)', rp(r[F.KUM]), 'sum'),
        noWajib ? null : li('Total Simpanan', rp(r[F.TOTAL]), 'sum')),
      h('h3', { text: k === 't' ? 'Perhitungan pelunasan tanpa margin' : 'Perhitungan pelunasan' }),
      h('ul', { class: 'rows' },
        k === 't' ? [
          li('Total Balance', rp(r[F.BALANCE])),
          li('Kum. Simpanan', '- ' + rp(r[F.KUM])),
          w ? li('Simp. Wajib', '- ' + rp(r[F.SIMP])) : null,
          li(w ? 'Total Balance - Total Simpanan' : 'Total Balance - Kum. Simpanan', rp(val('sisaTm')), 'sum'),
        ] : [
          li('Total Balance', rp(r[F.BALANCE])),
          li('Penyesuaian Margin (' + nf1.format(mg) + '%)', '+ ' + rp(r[F.TBADJ] - r[F.BALANCE])),
          li('Total Balance + Margin (nilai akhir)', rp(r[F.TBADJ]), 'sum'),
          li('Kum. Simpanan', '- ' + rp(r[F.KUM])),
          w ? li('Simp. Wajib', '- ' + rp(r[F.SIMP])) : null,
          li('Sisa Pelunasan', rp(val('sisa')), 'sum'),
          k === 's' ? null : li(inc('t') ? 'Total Balance - Total Simpanan' : 'Total Balance - Kum. Simpanan', rp(val('sisaTm', 't'))),   // info: hitungan kategori Tanpa Margin
        ],
        li('Jumlah angsuran dari Simp. Sukarela' + (inc('a') ? ' + Wajib' : ''), fmt(val('angsuran', 'a')) + ' kali')),   // info: hitungan kategori Angsuran By Sukarela
      note ? h('p', null, h('span', { class: 'warn-chip', text: note })) : null,
    ],
    h('div', { class: 'dlg-act' },
      h('button', { class: 'btn ghost', type: 'button', text: 'Salin Loan No', onclick: () => copyText(r[F.LOAN]) })));
  $('detail').showModal();
}
async function copyText(text) {
  try { await navigator.clipboard.writeText(text); toast('Loan No disalin.'); }
  catch (e) { toast('Tidak bisa menyalin. Salin manual dari layar.', true); }
}

/* ---------------------------------------------------------------------
   8. Ringkasan (per cabang / per staf)
   --------------------------------------------------------------------- */
function groupsOf(kind) {
  if (state.groups[kind]) return state.groups[kind];
  const idx = kind === 'branch' ? F.BRANCH : F.STAF, map = new Map();
  state.data.rows.forEach(r => {
    const key = r[idx];
    let g = map.get(key);
    if (!g) map.set(key, g = { name: key, branch: r[F.BRANCH], n: 0, a: 0, 1: 0, 2: 0, 5: 0, s: 0, t: 0, balance: 0 });
    g.n++; g.balance += r[F.BALANCE];
    for (const ch of r[F.FLAGS]) g[ch]++;
  });
  return (state.groups[kind] = [...map.values()]);
}

function drawSummary() {
  const tbl = $('stbl');
  if (!state.data || !state.data.scanned) { tbl.replaceChildren(); return setNote('smsg', state.data ? 'Belum ada data.' : 'Data masih dimuat.'); }
  setNote('smsg', '');
  const kind = state.sumBy, q = state.sumQ.trim().toLowerCase(), sort = state.sumSort;
  $('sq').placeholder = kind === 'branch' ? 'Cari cabang…' : 'Cari staf atau cabang…';
  document.querySelectorAll('.seg button').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.g === kind)));

  let list = groupsOf(kind);
  if (q) list = list.filter(g => (g.name + ' ' + g.branch).toLowerCase().indexOf(q) > -1);
  const textKey = sort.key === 'name' || sort.key === 'branch';
  list = list.slice().sort((a, b) => textKey ? collator.compare(a[sort.key], b[sort.key]) * sort.dir : (a[sort.key] - b[sort.key]) * sort.dir);

  const heads = [['name', kind === 'branch' ? 'Cabang' : 'Kode/Staf'], ...(kind === 'staf' ? [['branch', 'Cabang']] : []), ['n', 'Pinjaman'],
    ...CATS.map(c => [c.k, c.title]), ['balance', 'Total Balance']];
  const max = {}; heads.forEach(([k]) => { max[k] = list.reduce((m, g) => Math.max(m, typeof g[k] === 'number' ? g[k] : 0), 0); });
  const textCol = k => k === 'name' || k === 'branch';

  const hr = h('tr', null, heads.map(([k, label]) => h('th', { scope: 'col', class: textCol(k) ? '' : 'num', 'aria-sort': sort.key === k ? (sort.dir > 0 ? 'ascending' : 'descending') : null },
    h('button', { type: 'button', class: 'th', 'data-s': k, text: label }))));
  const body = h('tbody', null, list.map(g => h('tr', null, heads.map(([k]) => {
    if (k === 'name') return h('td', null, h('button', { type: 'button', class: 'lnk', 'data-drill': kind, 'data-v': g.name, 'data-b': g.branch, text: g.name || '(kosong)' }));
    if (k === 'branch') return h('td', { class: 'dim', text: g.branch });
    const v = g[k], td = h('td', { class: 'num' + (k === 'balance' || k === 'n' ? '' : ' dbar'), style: '--w:' + (max[k] ? (v / max[k] * 100).toFixed(1) : 0) + '%' });
    if (k === 'balance') { td.title = rp(v); td.textContent = fmt(v); }
    else if (k === 'n' || !v) td.textContent = fmt(v);
    else td.append(h('button', { type: 'button', class: 'lnk', 'data-drill': kind, 'data-v': g.name, 'data-b': g.branch, 'data-c': k, text: fmt(v) }));
    return td;
  }))));
  const sum = k => list.reduce((a, g) => a + g[k], 0);
  const foot = h('tfoot', null, h('tr', null, heads.map(([k], i) =>
    h('td', { class: textCol(k) ? '' : 'num', text: i === 0 ? 'Total (' + fmt(list.length) + ')' : k === 'branch' ? '' : fmt(sum(k)) }))));
  tbl.replaceChildren(h('thead', null, hr), body, foot);
  if (!list.length) setNote('smsg', 'Tidak ada hasil untuk "' + state.sumQ.trim() + '".');
}

// Dari ringkasan ke daftar: pasang filter lalu buka tab Dashboard.
function drill(kind, value, branch, cat) {
  state.q = ''; $('dq').value = ''; state.page = 0;
  if (kind === 'branch') { state.branch = value; state.staf = ''; } else { state.branch = branch; state.staf = value; }
  $('dbranch').value = state.branch; fillStafOptions(); $('dstaf').value = state.staf;
  state.dashboardPage = 'dash';
  selectTab('dashboard');
  showCat(cat || state.cat, true);
  $('dtitle').scrollIntoView({ block: 'start' });
}

/* ---------------------------------------------------------------------
   9. Baca Excel (di browser; hanya ADMINISTRATOR yang memakainya)
   --------------------------------------------------------------------- */
const isoOf = (y, m, d) => { y = +y; m = +m; d = +d; return y >= 1900 && y <= 2200 && m >= 1 && m <= 12 && d >= 1 && d <= 31 ? y + '-' + pad(m) + '-' + pad(d) : ''; };
function toDate(v) {
  if (v === '' || v == null) return '';
  if (v instanceof Date) return isNaN(v) ? '' : isoOf(v.getFullYear(), v.getMonth() + 1, v.getDate());
  if (typeof v === 'number') {                            // nomor seri tanggal Excel
    if (v < 20000 || v > 80000) return '';
    const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 864e5);
    return isoOf(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
  }
  const t = String(v).trim(); let m;
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(t))) return isoOf(m[1], m[2], m[3]);
  if ((m = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/.exec(t))) return isoOf(m[3], m[2], m[1]);   // hari/bulan/tahun
  if ((m = /^(\d{1,2})\s+([A-Za-z]+)\.?\s+(\d{4})/.exec(t)) && MONTHS_EN_ID[m[2].slice(0, 3).toLowerCase()]) return isoOf(m[3], MONTHS_EN_ID[m[2].slice(0, 3).toLowerCase()], m[1]);
  return '';
}
function toNum(v) {
  if (typeof v === 'number') return isFinite(v) ? v : '';
  const t = String(v == null ? '' : v).replace(/\s/g, '');
  if (!t || t === '-') return '';
  if (/^-?\d+(\.\d{3})*(,\d+)?$/.test(t)) return +t.replace(/\./g, '').replace(',', '.');
  if (/^-?\d+(,\d{3})*(\.\d+)?$/.test(t)) return +t.replace(/,/g, '');
  return '';
}
const toTxt = v => v == null ? '' : typeof v === 'number' ? (Number.isInteger(v) ? BigInt(v).toString() : String(v)) : String(v).trim();
const blank = v => v == null || String(v).trim() === '';

// Pustaka Excel dimuat hanya saat dibutuhkan (unggah / unduh)
let xlsxLoading;
function loadXLSX() {
  if (window.XLSX) return Promise.resolve();
  return xlsxLoading || (xlsxLoading = new Promise((ok, no) => {
    const s = h('script', { src: XLSX_URL });
    s.onload = ok;
    s.onerror = () => { xlsxLoading = null; no(new Error('Pustaka pembaca Excel gagal dimuat. Periksa koneksi internet.')); };
    document.head.append(s);
  }));
}

function parseUploadFilename(filename) {
  const name = String(filename || '').trim();
  if (!/\.(xlsx|xls)$/i.test(name)) throw new Error('File wajib berformat Excel .xls atau .xlsx. File .xlsm tidak diterima.');
  const base = name.replace(/\.(xlsx|xls)$/i, '').trim();
  const hasAnalisa = /(?:^|[^A-Z0-9])ANALISA(?=[^A-Z0-9])/i.test(base);
  const hasPar = /(?:^|[^A-Z0-9])PAR(?=[^A-Z0-9])/i.test(base);
  if (hasAnalisa && hasPar) throw new Error('Nama file tidak boleh memuat kata Analisa dan PAR sekaligus. Gunakan satu jenis proses per file.');
  const re = hasAnalisa ? /(?:^|[^A-Z0-9])ANALISA(?=[^A-Z0-9]).*?(\d{2}-\d{2}-\d{4})$/i
    : /(?:^|[^A-Z0-9])PAR(?=[^A-Z0-9]).*?(\d{2}-\d{2}-\d{4})$/i;
  let m = re.exec(base);
  const type = hasAnalisa ? 'ANALISA' : 'PAR';
  if (!m) throw new Error('Nama file harus mengikuti salah satu pola: "Analisa dd-mm-yyyy.xlsx" atau "PAR [nama opsional] dd-mm-yyyy.xlsx".');
  const parts = m[1].split('-').map(Number), d = new Date(Date.UTC(parts[2], parts[1] - 1, parts[0]));
  if (d.getUTCFullYear() !== parts[2] || d.getUTCMonth() !== parts[1] - 1 || d.getUTCDate() !== parts[0]) throw new Error('Tanggal pada nama file tidak valid. Gunakan dd-mm-yyyy.');
  return { type, periodText: m[1], period: parts[2] + '-' + pad(parts[1]) + '-' + pad(parts[0]) };
}

async function readAnalisaExcel(file, info) {
  if (file.size > UP_MAX_MB * 1048576) throw new Error('Ukuran file melebihi ' + UP_MAX_MB + ' MB.');
  const buf = await file.arrayBuffer(), b = new Uint8Array(buf.slice(0, 4));
  const zip = b[0] === 0x50 && b[1] === 0x4B, ole = b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0;
  if (!(zip || ole)) throw new Error('Isi file bukan workbook Excel yang valid.');
  await loadXLSX();
  const wb = XLSX.read(buf, { type: 'array', dense: true });
  let ws = null, block = null;
  const compact = v => String(v == null ? '' : v).toLowerCase().replace(/[^a-z0-9]/g, '');
  for (const sheetName of wb.SheetNames) {
    const candidate = wb.Sheets[sheetName];
    if (!candidate['!ref']) continue;
    const data = XLSX.utils.sheet_to_json(candidate, { header: 1, range: 'A4:AH22', raw: true, defval: '' });
    const header = data[0] || [];
    if (compact(header[5]) === 'name') { ws = candidate; block = data; break; }
  }
  if (!ws || !block) throw new Error('Sheet Analisa tidak ditemukan. Header kolom F pada baris 4 harus berisi "Name".');
  const header = block[0] || [];
  const checks = [
    [5, 'F (Name)', x => x === 'name'], [7, 'H (Member)', x => x === 'member'],
    [24, 'Y (Outs. Loan)', x => x === 'outsloan'], [20, 'U (Disbursement)', x => x === 'disbursement'],
    [31, 'AF (Tot. Sav. Balance)', x => x === 'totsavbalance'], [21, 'V (Loan Princ. Rep.)', x => x.includes('loan') && (x.includes('princ') || x.includes('principal'))],
    [22, 'W (Interest Coll.)', x => x.includes('interest') && (x.includes('coll') || x.includes('collect'))],
    [32, 'AG (Memb. PAR)', x => x === 'membpar'], [33, 'AH (Outs. PAR)', x => x === 'outspar']
  ];
  for (const [idx, label, test] of checks) {
    if (!test(compact(header[idx]))) throw new Error('Header baris 4 pada kolom ' + label + ' tidak sesuai. Kolom terdeteksi: "' + String(header[idx] || '(kosong)') + '".');
  }
  const map = [5, 7, 24, 20, 31, 21, 22, 32, 33];
  const rows = [], seen = new Set(); let skipped = 0, badNumeric = 0;
  for (let i = 1; i < block.length; i++) { // baris Excel 5 sampai 22
    const r = block[i] || [];
    const values = map.map(c => r[c] == null ? '' : r[c]);
    if (values.every(blank)) { skipped++; continue; }
    const branch = toTxt(values[0]);
    if (!branch) throw new Error('Nama Cabang kosong pada baris Excel ' + (i + 4) + '.');
    const key = branch.replace(/\s+/g, ' ').trim().toLocaleUpperCase('id-ID');
    if (seen.has(key)) throw new Error('Nama Cabang "' + branch + '" tercantum lebih dari satu kali pada file.');
    seen.add(key);
    const numeric = values.slice(1).map((v, j) => {
      if (blank(v)) return '';
      const cleaned = typeof v === 'string' ? v.replace(/^Rp\.?/i, '').trim() : v;
      const n = toNum(cleaned);
      if (n === '') { badNumeric++; throw new Error('Nilai angka tidak valid pada baris Excel ' + (i + 4) + ', kolom ' + ['H','Y','U','AF','V','W','AG','AH'][j] + '.'); }
      return n;
    });
    rows.push([branch].concat(numeric));
  }
  if (!rows.length) throw new Error('Tidak ada data cabang pada baris 5 sampai 22.');
  return { name: file.name, type: 'ANALISA', period: info.period, periodText: info.periodText, rows, skipped };
}

async function readExcel(file) {
  if (file.size > UP_MAX_MB * 1048576) throw new Error('Ukuran file melebihi ' + UP_MAX_MB + ' MB.');
  const buf = await file.arrayBuffer(), b = new Uint8Array(buf.slice(0, 4));
  const zip = b[0] === 0x50 && b[1] === 0x4B, ole = b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0;
  if (!/\.(xlsx|xls)$/i.test(file.name) || !(zip || ole)) throw new Error('File harus berformat Excel (.xlsx atau .xls).');
  await loadXLSX();
  const wb = XLSX.read(buf, { type: 'array', dense: true });
  let ws = null;
  for (const n of wb.SheetNames) {                        // sheet pertama yang kolom C baris 1-4 berisi "Loan No"
    const w = wb.Sheets[n]; if (!w['!ref']) continue;
    const top = XLSX.utils.sheet_to_json(w, { header: 1, range: 'A1:AK4', raw: true, defval: '' });
    if (top.some(r => /^loan\s*no\.?$/i.test(String(r[2] || '').trim()))) { ws = w; break; }
  }
  if (!ws) throw new Error('Kolom "Loan No" tidak ditemukan di kolom C (baris 1-4). Pastikan ini file yang benar.');
  const last = XLSX.utils.decode_range(ws['!ref']).e.r + 1;
  const all = XLSX.utils.sheet_to_json(ws, { header: 1, range: 'A1:AK' + last, raw: true, defval: '' });
  const rows = []; let skipped = 0, badDate = 0, badNum = 0;
  for (let i = UP_FIRST_ROW - 1; i < all.length; i++) {
    const r = all[i];
    if (blank(r[2])) { if (r.some(x => !blank(x))) skipped++; continue; }   // tanpa Loan No = bukan baris data (total / catatan)
    rows.push(UP_COLS.map(([, c, t]) => {
      const v = r[c];
      if (t === 'd') { const x = toDate(v); if (x === '' && !blank(v)) badDate++; return x; }
      if (t === 'n') { const x = toNum(v); if (x === '' && !blank(v)) badNum++; return x; }
      return toTxt(v);
    }));
  }
  if (!rows.length) throw new Error('Tidak ada baris data (Loan No terisi) mulai baris ' + UP_FIRST_ROW + '.');
  return { name: file.name, rows, skipped, badDate, badNum };
}

/* ---------------------------------------------------------------------
   10. Unggah + pemeriksaan perubahan Loan No
   --------------------------------------------------------------------- */
let pending = null;

function upBar(f) { $('upbar').hidden = f == null; if (f != null) $('upfill').style.width = Math.round(f * 100) + '%'; }
function resetUpload() {
  pending = null; $('upinfo').hidden = true; $('file').value = ''; $('upstats').replaceChildren(); $('uptbl').replaceChildren();
  $('updiff').hidden = true; $('updiffstats').replaceChildren(); $('updiff-new-samples').textContent = 'Tidak ada.'; $('updiff-done-samples').textContent = 'Tidak ada.';
  $('upgo').textContent = 'Periksa / Simpan Excel'; $('upgo').disabled = false;
  setNote('upmsg', ''); upBar(null);
}

function renderPreview(d) {
  const chip = (text, warn) => h('span', { class: 'st' + (warn ? ' warn' : ''), text });
  put($('upstats'), chip(d.name), chip(fmt(d.rows.length) + ' baris data'),
    d.skipped ? chip(fmt(d.skipped) + ' baris dilewati (tanpa Loan No)', true) : null,
    d.badDate ? chip(fmt(d.badDate) + ' tanggal tidak terbaca (dikosongkan)', true) : null,
    d.badNum ? chip(fmt(d.badNum) + ' angka tidak terbaca (dikosongkan)', true) : null);
  const head = h('tr', null, UP_PREVIEW.map(i => h('th', { scope: 'col', class: UP_COLS[i][2] === 'n' ? 'num' : '', text: UP_COLS[i][0] })));
  const body = d.rows.slice(0, 5).map(r => h('tr', null, UP_PREVIEW.map(i => {
    const num = UP_COLS[i][2] === 'n';
    return h('td', { class: num ? 'num' : '', text: num && r[i] !== '' ? fmt(r[i]) : r[i] });
  })));
  $('uptbl').replaceChildren(h('thead', null, head), h('tbody', null, body));
}

function renderAnalisaPreview(d) {
  const chip = (text, warn) => h('span', { class: 'st' + (warn ? ' warn' : ''), text });
  put($('upstats'), chip('FITUR ANALISA'), chip(d.name), chip('Periode: ' + d.periodText), chip(fmt(d.rows.length) + ' cabang terbaca'),
    d.skipped ? chip(fmt(d.skipped) + ' baris kosong dilewati', true) : null);
  const labels = ['Nama Cabang', 'Member', 'Outstanding', 'Disbursement', 'Saving', 'Repayment', 'Margin', 'Member PAR', 'Outstanding PAR'];
  const head = h('tr', null, labels.map((v, i) => h('th', { scope: 'col', class: i ? 'num' : '', text: v })));
  const body = d.rows.slice(0, 8).map(r => h('tr', null, r.map((v, i) => h('td', { class: i ? 'num' : '', text: i && v !== '' ? fmt(v) : v }))));
  $('uptbl').replaceChildren(h('thead', null, head), h('tbody', null, body));
  $('updiff').hidden = true;
  $('upgo').textContent = 'Simpan ke Sheet Analisa';
}

function renderUploadCompare(c) {
  const chip = (text, warn) => h('span', { class: 'st' + (warn ? ' warn' : ''), text });
  put($('updiffstats'),
    chip(fmt(c.uploadCount || 0) + ' LOAN NO pada Excel'),
    chip(fmt(c.currentCount || 0) + ' LOAN NO pada Data_Upload'),
    chip(fmt(c.newCount || 0) + ' PAR Baru', (c.newCount || 0) > 0),
    chip(fmt(c.doneCount || 0) + ' PAR Selesai', (c.doneCount || 0) > 0),
    c.progressDate ? chip('Tanggal Progress: ' + c.progressDate) : null);
  $('updiff-new-samples').textContent = (c.newSamples && c.newSamples.length) ? c.newSamples.join(' · ') : 'Tidak ditemukan LOAN NO baru.';
  $('updiff-done-samples').textContent = (c.doneSamples && c.doneSamples.length) ? c.doneSamples.join(' · ') : 'Tidak ditemukan LOAN NO selesai.';
  $('updiff').hidden = false;
}

async function pickFile(file) {
  if (!file) return;
  resetUpload(); $('upinfo').hidden = false; $('upgo').disabled = true; setNote('upmsg', 'Memeriksa nama dan membaca file…');
  await nextFrame();
  try {
    const info = parseUploadFilename(file.name);
    pending = info.type === 'ANALISA' ? await readAnalisaExcel(file, info) : await readExcel(file);
    if (info.type === 'PAR') pending.type = 'PAR';
    if (pending.type === 'ANALISA') {
      renderAnalisaPreview(pending);
      setNote('upmsg', 'File dikenali sebagai Analisa. Data akan disimpan ke sheet Analisa dan tidak mengubah Data_Upload. Jika periode sama sudah ada, data periode itu akan diperbarui.');
    } else {
      renderPreview(pending);
      $('upgo').textContent = 'Periksa LOAN NO';
      setNote('upmsg', 'File dikenali sebagai PAR. Periksa pratinjau, lalu jalankan pemeriksaan LOAN NO. Data_Upload lama belum akan diubah.');
    }
    $('upgo').disabled = false;
  } catch (e) {
    pending = null; $('upstats').replaceChildren(); $('uptbl').replaceChildren();
    setNote('upmsg', e.message || 'File tidak bisa dibaca.', 'bad');
  }
}

function upFail(r, sent) {
  if (r.code === 'SESSION_REQUIRED') return logout(r.error);
  upBar(null); $('upgo').disabled = false;
  setNote('upmsg', (r.error || 'Gagal memproses data.') + (sent ? ' Staging menerima ' + fmt(sent) + ' dari ' + fmt(pending && pending.totalRows || sent) + ' baris; Data_Upload lama masih dipertahankan.' : ''), 'bad');
}

async function stageAndCheck() {
  if (!pending || !Array.isArray(pending.rows)) return;
  const total = pending.rows.length, tok = (session.get() || {}).token;
  const ok = await askConfirm('Periksa file Excel?', 'File akan disiapkan sementara dan dibandingkan berdasarkan Loan No. Data_Upload lama tidak akan berubah sebelum Anda mengonfirmasi hasil pemeriksaan.', 'Mulai pemeriksaan');
  if (!ok) return;
  $('upgo').disabled = true; upBar(0); setNote('upmsg', 'Menyiapkan area pemeriksaan…');
  let r = await api('webUploadStart', { token: tok, total, file: pending.name });
  if (!r.ok) return upFail(r, 0);
  const id = r.uploadId; let sent = 0;
  for (let seq = 0; sent < total; seq++) {
    const part = pending.rows.slice(sent, sent + UP_CHUNK); let good = false, res;
    for (let t = 0; t < 3 && !good; t++) {
      res = await api('webUploadChunk', { token: tok, uploadId: id, seq, rows: JSON.stringify(part) });
      good = res.ok;
      if (!good) { if (res.code === 'SESSION_REQUIRED') break; await sleep(1500 * (t + 1)); }
    }
    if (!good) { pending.uploadId = id; return upFail(res, sent); }
    sent += part.length; upBar(sent / total);
    setNote('upmsg', 'Menyiapkan pemeriksaan ' + fmt(sent) + ' dari ' + fmt(total) + ' baris…');
  }
  setNote('upmsg', 'Membandingkan seluruh LOAN NO dengan Data_Upload…');
  r = await api('webUploadCheck', { token: tok, uploadId: id, total }, { timeoutMs: 300000 });
  if (!r.ok) { pending.uploadId = id; return upFail(r, total); }
  pending.uploadId = id; pending.totalRows = total; pending.checked = true; pending.compare = r.compare || {};
  pending.rows = null; // file sudah aman di sheet staging; bebaskan memori browser
  renderUploadCompare(pending.compare);
  upBar(null);
  $('upgo').textContent = 'Konfirmasi & Upload ke Data_Upload'; $('upgo').disabled = false;
  setNote('upmsg', 'Pemeriksaan selesai. Data_Upload lama masih utuh dan belum diganti.', 'good');
}

async function commitUpload() {
  if (!pending || !pending.checked || !pending.uploadId) return;
  const c = pending.compare || {};
  const ok = await askConfirm('Konfirmasi hasil pemeriksaan?',
    'PAR Baru: ' + fmt(c.newCount || 0) + ' baris. PAR Selesai: ' + fmt(c.doneCount || 0) + ' baris. Tanggal Progress: ' + (c.progressDate || '-') + '. Arsip lama tetap dipertahankan dan data baru ditambahkan di bawah baris terakhir. Setelah arsip disimpan, Data_Upload akan diganti dengan ' + fmt(pending.totalRows) + ' baris dari Excel.',
    'Simpan arsip & upload');
  if (!ok) return;
  $('upgo').disabled = true; upBar(null); setNote('upmsg', 'Menyimpan PAR Baru dan PAR Selesai, lalu memperbarui Data_Upload…');
  const tok = (session.get() || {}).token;
  const r = await api('webUploadCommit', { token: tok, uploadId: pending.uploadId }, { timeoutMs: 300000 });
  if (!r.ok) {
    upBar(null); $('upgo').disabled = false;
    $('upgo').textContent = 'Ulangi Konfirmasi & Upload';
    return setNote('upmsg', r.error || 'Penyimpanan belum selesai. Tekan ulang untuk melanjutkan proses.', 'bad');
  }
  upBar(1);
  setNote('upmsg', 'Selesai. Data_Upload: ' + fmt(r.rows) + ' baris · PAR Baru: ' + fmt(r.newCount || c.newCount || 0) + ' baris · PAR Selesai: ' + fmt(r.doneCount || c.doneCount || 0) + ' baris · Tanggal Progress: ' + (r.progressDate || c.progressDate || '-') + '.', 'good');
  toast('Pemeriksaan dan unggah Excel berhasil. Dashboard dimuat ulang.');
  // Unggahan mengubah arsip dan Data_Upload; batalkan data sub-dashboard yang tersimpan di memori.
  state.progress = null; state.progressLoadedAt = 0;
  state.rekening = null; state.rekeningLoadedAt = 0;
  pending = null; $('file').value = ''; $('upgo').disabled = true; loadDash();
}

async function saveAnalisaUpload() {
  if (!pending || pending.type !== 'ANALISA' || !Array.isArray(pending.rows)) return;
  const ok = await askConfirm('Simpan data Analisa?',
    'File: ' + pending.name + '. Periode Analisa: ' + pending.periodText + '. Cabang terbaca: ' + fmt(pending.rows.length) + '. Jika periode ini sudah ada, seluruh data pada periode tersebut akan diganti dengan isi file terbaru. Periode lain tetap dipertahankan. Data_Upload tidak akan diubah.',
    'Simpan Analisa');
  if (!ok) return;
  $('upgo').disabled = true; upBar(null); setNote('upmsg', 'Memvalidasi dan menyimpan data ke sheet Analisa…');
  const r = await api('webAnalisaUpload', { token: (session.get() || {}).token, file: pending.name, period: pending.period, rows: JSON.stringify(pending.rows) }, { timeoutMs: 180000 });
  if (r.code === 'SESSION_REQUIRED') return logout(r.error);
  if (!r.ok) { $('upgo').disabled = false; return setNote('upmsg', r.error || 'Gagal menyimpan Analisa.', 'bad'); }
  const message = (r.message || 'Data Analisa berhasil disimpan.') + ' ' + fmt(r.inserted || 0) + ' baris tersimpan; ' + fmt(r.replaced || 0) + ' baris periode lama diperbarui. Data_Upload tidak berubah.';
  setNote('upmsg', message, 'good'); toast('Data Analisa berhasil disimpan.');
  pending = null; $('file').value = ''; $('upgo').disabled = true; $('upgo').textContent = 'Pilih file berikutnya';
  state.analisaCabang = null; state.analisaLoadedAt = 0;
  if (!$('p-dashboard').hidden && state.dashboardPage === 'analisa') loadAnalisaCabang();
}

async function sendUpload() {
  if (!pending) return;
  if (pending.type === 'ANALISA') return saveAnalisaUpload();
  if (pending.checked) return commitUpload();
  return stageAndCheck();
}

/* ---------------------------------------------------------------------
   11. Unduh Excel
   Susunan tiap sheet: judul -> keterangan (isi kategori, rumus, asal data) -> judul kolom berwarna (bisa disaring, tetap terlihat
   saat digulir) -> data -> baris TOTAL. Sheet "Petunjuk" menjelaskan istilah dan arti tiap kolom.
   Warna ditulis 6 digit (RRGGBB) sesuai dokumentasi xlsx-js-style; palet sama dengan style.css.
   --------------------------------------------------------------------- */
const XL_NUM = '#,##0;[Red]-#,##0;0';      // pemisah ribuan mengikuti bahasa Excel pengguna; angka minus merah
const XL_DARK = '064E3B', XL_SOFT = 'E3F5ED', XL_LINE = 'B9CFC5', XL_INK = '10261F', XL_MUTE = '55695F';
const xlSide = (style, rgb) => ({ style, color: { rgb } });
const xlBox = rgb => ({ top: xlSide('thin', rgb), bottom: xlSide('thin', rgb), left: xlSide('thin', rgb), right: xlSide('thin', rgb) });
const xlFill = rgb => ({ patternType: 'solid', fgColor: { rgb } });
const XS = {
  title: { font: { sz: 16, bold: true, color: { rgb: XL_DARK } }, alignment: { vertical: 'center' } },
  info: { font: { sz: 10, color: { rgb: XL_MUTE } }, alignment: { vertical: 'center' } },
  rule: { font: { sz: 10, bold: true, color: { rgb: XL_INK } }, alignment: { vertical: 'center' } },
  head: { font: { sz: 10, bold: true, color: { rgb: 'FFFFFF' } }, fill: xlFill(XL_DARK), border: xlBox('FFFFFF'),
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true } },
  text: { font: { sz: 10 }, alignment: { vertical: 'center' }, border: xlBox(XL_LINE) },
  wrap: { font: { sz: 10 }, alignment: { vertical: 'top', wrapText: true }, border: xlBox(XL_LINE) },
  mid: { font: { sz: 10 }, alignment: { horizontal: 'center', vertical: 'center' }, border: xlBox(XL_LINE) },
  num: { font: { sz: 10 }, alignment: { horizontal: 'right', vertical: 'center' }, border: xlBox(XL_LINE), numFmt: XL_NUM },
  dup: { font: { sz: 10, bold: true, color: { rgb: 'B42318' } }, fill: xlFill('FFD6D6'), alignment: { vertical: 'center' }, border: xlBox(XL_LINE) },
  totT: { font: { sz: 10, bold: true, color: { rgb: XL_DARK } }, fill: xlFill(XL_SOFT), alignment: { vertical: 'center' },
    border: { top: xlSide('medium', XL_DARK), bottom: xlSide('thin', XL_LINE), left: xlSide('thin', XL_LINE), right: xlSide('thin', XL_LINE) } },
  totN: { font: { sz: 10, bold: true, color: { rgb: XL_DARK } }, fill: xlFill(XL_SOFT), alignment: { horizontal: 'right', vertical: 'center' }, numFmt: XL_NUM,
    border: { top: xlSide('medium', XL_DARK), bottom: xlSide('thin', XL_LINE), left: xlSide('thin', XL_LINE), right: xlSide('thin', XL_LINE) } },
  sec: { font: { sz: 10, bold: true, color: { rgb: 'FFFFFF' } }, fill: xlFill(XL_DARK), alignment: { vertical: 'center' }, border: xlBox(XL_DARK) },
  key: { font: { sz: 10, bold: true, color: { rgb: XL_INK } }, alignment: { vertical: 'top', wrapText: true }, border: xlBox(XL_LINE) },
};

const stampNow = () => { const t = new Date(); return today() + ' ' + pad(t.getHours()) + ':' + pad(t.getMinutes()); };
// Jumlah karakter angka setelah diberi pemisah ribuan (untuk lebar kolom).
const xlDigits = v => { const a = Math.abs(Math.trunc(v)), n = a < 10 ? 1 : Math.floor(Math.log10(a)) + 1; return n + Math.floor((n - 1) / 3) + (v < 0 ? 1 : 0); };

// Rumus tiap kategori, sesuai kolom yang benar-benar ada di sheet (bergantung pada pilihan "Include: Simpanan Wajib").
function xlRule(c) {
  const w = inc(c.k);
  if (c.k === 'a') return 'Rumus: Jumlah Angsuran = (' + (w ? 'Simp. Sukarela + Simp. Wajib' : 'Simp. Sukarela') + ') ÷ Cicilan, dibulatkan ke bawah.';
  if (c.k === 't') return 'Rumus: Total Balance - Simpanan = Total Balance - ' + (w ? 'Total Simpanan' : 'Kum. Simpanan') + ' (tanpa Penyesuaian Margin).';
  return 'Rumus: Sisa Pelunasan = Total Balance + Margin - Kum. Simpanan' + (w ? ' - Simp. Wajib' : '') + '.';
}

// Keterangan umum di bagian atas sheet. rows = daftar pinjaman pada sheet itu (kosongkan untuk sheet ringkasan).
function xlSource(rows) {
  const d = state.data, u = state.user || {};
  const origin = (d.updated ? 'Data per ' + fmtStamp(d.updated) + (d.file ? ' · Sumber: ' + d.file : '') + ' · ' : '') + 'Penyesuaian Margin ' + nf1.format(d.margin || 0) + '%';
  const by = 'Diunduh ' + fmtStamp(stampNow()) + ' oleh ' + (u.nama || u.nik || '-')
    + (rows ? ' · ' + fmt(rows.length) + ' pinjaman di ' + fmt(new Set(rows.map(r => r[F.BRANCH])).size) + ' cabang' : '')
    + ' · Penjelasan kolom ada di sheet "Petunjuk"';
  return [origin, by];
}

// Keterangan filter & urutan untuk "Unduh daftar ini", supaya jelas mengapa jumlah barisnya berbeda dari kategori penuh.
function xlFilterNote() {
  const f = [], s = state.sort, sc = s.label && screenCols(catOf(state.cat)).find(x => x.label === s.label);
  if (state.branch) f.push('Cabang ' + state.branch);
  if (state.staf) f.push('Staf ' + state.staf);
  if (state.q.trim()) f.push('Kata kunci "' + state.q.trim() + '"');
  const order = sc ? 'Urutan: ' + sc.label + (sc.num ? (s.dir > 0 ? ' (terkecil dulu)' : ' (terbesar dulu)') : (s.dir > 0 ? ' (A-Z)' : ' (Z-A)')) : '';
  if (!f.length) return order || 'Tanpa filter: seluruh pinjaman di kategori ini.';
  return 'Daftar sudah disaring: ' + f.join(' · ') + (order ? ' · ' + order : '');
}

// Satu sheet dari daftar kolom + baris. meta = { title, lines: [teks | [teks, gaya]], total: tampilkan baris TOTAL }.
// Angka diberi pemisah ribuan; baris TOTAL memakai SUBTOTAL sehingga ikut berubah saat data disaring di Excel.
function sheetOf(cols, rows, k, meta) {
  const lines = (meta.lines || []).filter(Boolean).map(l => typeof l === 'string' ? [l, 'info'] : l);
  const nc = cols.length, n = rows.length;
  // Sel Client ID yang berulang diberi warna merah (sama seperti tanda DUPLIKAT di layar).
  const dupCol = cols.findIndex(c => c.label === 'Client ID');
  const isDup = r => { const id = clientIdKey(r[F.CLIENT]); return !!id && (state.clientIdCounts.get(id) || 0) > 1; };
  const dupN = dupCol < 0 ? 0 : rows.reduce((a, r) => a + (isDup(r) ? 1 : 0), 0);
  if (dupN) lines.push(['Client ID berlatar merah = duplikat (muncul lebih dari sekali di data Anda): ' + fmt(dupN) + ' baris.', 'info']);

  const hr = lines.length + 1;                  // indeks baris judul kolom (0 = baris pertama sheet)
  const last = hr + n;                          // indeks baris data terakhir
  const L = cols.map((c, j) => XLSX.utils.encode_col(j));
  // Lebar awal dari judul kolom: boleh turun ke 2 baris, tetapi kata terpanjangnya harus muat.
  const width = cols.map(c => Math.max(9, Math.min(60, Math.max(Math.ceil(c.label.length / 2), ...c.label.split(' ').map(w => w.length)) + 2)));
  const sums = cols.map(() => 0);

  const aoa = [[meta.title]].concat(lines.map(l => [l[0]]), [cols.map(c => c.label)]);
  rows.forEach(r => {
    aoa.push(cols.map((c, j) => {
      const v = c.get(r, k);
      const w = c.num ? xlDigits(v) + 2 : Math.min(c.wrap ? 60 : 44, String(v == null ? '' : v).length + 2);
      if (w > width[j]) width[j] = w;
      if (c.num) sums[j] += v;
      return v;
    }));
  });
  if (n && meta.total) {
    aoa.push([]);                               // satu baris kosong supaya baris TOTAL tidak ikut tersaring
    aoa.push(cols.map((c, j) => j === 0 ? 'TOTAL' : c.num && !c.noSum ? sums[j] : ''));
  } else if (!n) aoa.push(['Tidak ada pinjaman di kategori ini.']);

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const cell = (r, j) => ws[L[j] + (r + 1)];     // r = indeks baris (0 = baris pertama)
  ws['A1'].s = XS.title;
  lines.forEach((l, i) => { ws['A' + (i + 2)].s = XS[l[1]]; });
  cols.forEach((c, j) => { cell(hr, j).s = XS.head; });
  if (!n) ws['A' + (hr + 2)].s = XS.info;

  for (let i = 0; i < n; i++) {
    for (let j = 0; j < nc; j++) {
      const x = cell(hr + 1 + i, j);
      if (!x) continue;
      if (cols[j].num) { x.s = XS.num; x.z = XL_NUM; } else x.s = cols[j].wrap ? XS.wrap : cols[j].mid ? XS.mid : XS.text;
    }
    if (dupN && isDup(rows[i])) { const x = cell(hr + 1 + i, dupCol); if (x) x.s = XS.dup; }
  }

  if (n && meta.total) {
    const tr = last + 2;                         // indeks baris TOTAL
    cols.forEach((c, j) => {
      const x = cell(tr, j), sum = j > 0 && c.num && !c.noSum;
      x.s = sum ? XS.totN : XS.totT;
      if (sum) { x.z = XL_NUM; x.f = 'SUBTOTAL(109,' + L[j] + (hr + 2) + ':' + L[j] + (last + 1) + ')'; }
    });
  }

  const hts = [];
  hts[0] = { hpt: 26 }; hts[hr] = { hpt: 34 };
  cols.forEach((c, j) => {
    if (!c.wrap) return;
    rows.forEach((r, i) => {
      const need = Math.ceil(String(c.get(r, k)).length / (width[j] * 1.1)) * 13.5 + 3, cur = hts[hr + 1 + i];
      if (need > 17 && (!cur || cur.hpt < need)) hts[hr + 1 + i] = { hpt: need };
    });
  });
  ws['!rows'] = hts;
  ws['!cols'] = width.map(w => ({ wch: w }));
  ws['!freeze'] = { ySplit: String(hr + 1), topLeftCell: 'A' + (hr + 2), activePane: 'bottomLeft', state: 'frozen' };   // judul kolom tetap terlihat
  if (n) ws['!autofilter'] = { ref: L[0] + (hr + 1) + ':' + L[nc - 1] + (last + 1) };
  return ws;
}

// Sheet "Petunjuk": asal data, isi tiap sheet, arti kolom, dan arti warna. sheets = [[nama sheet, penjelasan]] yang ikut di file ini.
function guideSheet(sheets) {
  const d = state.data, u = state.user || {}, mg = nf1.format(d.margin || 0) + '%';
  const wajib = CATS.filter(c => inc(c.k)).map(c => c.title);
  const rows = [];
  const sec = t => rows.push([t, '', true]);
  const add = (a, b) => rows.push([a, b, false]);

  sec('Tentang data ini');
  add('Sumber data', d.updated ? 'File ' + (d.file || '-') + ', diunggah ' + fmtStamp(d.updated) + '.' : 'Dihitung langsung oleh server.');
  add('Diunduh', fmtStamp(stampNow()) + ' oleh ' + (u.nama || u.nik || '-') + (u.tingkat ? ' (' + u.tingkat + ')' : '') + '.');
  add('Cakupan', 'Hanya cabang yang boleh diakses akun yang mengunduh (' + fmt(state.branches.size) + ' cabang pada data ini).');
  add('Penyesuaian Margin', mg + '. Total Balance + Margin = Total Balance × (1 + ' + mg + '), dibulatkan ke rupiah. Dipakai pada Pelunasan < 100.000, < 200.000, < 500.000 dan Pelunasan dari Simpanan.');
  add('Cara hitung simpanan', 'Saldo tiap jenis simpanan dikurangi 2.000 lebih dulu. Simpanan yang saldonya di bawah 2.000 dihitung 0 dan ditulis di kolom Catatan.');
  add('Simp. Wajib ikut dihitung', wajib.length ? 'Hanya pada: ' + wajib.join(', ') + ' (diatur di menu Setting).' : 'Tidak pada kategori mana pun (diatur di menu Setting).');

  sec('Isi tiap sheet');
  sheets.forEach(s => add(s[0], s[1]));

  sec('Arti kolom');
  [
    ['Branch Name', 'Nama cabang.'],
    ['Client ID / Client Name', 'Nomor dan nama klien.'],
    ['Loan No', 'Nomor pinjaman.'],
    ['Product Name', 'Nama produk pinjaman.'],
    ['Kode/Staf', 'Kode staf yang memegang pinjaman tersebut.'],
    ['Total Balance', 'Total saldo pinjaman klien menurut file Excel yang diunggah.'],
    ['Total Balance + Margin', 'Total Balance ditambah Penyesuaian Margin (' + mg + ').'],
    ['Cicilan', 'Besar satu kali angsuran (Pokok + Margin).'],
    ['Simp. Sukarela, Simp. Wajib', 'Saldo simpanan setelah dikurangi 2.000. Kolom Simp. Wajib hanya ada di kategori yang menghitungnya.'],
    ['Kum. Simpanan', 'Jumlah semua simpanan kecuali Wajib (Sukarela + Pensiun + Hari Raya + Khusus + Qurban + Sipadan).'],
    ['Total Simpanan', 'Kum. Simpanan + Simp. Wajib.'],
    ['Sisa Pelunasan', 'Yang masih harus dibayar klien untuk melunasi pinjaman setelah dikurangi simpanan. Minus (angka merah) berarti simpanan sudah melebihi sisa pinjaman.'],
    ['Total Balance - Simpanan', 'Sama seperti Sisa Pelunasan, tetapi tanpa Penyesuaian Margin (kategori Pelunasan Tanpa Margin).'],
    ['Jumlah Angsuran', 'Berapa kali angsuran yang dapat ditutup dari simpanan: simpanan ÷ Cicilan, dibulatkan ke bawah.'],
    ['Keterangan', 'Tertulis YA: pinjaman memenuhi syarat kategori pada sheet tersebut.'],
    ['Catatan', 'Peringatan bila ada simpanan yang saldonya di bawah 2.000 (disebut jenis simpanannya).'],
    ['Pinjaman Masuk Kategori', 'Sheet Per Cabang: jumlah pinjaman cabang itu yang masuk minimal satu kategori.'],
  ].forEach(r => add(r[0], r[1]));

  sec('Tanda dan tips');
  add('Client ID berlatar merah', 'Client ID yang sama muncul lebih dari sekali di data akun Anda (duplikat).');
  add('Angka merah', 'Angka negatif, ditulis merah dengan tanda minus.');
  add('Baris TOTAL', 'Ada di bawah tabel. Memakai rumus SUBTOTAL, jadi ikut berubah saat Anda menyaring data lewat tanda panah di judul kolom.');
  add('Satu pinjaman, banyak sheet', 'Satu pinjaman bisa masuk lebih dari satu kategori, sehingga jumlah pinjaman antar sheet kategori tidak boleh dijumlahkan.');

  const aoa = [['Petunjuk Membaca File Ini'], ['Dashboard Operasional · ' + fmtStamp(stampNow())], []].concat(rows.map(r => [r[0], r[1]]));
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['A1'].s = XS.title; ws['A2'].s = XS.info;
  const hts = []; hts[0] = { hpt: 26 };
  rows.forEach((r, i) => {
    const x = i + 3, a = ws['A' + (x + 1)], b = ws['B' + (x + 1)];
    a.s = r[2] ? XS.sec : XS.key; b.s = r[2] ? XS.sec : XS.wrap;
    const lines = Math.max(Math.ceil(r[1].length / 105), Math.ceil(r[0].length / 31));
    hts[x] = { hpt: r[2] ? 20 : Math.max(17, lines * 13.5 + 4) };
  });
  ws['!rows'] = hts;
  ws['!cols'] = [{ wch: 30 }, { wch: 105 }];
  return ws;
}

async function download(button, work, name) {
  button.disabled = true; setNote('dmsg', 'Menyiapkan file Excel…');
  try {
    await Promise.all([loadXLSX(), nextFrame()]);
    const wb = XLSX.utils.book_new();
    work(wb);
    XLSX.writeFile(wb, name);
    setNote('dmsg', 'File Excel berhasil dibuat.', 'good');
  } catch (e) { setNote('dmsg', e.message || 'Gagal membuat file Excel.', 'bad'); }
  button.disabled = false;
}

// Penjelasan satu sheet kategori untuk sheet Petunjuk.
const xlCatNote = (c, n) => c.title + ': ' + c.hint() + ' (' + fmt(n) + ' pinjaman; rumusnya tertulis di bagian atas sheet).';

function exportAll() {
  if (!state.data) return;
  download($('ddl'), wb => {
    const d = state.data, add = (ws, name) => XLSX.utils.book_append_sheet(wb, ws, name);
    const sumCols = [
      { label: 'Kategori', get: c => c.title },
      { label: 'Jumlah Pinjaman', get: c => state.byCat[c.k].length, num: true, noSum: true },
      { label: 'Total Balance', get: c => state.byCat[c.k].reduce((a, r) => a + r[F.BALANCE], 0), num: true },
      { label: 'Simp. Wajib Dihitung', get: c => inc(c.k) ? 'YA' : 'TIDAK', mid: true },
      { label: 'Syarat Masuk Kategori', get: c => c.hint(), wrap: true },
    ];
    const kpi = 'Pinjaman dicek: ' + fmt(d.scanned) + ' · Masuk minimal satu kategori: ' + fmt(d.rows.length)
      + (d.scanned ? ' (' + nf1.format(d.rows.length / d.scanned * 100) + '%)' : '')
      + (d.stats ? ' · Total Balance semua pinjaman dicek: ' + rp(d.stats.balance) + ' · Total Simpanan: ' + rp(d.stats.savings) : '');
    add(sheetOf(sumCols, CATS, null, { title: 'Ringkasan per Kategori', total: false, lines: [
      'Jumlah pinjaman dan Total Balance tiap kategori. Satu pinjaman bisa masuk lebih dari satu kategori, jadi angka antar kategori tidak dijumlahkan.',
      ...xlSource(), [kpi, 'rule']] }), 'Ringkasan Kategori');

    const bg = groupsOf('branch').slice().sort((a, b) => collator.compare(a.name, b.name));
    add(sheetOf([{ label: 'Cabang', get: g => g.name }, { label: 'Pinjaman Masuk Kategori', get: g => g.n, num: true },
      ...CATS.map(c => ({ label: c.title, get: g => g[c.k], num: true })), { label: 'Total Balance', get: g => g.balance, num: true }], bg, null,
      { title: 'Ringkasan per Cabang', total: true, lines: [
        'Jumlah pinjaman tiap cabang pada setiap kategori. Satu pinjaman dapat masuk lebih dari satu kategori, jadi angka antar kategori tidak untuk dijumlahkan.',
        ...xlSource()] }), 'Per Cabang');

    CATS.forEach(c => {
      const rows = state.byCat[c.k];
      add(sheetOf(excelCols(c), rows, c.k, { title: c.title, total: true, lines: [c.hint(), [xlRule(c), 'rule'], ...xlSource(rows)] }), c.sheet);
    });
    add(guideSheet([
      ['Ringkasan Kategori', 'Jumlah pinjaman dan Total Balance tiap kategori, beserta syarat masuknya.'],
      ['Per Cabang', 'Jumlah pinjaman tiap cabang pada setiap kategori.'],
      ...CATS.map(c => [c.sheet, xlCatNote(c, state.byCat[c.k].length)]),
      ['Petunjuk', 'Halaman ini.'],
    ]), 'Petunjuk');
  }, 'Dashboard_' + today() + '.xlsx');
}

function exportView() {   // kategori yang sedang dibuka, lengkap dengan filter & urutan saat ini
  if (!state.data || !state.view.length) return;
  const c = catOf(state.cat), rows = state.view;
  download($('dxl'), wb => {
    XLSX.utils.book_append_sheet(wb, sheetOf(excelCols(c), rows, c.k, { title: c.title, total: true,
      lines: [c.hint(), [xlRule(c), 'rule'], ...xlSource(rows), xlFilterNote()] }), c.sheet);
    XLSX.utils.book_append_sheet(wb, guideSheet([[c.sheet, xlCatNote(c, rows.length)], ['Petunjuk', 'Halaman ini.']]), 'Petunjuk');
  }, 'Dashboard_' + c.sheet.replace(/\s+/g, '_') + '_' + today() + '.xlsx');
}

/* ---------------------------------------------------------------------
   Dashboard: Rekening Bergerak
   --------------------------------------------------------------------- */
function rbDateLabel(value) {
  return value ? fmtStamp(value) : 'Belum diatur';
}

function rbSortValue(row, key, kind) {
  if (key === 'name') return kind === 'branch' ? (row.branch || '') : (row.name || '');
  return Number(row[key]) || 0;
}

function rbRenderTable(id, rows, firstLabel, year, kind) {
  const table = $(id);
  const headers = [
    { key: 'name', label: firstLabel, text: true },
    { key: 'total', label: 'Total Rekening PAR' },
    { key: 'active', label: 'Rekening Bergerak Tahun ' + (year || '…') },
    { key: 'deviation', label: 'Deviasi' },
    { key: 'percent', label: 'Persentasi (%)' },
  ];
  const query = (state.rbSearch[kind] || '').trim().toLocaleLowerCase('id-ID');
  const sort = state.rbSort[kind];
  let visible = (rows || []).filter(r => !query || String(kind === 'branch' ? r.branch : r.name || '').toLocaleLowerCase('id-ID').includes(query));
  visible = visible.slice().sort((a, b) => {
    const av = rbSortValue(a, sort.key, kind), bv = rbSortValue(b, sort.key, kind);
    const cmp = typeof av === 'string' ? collator.compare(av, bv) : av - bv;
    return cmp * sort.dir;
  });
  const totals = visible.reduce((a, r) => {
    a.total += Number(r.total) || 0;
    a.active += Number(r.active) || 0;
    return a;
  }, { total: 0, active: 0 });
  const totalDeviation = totals.active - totals.total;
  const totalPercent = totals.total ? Math.round(totals.active / totals.total * 1000) / 10 : 0;
  const thead = h('thead', null, h('tr', null, headers.map((item, i) => {
    const activeSort = sort.key === item.key;
    return h('th', { scope: 'col', class: i ? 'num' : '', 'aria-sort': activeSort ? (sort.dir > 0 ? 'ascending' : 'descending') : 'none' },
      h('button', { type: 'button', class: 'th', 'data-rb-sort': item.key, 'data-rb-kind': kind,
        'aria-label': 'Urutkan berdasarkan ' + item.label, text: item.label }));
  })));
  const tbody = h('tbody');
  if (!visible.length) {
    tbody.append(h('tr', { class: 'rb-empty-row' }, h('td', { colspan: '5', text: query ? 'Tidak ada hasil yang cocok dengan pencarian.' : 'Tidak ada data pada cakupan Wilayah Kerja ini.' })));
  } else {
    visible.forEach(r => tbody.append(h('tr', null,
      h('td', { text: kind === 'branch' ? (r.branch || '(tanpa nama)') : (r.name || '(tanpa nama)') }),
      h('td', { class: 'num', text: fmt(r.total) }),
      h('td', { class: 'num', text: fmt(r.active) }),
      h('td', { class: 'num rb-deviation', text: fmt(r.deviation) }),
      h('td', { class: 'num', text: nf1.format(r.percent) + '%' })
    )));
  }
  const tfoot = h('tfoot', null, h('tr', null,
    h('td', { text: query ? 'TOTAL HASIL FILTER' : 'TOTAL' }),
    h('td', { class: 'num', text: fmt(totals.total) }),
    h('td', { class: 'num', text: fmt(totals.active) }),
    h('td', { class: 'num rb-deviation', text: fmt(totalDeviation) }),
    h('td', { class: 'num', text: nf1.format(totalPercent) + '%' })
  ));
  put(table, thead, tbody, tfoot);
  const count = $(kind === 'branch' ? 'rb-branch-count' : 'rb-coordinator-count');
  if (count) count.textContent = 'Menampilkan ' + fmt(visible.length) + ' dari ' + fmt((rows || []).length) + (kind === 'branch' ? ' cabang' : ' Branch Coordinator');
}

function refreshRBTable(kind) {
  const data = state.rekening;
  if (!data || data.settingRequired) return;
  if (kind === 'branch') rbRenderTable('rb-branch-tbl', data.branchRows || [], 'Cabang', data.year, 'branch');
  else rbRenderTable('rb-coordinator-tbl', data.coordinatorRows || [], 'Branch Coordinator (Nama)', data.year, 'coordinator');
}

function selectRekeningTable(which) {
  if (which !== 'branch' && which !== 'coordinator') which = 'branch';
  state.rbView = which;
  $('rb-panel-branch').hidden = which !== 'branch';
  $('rb-panel-coordinator').hidden = which !== 'coordinator';
  document.querySelectorAll('[data-rb-view]').forEach(b => b.setAttribute('aria-selected', String(b.dataset.rbView === which)));
}

function paintRBKpis(data) {
  const branchRows = data.branchRows || [];
  const totals = branchRows.reduce((a, r) => {
    a.total += Number(r.total) || 0;
    a.active += Number(r.active) || 0;
    return a;
  }, { total: 0, active: 0 });
  $('rb-kpi-total').textContent = fmt(data.totalRows == null ? totals.total : data.totalRows);
  $('rb-kpi-active').textContent = fmt(totals.active);
  $('rb-kpi-percent').textContent = (totals.total ? nf1.format(totals.active / totals.total * 100) : '0') + '%';
  $('rb-kpi-missing').textContent = fmt(data.missingDateCount || 0);
  $('rb-period-chip').textContent = 'Tahun acuan ' + (data.year || '—');
}

function paintRekeningBergerak(data) {
  const isRequired = !!data.settingRequired;
  $('rb-setting-current').textContent = rbDateLabel(data.settingDate);
  if (document.activeElement !== $('rb-setting-date')) $('rb-setting-date').value = data.settingDate || '';
  if (isRequired) {
    $('rb-kpis').hidden = true;
    $('rb-period-chip').textContent = 'Tanggal N belum diatur';
    $('rb-meta').textContent = 'Tanggal N belum ditentukan. Penghitungan Rekening Bergerak belum dapat ditampilkan.';
    $('rb-config-note').hidden = false;
    $('rb-config-note').textContent = canSetting(state.user)
      ? 'Atur tanggal N melalui menu Setting → Fitur: Rekening PAR Bergerak, kemudian muat ulang dashboard.'
      : 'Administrator atau User 0 belum menetapkan tanggal N. Hubungi administrator untuk mengaktifkan dashboard ini.';
    $('rb-warning').hidden = true;
    put($('rb-branch-tbl')); put($('rb-coordinator-tbl'));
    $('rb-branch-count').textContent = 'Menunggu tanggal N';
    $('rb-coordinator-count').textContent = 'Menunggu tanggal N';
    $('rb-xlsx').disabled = true;
    return;
  }
  $('rb-config-note').hidden = true;
  $('rb-kpis').hidden = false;
  paintRBKpis(data);
  $('rb-meta').textContent = 'Tanggal N: ' + rbDateLabel(data.settingDate) + ' · Tahun acuan: ' + data.year
    + ' · Total rekening pada cakupan Anda: ' + fmt(data.totalRows || 0);
  refreshRBTable('branch');
  refreshRBTable('coordinator');
  const warnings = [];
  if (data.missingDateCount) warnings.push(fmt(data.missingDateCount) + ' baris memiliki Trans Pinj Terakhir kosong. Data tetap dihitung ke Total Rekening PAR, tetapi tidak dihitung sebagai Rekening Bergerak.');
  if (data.invalidDateCount) warnings.push(fmt(data.invalidDateCount) + ' baris memiliki Trans Pinj Terakhir dengan format tanggal tidak valid; data tetap dihitung ke Total Rekening PAR, tetapi tidak dihitung sebagai Rekening Bergerak.');
  if (data.unassignedBranchCount) warnings.push(fmt(data.unassignedBranchCount) + ' baris tidak memiliki Branch Name sehingga tidak dapat dimasukkan ke rekap per cabang.');
  if (data.scopeMessage) warnings.push(data.scopeMessage);
  $('rb-warning').hidden = !warnings.length;
  $('rb-warning').textContent = warnings.join(' ');
  $('rb-xlsx').disabled = false;
}

async function loadRekeningBergerak(options) {
  if (state.rekeningLoading || !state.user) return;
  state.rekeningLoading = true;
  $('rb-ref').disabled = true;
  $('rb-xlsx').disabled = true;
  if (!(options && options.automatic)) $('rb-sync').textContent = 'Mengambil rekap dari Data_Upload…';
  setNote('rb-msg', '');
  try {
    if (options && options.automatic && !reserveAutoSync()) return;
    const r = await api('webRekeningBergerak', { token: (session.get() || {}).token }, { retry: true });
    if (r.code === 'SESSION_REQUIRED') return logout(r.error);
    if (!r.ok) {
      state.rekening = null;
      $('rb-meta').textContent = 'Data Rekening Bergerak gagal dimuat.';
      $('rb-kpis').hidden = true;
      $('rb-config-note').hidden = false;
      $('rb-config-note').textContent = r.error || 'Terjadi kesalahan saat mengambil data.';
      put($('rb-branch-tbl')); put($('rb-coordinator-tbl'));
      $('rb-branch-count').textContent = 'Data tidak tersedia';
      $('rb-coordinator-count').textContent = 'Data tidak tersedia';
      $('rb-warning').hidden = true;
      put($('rb-branch-tbl')); put($('rb-coordinator-tbl'));
      return;
    }
    state.rekening = r;
    state.rekeningLoadedAt = Date.now();
    paintRekeningBergerak(r);
    $('rb-sync').textContent = 'Terakhir diperbarui: ' + new Date(state.rekeningLoadedAt).toLocaleString('id-ID');
  } catch (err) {
    state.rekening = null;
    $('rb-meta').textContent = 'Data Rekening Bergerak gagal dimuat.';
    $('rb-kpis').hidden = true;
    $('rb-config-note').hidden = false;
    $('rb-config-note').textContent = err && err.message ? err.message : 'Jaringan tidak dapat dihubungi. Periksa koneksi lalu muat ulang.';
    put($('rb-branch-tbl')); put($('rb-coordinator-tbl'));
    $('rb-branch-count').textContent = 'Data tidak tersedia';
    $('rb-coordinator-count').textContent = 'Data tidak tersedia';
    $('rb-xlsx').disabled = true;
  } finally {
    state.rekeningLoading = false;
    $('rb-ref').disabled = false;
  }
}

async function saveRekeningBergerakSetting() {
  const value = $('rb-setting-date').value;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    $('rb-setting-date').classList.add('bad');
    return setNote('rb-setting-msg', 'Pilih tanggal N yang valid.', 'bad');
  }
  const ok = await askConfirm('Simpan tanggal N?', 'Rekening dengan Trans Pinj Terakhir pada atau setelah ' + rbDateLabel(value)
    + ' akan dihitung sebagai Rekening Bergerak. Pengaturan ini berlaku untuk semua pengguna.', 'Simpan');
  if (!ok) return;
  $('rb-setting-save').disabled = true;
  setNote('rb-setting-msg', 'Menyimpan tanggal N…');
  try {
    const r = await api('webSetRekeningBergerak', { token: (session.get() || {}).token, value: value });
    if (r.code === 'SESSION_REQUIRED') return logout(r.error);
    if (!r.ok) return setNote('rb-setting-msg', r.error || 'Gagal menyimpan tanggal N.', 'bad');
    $('rb-setting-date').value = r.dateN;
    $('rb-setting-current').textContent = rbDateLabel(r.dateN);
    setNote('rb-setting-msg', 'Tanggal N tersimpan: ' + rbDateLabel(r.dateN) + '.', 'good');
    state.rekening = null;
    await loadRekeningBergerak();
  } catch (err) {
    setNote('rb-setting-msg', err && err.message ? err.message : 'Gagal menyimpan tanggal N.', 'bad');
  } finally {
    $('rb-setting-save').disabled = false;
  }
}

function rbExcelSheet(rows, title, firstLabel, year, data) {
  const header = [firstLabel, 'Total Rekening PAR', 'Rekening Bergerak Tahun ' + year, 'Deviasi', 'Persentasi (%)'];
  const total = rows.reduce((a, r) => {
    a.total += Number(r.total) || 0;
    a.active += Number(r.active) || 0;
    return a;
  }, { total: 0, active: 0 });
  const percent = total.total ? Math.round(total.active / total.total * 1000) / 10 : 0;
  const aoa = [[title], ['Tanggal N: ' + rbDateLabel(data.settingDate)], [], header];
  rows.forEach(r => aoa.push([r.branch || r.name, r.total, r.active, r.deviation, r.percent]));
  aoa.push(['TOTAL', total.total, total.active, total.active - total.total, percent]);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } }];
  ws['!cols'] = [{ wch: 38 }, { wch: 22 }, { wch: 30 }, { wch: 16 }, { wch: 18 }];
  ws['!rows'] = [{ hpt: 28 }, { hpt: 22 }, { hpt: 8 }, { hpt: 34 }];
  const style = (fill, color, bold, align) => ({
    font: { name: 'Aptos', sz: 10, bold: !!bold, color: { rgb: color || '18352B' } },
    fill: { patternType: 'solid', fgColor: { rgb: fill } },
    alignment: { vertical: 'center', horizontal: align || 'left', wrapText: true },
    border: { bottom: { style: 'thin', color: { rgb: 'D5E5DC' } } }
  });
  if (ws.A1) ws.A1.s = { font: { name: 'Aptos Display', sz: 16, bold: true, color: { rgb: 'FFFFFF' } }, fill: { patternType: 'solid', fgColor: { rgb: '064E3B' } }, alignment: { vertical: 'center', horizontal: 'left' } };
  if (ws.A2) ws.A2.s = { font: { name: 'Aptos', sz: 10, color: { rgb: '405D50' }, italic: true }, alignment: { vertical: 'center' } };
  for (let c = 0; c < 5; c++) {
    const cell = XLSX.utils.encode_cell({ r: 3, c });
    if (ws[cell]) ws[cell].s = style('0A7A5A', 'FFFFFF', true, c ? 'center' : 'left');
  }
  for (let r = 0; r < rows.length; r++) {
    const rowIndex = 4 + r;
    ws['!rows'][rowIndex] = { hpt: 20 };
    for (let c = 0; c < 5; c++) {
      const cell = XLSX.utils.encode_cell({ r: rowIndex, c });
      if (!ws[cell]) continue;
      ws[cell].s = style(r % 2 ? 'F5FAF7' : 'FFFFFF', '18352B', false, c ? 'right' : 'left');
      if (c > 0) ws[cell].z = c === 4 ? '0.0"%"' : '#,##0;[Red]-#,##0';
    }
  }
  const totalRow = 4 + rows.length;
  ws['!rows'][totalRow] = { hpt: 24 };
  for (let c = 0; c < 5; c++) {
    const cell = XLSX.utils.encode_cell({ r: totalRow, c });
    if (!ws[cell]) continue;
    ws[cell].s = style('E3F5ED', '064E3B', true, c ? 'right' : 'left');
    if (c > 0) ws[cell].z = c === 4 ? '0.0"%"' : '#,##0;[Red]-#,##0';
  }
  ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 3, c: 0 }, e: { r: 3 + rows.length, c: 4 } }) };
  return ws;
}

async function exportRekeningBergerak() {
  const data = state.rekening;
  if (!data || data.settingRequired || !data.settingDate) return;
  $('rb-xlsx').disabled = true;
  setNote('rb-msg', 'Menyiapkan file Excel…');
  try {
    await Promise.all([loadXLSX(), nextFrame()]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, rbExcelSheet(data.branchRows || [], 'REKENING BERGERAK PER CABANG', 'Cabang', data.year, data), 'Per Cabang');
    XLSX.utils.book_append_sheet(wb, rbExcelSheet(data.coordinatorRows || [], 'REKENING BERGERAK PER BRANCH COORDINATOR', 'Branch Coordinator (Nama)', data.year, data), 'Per BC');
    XLSX.writeFile(wb, 'Rekening_Bergerak_' + data.settingDate + '.xlsx');
    setNote('rb-msg', 'File Excel berhasil dibuat dengan dua sheet dan baris total.', 'good');
  } catch (err) {
    setNote('rb-msg', err && err.message ? err.message : 'Gagal membuat file Excel.', 'bad');
  } finally {
    $('rb-xlsx').disabled = false;
  }
}

/* ---------------------------------------------------------------------
   Dashboard: Progres Harian
   Data dari PAR Baru/PAR Selesai. Penghitungan periode dan cakupan cabang
   dilakukan di server; browser hanya menyajikan ringkasan dan ekspor.
   --------------------------------------------------------------------- */
function progresChartCompact(value, money) {
  const n = Number(value) || 0;
  if (money) {
    const a = Math.abs(n), sign = n < 0 ? '-' : '';
    if (a >= 1e9) return sign + nf1.format(a / 1e9) + ' M';
    if (a >= 1e6) return sign + nf1.format(a / 1e6) + ' jt';
    if (a >= 1e3) return sign + nf1.format(a / 1e3) + ' rb';
    return sign + nf1.format(a);
  }
  return fmt(n);
}

function progresDayChart(hostId, daily, firstKey, secondKey, money) {
  const host = $(hostId);
  const data = Array.isArray(daily) ? daily : [];
  if (!data.length) { put(host, h('div', { class: 'ph-chart-empty', text: 'Belum ada data pada bulan berjalan.' })); return; }
  const ns = 'http://www.w3.org/2000/svg';
  const width = Math.max(620, data.length * 25 + 82), height = 250;
  const left = 72, right = 18, top = 16, bottom = 42, plotW = width - left - right, plotH = height - top - bottom;
  const maxValue = Math.max(0, ...data.map(d => Math.max(Number(d[firstKey]) || 0, Number(d[secondKey]) || 0)));
  const scaleMax = maxValue > 0 ? maxValue * 1.12 : 1;
  const step = plotW / data.length, barW = Math.max(3, Math.min(10, step * 0.31));
  const svg = document.createElementNS(ns, 'svg');
  svg.setAttribute('viewBox', '0 0 ' + width + ' ' + height);
  svg.setAttribute('width', width); svg.setAttribute('height', height);
  svg.setAttribute('role', 'img');
  svg.setAttribute('aria-label', money ? 'Grafik Out PAR baru dibanding PAR selesai per hari' : 'Grafik jumlah Client PAR baru dibanding selesai per hari');
  const add = (tag, attrs, text) => {
    const e = document.createElementNS(ns, tag);
    Object.keys(attrs || {}).forEach(k => e.setAttribute(k, attrs[k]));
    if (text != null) e.textContent = text;
    svg.appendChild(e); return e;
  };
  for (let t = 0; t <= 4; t++) {
    const y = top + plotH - (plotH * t / 4), value = scaleMax * t / 4;
    add('line', { x1: left, y1: y, x2: width - right, y2: y, stroke: '#e5eee9', 'stroke-width': 1 });
    add('text', { x: left - 9, y: y + 4, fill: '#66796f', 'font-size': 10, 'text-anchor': 'end' }, progresChartCompact(value, money));
  }
  add('line', { x1: left, y1: top + plotH, x2: width - right, y2: top + plotH, stroke: '#9bb5a8', 'stroke-width': 1.2 });
  data.forEach((d, i) => {
    const cx = left + i * step + step / 2;
    const vals = [Number(d[firstKey]) || 0, Number(d[secondKey]) || 0];
    const colors = ['#0a7a5a', '#f0c419'];
    vals.forEach((v, j) => {
      const bh = Math.max(0, v / scaleMax * plotH);
      if (bh > 0) {
        const rect = add('rect', { x: cx + (j === 0 ? -barW - 1 : 1), y: top + plotH - bh, width: barW, height: bh,
          rx: 2, fill: colors[j], opacity: .94 });
        const title = document.createElementNS(ns, 'title');
        title.textContent = d.date + ' · ' + (j === 0 ? 'PAR Baru' : 'PAR Selesai') + ': ' + (money ? rp(v) : fmt(v) + ' Loan No');
        rect.appendChild(title);
      }
    });
    const day = Number(String(d.date).slice(8, 10));
    if (day === 1 || day % 3 === 1 || i === data.length - 1) {
      add('text', { x: cx, y: height - 17, fill: '#66796f', 'font-size': 10, 'text-anchor': 'middle' }, String(day));
    }
  });
  const scroller = h('div', { class: 'ph-chart-scroll' });
  scroller.append(svg); put(host, scroller);
}

function progresRenderRanking(hostId, rows, kind) {
  const host = $(hostId);
  if (!rows || !rows.length) {
    put(host, h('div', { class: 'ph-empty-row', text: 'Belum ada progres untuk bulan berjalan.' })); return;
  }
  put(host, ...rows.map((r, i) => {
    const amount = kind === 'new' ? r.newBalance : r.doneBalance;
    const count = kind === 'new' ? r.newCount : r.doneCount;
    return h('article', { class: 'ph-rank-row' },
      h('span', { class: 'ph-rank-number', text: String(i + 1).padStart(2, '0') }),
      h('div', { class: 'ph-rank-main' }, h('b', { text: r.branch }), h('span', { text: fmt(count) + ' Client PAR (Loan No)' })),
      h('div', { class: 'ph-rank-value' }, h('b', { text: rpShort(amount), title: rp(amount) }), h('small', { text: 'Out. PAR' })));
  }));
}

function progresRenderTopUp(data) {
  const table = $('ph-topup-table');
  const head = ['Kategori TOP UP', 'PAR Baru · Client', 'PAR Baru · Out. PAR', 'PAR Selesai · Client', 'PAR Selesai · Out. PAR'];
  const thead = h('thead', {}, h('tr', {}, ...head.map(x => h('th', { text: x }))));
  const tbody = h('tbody', {});
  (data.topUp || []).forEach(r => {
    tbody.append(h('tr', {},
      h('th', { scope: 'row', text: r.category }),
      h('td', { text: fmt(r.newCount) }), h('td', { text: rp(r.newBalance) }),
      h('td', { text: fmt(r.doneCount) }), h('td', { text: rp(r.doneBalance) })));
  });
  const total = (data.topUp || []).reduce((a, r) => {
    a.newCount += Number(r.newCount) || 0; a.newBalance += Number(r.newBalance) || 0;
    a.doneCount += Number(r.doneCount) || 0; a.doneBalance += Number(r.doneBalance) || 0; return a;
  }, { newCount: 0, newBalance: 0, doneCount: 0, doneBalance: 0 });
  const tfoot = h('tfoot', {}, h('tr', {}, h('th', { text: 'TOTAL 3 KATEGORI' }),
    h('td', { text: fmt(total.newCount) }), h('td', { text: rp(total.newBalance) }),
    h('td', { text: fmt(total.doneCount) }), h('td', { text: rp(total.doneBalance) })));
  put(table, thead, tbody, tfoot);
  const unknown = data.unknownTopUp || { newCount: 0, doneCount: 0 };
  $('ph-topup-note').textContent = (unknown.newCount || unknown.doneCount)
    ? 'Ada nilai Top Up di luar tiga kategori yang diminta: ' + fmt(unknown.newCount) + ' baris PAR Baru dan ' + fmt(unknown.doneCount) + ' baris PAR Selesai. Nilai tersebut tetap masuk kartu total dan grafik, tetapi tidak dimasukkan ke tabel tiga kategori.'
    : 'Seluruh data dengan Loan No pada bulan berjalan memiliki kategori TOP UP yang sesuai dengan tiga kategori di atas.';
}

function paintProgresHarian(data) {
  $('ph-meta').textContent = data.monthLabel + ' · periode ' + fmtStamp(data.monthKey + '-01') + ' s.d. ' + fmtStamp(data.asOf)
    + ' · sumber: PAR Baru dan PAR Selesai';
  $('ph-scope').textContent = data.scope || 'Sesuai Wilayah Kerja';
  $('ph-new-balance').textContent = rpShort(data.totals.newBalance);
  $('ph-new-balance').title = rp(data.totals.newBalance);
  $('ph-new-count').textContent = fmt(data.totals.newCount);
  $('ph-done-balance').textContent = rpShort(data.totals.doneBalance);
  $('ph-done-balance').title = rp(data.totals.doneBalance);
  $('ph-done-count').textContent = fmt(data.totals.doneCount);
  const netBalance = Number(data.netBalance) || 0, netCount = Number(data.netCount) || 0;
  $('ph-net-balance').textContent = (netBalance > 0 ? '+' : '') + rpShort(netBalance);
  $('ph-net-balance').title = rp(netBalance);
  $('ph-net-count').textContent = (netCount > 0 ? '+' : '') + fmt(netCount);
  $('ph-net-balance').className = netBalance > 0 ? 'negative' : netBalance < 0 ? 'positive' : '';
  $('ph-net-count').className = netCount > 0 ? 'negative' : netCount < 0 ? 'positive' : '';
  $('ph-net-balance-note').textContent = netBalance > 0 ? 'Kenaikan lebih besar daripada penyelesaian' : netBalance < 0 ? 'Penyelesaian lebih besar daripada kenaikan' : 'Kenaikan dan penyelesaian seimbang';
  $('ph-net-count-note').textContent = netCount > 0 ? 'Loan No baru lebih banyak daripada yang selesai' : netCount < 0 ? 'Loan No selesai lebih banyak daripada yang baru' : 'Jumlah Loan No seimbang';
  progresRenderRanking('ph-top-new', data.topNew, 'new');
  progresRenderRanking('ph-top-done', data.topDone, 'done');
  progresDayChart('ph-chart-out', data.daily, 'newBalance', 'doneBalance', true);
  progresDayChart('ph-chart-client', data.daily, 'newCount', 'doneCount', false);
  progresRenderTopUp(data);

  const warnings = [];
  if (data.scopeMessage) warnings.push(data.scopeMessage);
  if (data.missingSheets && data.missingSheets.length) warnings.push('Sheet belum ditemukan: ' + data.missingSheets.join(', ') + '. Buat sheet melalui proses unggah versi terbaru.');
  if (data.missingDate) warnings.push(fmt(data.missingDate) + ' baris arsip memiliki Tanggal Progress kosong. Baris tersebut tidak masuk hitungan bulan berjalan.');
  if (data.invalidDate) warnings.push(fmt(data.invalidDate) + ' baris arsip memiliki Tanggal Progress tidak valid. Baris tersebut tidak masuk hitungan bulan berjalan.');
  const unknown = data.unknownTopUp || {};
  if (unknown.newCount || unknown.doneCount) warnings.push('Kategori TOP UP di luar NON TOP-UP, REGULER, dan KHUSUS: PAR Baru ' + fmt(unknown.newCount || 0) + ' Loan No; PAR Selesai ' + fmt(unknown.doneCount || 0) + ' Loan No.');
  const notice = $('ph-notice'); notice.hidden = !warnings.length; notice.textContent = warnings.join(' ');
  $('ph-sync').textContent = 'Terakhir diperiksa: ' + (data.generatedAt || '').replace('T', ' ') + (data.timezone ? ' · Zona waktu ' + data.timezone : '');
  $('ph-export').disabled = false;
}

async function loadProgresHarian(options) {
  if (state.progressLoading || !state.user) return;
  const automatic = !!(options && options.automatic);
  state.progressLoading = true;
  $('ph-refresh').disabled = true; $('ph-export').disabled = true;
  if (!automatic) $('ph-sync').textContent = 'Mengambil arsip PAR Baru dan PAR Selesai…';
  setNote('ph-msg', '');
  try {
    if (automatic && !reserveAutoSync()) return;
    const r = await api('webProgresHarian', { token: (session.get() || {}).token }, { retry: true });
    if (r.code === 'SESSION_REQUIRED') return logout(r.error);
    if (!r.ok) {
      state.progress = null;
      $('ph-meta').textContent = 'Progres Harian gagal dimuat.';
      $('ph-notice').hidden = false; $('ph-notice').textContent = r.error || 'Terjadi kesalahan saat mengambil data.';
      $('ph-sync').textContent = '';
      return;
    }
    state.progress = r; state.progressLoadedAt = Date.now();
    paintProgresHarian(r);
  } catch (err) {
    $('ph-notice').hidden = false;
    $('ph-notice').textContent = err && err.message ? err.message : 'Jaringan tidak dapat dihubungi. Periksa koneksi lalu muat ulang.';
  } finally {
    state.progressLoading = false; $('ph-refresh').disabled = false;
    $('ph-export').disabled = !state.progress;
  }
}

function progresExcelSheet(title, subtitle, headers, rows, widths, numericColumns) {
  const aoa = [[title], [subtitle], [], headers].concat(rows);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const lastCol = Math.max(0, headers.length - 1), headerRow = 3;
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: lastCol } }, { s: { r: 1, c: 0 }, e: { r: 1, c: lastCol } }];
  ws['!cols'] = (widths || headers.map(() => 18)).map(w => ({ wch: w }));
  ws['!freeze'] = { xSplit: 0, ySplit: 4 };
  const paint = (cell, fill, color, bold, size) => { if (ws[cell]) ws[cell].s = {
    font: { name: 'Aptos', sz: size || 10, bold: !!bold, color: { rgb: color || '18352B' } },
    fill: { patternType: 'solid', fgColor: { rgb: fill } },
    alignment: { vertical: 'center', wrapText: true },
    border: { bottom: { style: 'thin', color: { rgb: 'D5E5DC' } } }
  }; };
  paint('A1', '064E3B', 'FFFFFF', true, 16); paint('A2', 'F3FAF6', '405D50', false, 10);
  for (let c = 0; c < headers.length; c++) paint(XLSX.utils.encode_cell({ r: headerRow, c }), '0A7A5A', 'FFFFFF', true, 10);
  rows.forEach((row, i) => {
    const ri = 4 + i;
    for (let c = 0; c < row.length; c++) {
      const cell = XLSX.utils.encode_cell({ r: ri, c });
      if (ws[cell]) {
        const isTotal = /^TOTAL/i.test(String(row[0] || ''));
        ws[cell].s = { font: { name: 'Aptos', sz: 10, bold: isTotal, color: { rgb: isTotal ? '064E3B' : '18352B' } },
          fill: { patternType: 'solid', fgColor: { rgb: isTotal ? 'E3F5ED' : (i % 2 ? 'F7FBF8' : 'FFFFFF') } },
          border: { bottom: { style: isTotal ? 'thin' : 'hair', color: { rgb: 'E2ECE6' } }, ...(isTotal ? { top: { style: 'thin', color: { rgb: '0A7A5A' } } } : {}) }, alignment: { vertical: 'center' } };
        if ((numericColumns || []).includes(c)) { ws[cell].t = typeof row[c] === 'number' ? 'n' : ws[cell].t; ws[cell].z = '#,##0;[Red]-#,##0'; }
      }
    }
  });
  ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: headerRow, c: 0 }, e: { r: headerRow + Math.max(rows.length, 1), c: lastCol } }) };
  return ws;
}

async function exportProgresHarian() {
  const d = state.progress; if (!d) return;
  const btn = $('ph-export'); btn.disabled = true; setNote('ph-msg', 'Menyiapkan laporan Excel…');
  try {
    await Promise.all([loadXLSX(), nextFrame()]);
    const wb = XLSX.utils.book_new(), add = (name, ws) => XLSX.utils.book_append_sheet(wb, ws, name);
    const sub = 'Periode ' + d.monthLabel + ' · s.d. ' + fmtStamp(d.asOf) + ' · ' + d.scope;
    const summaryRows = [
      ['PAR Baru · Client PAR', d.totals.newCount], ['PAR Baru · Out. PAR', d.totals.newBalance],
      ['PAR Selesai · Client PAR', d.totals.doneCount], ['PAR Selesai · Out. PAR', d.totals.doneBalance],
      ['Selisih Client PAR', d.netCount], ['Selisih Out. PAR', d.netBalance]
    ];
    add('Ringkasan', progresExcelSheet('PROGRES HARIAN', sub, ['Indikator', 'Nilai'], summaryRows, [38, 24], [1]));
    add('Top 3 PAR Baru', progresExcelSheet('3 CABANG KENAIKAN TERTINGGI', sub, ['Cabang', 'Client PAR / Loan No', 'Out. PAR'],
      (d.topNew || []).map(x => [x.branch, x.newCount, x.newBalance]), [38, 24, 24], [1, 2]));
    add('Top 3 PAR Selesai', progresExcelSheet('3 CABANG PENYELESAIAN TERTINGGI', sub, ['Cabang', 'Client PAR / Loan No', 'Out. PAR'],
      (d.topDone || []).map(x => [x.branch, x.doneCount, x.doneBalance]), [38, 24, 24], [1, 2]));
    const dailyOutRows = (d.daily || []).map(x => [x.date.split('-').reverse().join('-'), x.newBalance, x.doneBalance, x.newBalance - x.doneBalance]);
    dailyOutRows.push(['TOTAL BULAN', d.totals.newBalance, d.totals.doneBalance, d.netBalance]);
    const dailyClientRows = (d.daily || []).map(x => [x.date.split('-').reverse().join('-'), x.newCount, x.doneCount, x.newCount - x.doneCount]);
    dailyClientRows.push(['TOTAL BULAN', d.totals.newCount, d.totals.doneCount, d.netCount]);
    const topUpRows = (d.topUp || []).map(x => [x.category, x.newCount, x.newBalance, x.doneCount, x.doneBalance]);
    topUpRows.push(['TOTAL 3 KATEGORI', topUpRows.reduce((a, x) => a + x[1], 0), topUpRows.reduce((a, x) => a + x[2], 0),
      topUpRows.reduce((a, x) => a + x[3], 0), topUpRows.reduce((a, x) => a + x[4], 0)]);
    const branchRows = (d.branches || []).map(x => [x.branch, x.newCount, x.newBalance, x.doneCount, x.doneBalance]);
    branchRows.push(['TOTAL', d.totals.newCount, d.totals.newBalance, d.totals.doneCount, d.totals.doneBalance]);
    add('Progress Out PAR', progresExcelSheet('PROGRES OUT. PAR PER HARI', sub, ['Tanggal Progress', 'PAR Baru', 'PAR Selesai', 'Selisih Harian'],
      dailyOutRows, [20, 24, 24, 24], [1, 2, 3]));
    add('Progress Client', progresExcelSheet('PROGRES CLIENT PAR PER HARI', sub, ['Tanggal Progress', 'PAR Baru · Loan No', 'PAR Selesai · Loan No', 'Selisih Harian'],
      dailyClientRows, [20, 24, 26, 20], [1, 2, 3]));
    add('TOP UP', progresExcelSheet('KOMULATIF PAR BARU VS PAR SELESAI BERDASARKAN TOP UP', sub,
      ['Kategori', 'PAR Baru · Client', 'PAR Baru · Out. PAR', 'PAR Selesai · Client', 'PAR Selesai · Out. PAR'],
      topUpRows, [20, 20, 24, 24, 24], [1, 2, 3, 4]));
    add('Per Cabang', progresExcelSheet('REKAP PROGRES PER CABANG', sub,
      ['Cabang', 'PAR Baru · Client', 'PAR Baru · Out. PAR', 'PAR Selesai · Client', 'PAR Selesai · Out. PAR'],
      branchRows, [38, 20, 24, 24, 24], [1, 2, 3, 4]));
    add('Petunjuk', progresExcelSheet('PETUNJUK LAPORAN', sub, ['Istilah', 'Keterangan'], [
      ['Sumber', 'Sheet PAR Baru dan PAR Selesai.'],
      ['Periode', 'Tanggal Progress dari tanggal 1 bulan berjalan sampai tanggal proses laporan.'],
      ['Client PAR', 'Jumlah Loan No tidak kosong pada arsip.'],
      ['Out. PAR', 'Penjumlahan Total Balance pada baris arsip.'],
      ['PAR Baru', 'Loan No yang tidak ditemukan pada Data_Upload ketika pemeriksaan upload.'],
      ['PAR Selesai', 'Loan No pada Data_Upload lama yang tidak ditemukan pada Excel terbaru.'],
      ['Cakupan', 'Administrator melihat seluruh cabang; pengguna lain hanya cabang pada Wilayah_Kerja.'],
      ['TOP UP lainnya', 'Nilai TOP UP di luar NON TOP-UP, REGULER, dan KHUSUS tetap masuk total utama namun tidak masuk tabel tiga kategori.']
    ], [24, 100], []));
    XLSX.writeFile(wb, 'Progres_Harian_' + d.monthKey + '.xlsx');
    setNote('ph-msg', 'Laporan Excel berhasil dibuat (8 sheet).', 'good');
  } catch (err) {
    setNote('ph-msg', err && err.message ? err.message : 'Gagal membuat laporan Excel.', 'bad');
  } finally { btn.disabled = false; }
}

/* ---------------------------------------------------------------------
   Dashboard: Analisa Cabang
   Data telah dibatasi Wilayah_Kerja di server. Filter dan agregasi di sini
   hanya bekerja pada data yang memang boleh dilihat akun aktif.
   --------------------------------------------------------------------- */
const ANALISA_METRICS = [
  { key: 'member', label: 'Member', kind: 'count', hint: 'Jumlah member' },
  { key: 'outstanding', label: 'Outstanding', kind: 'money', hint: 'Saldo outstanding' },
  { key: 'disbursement', label: 'Disbursement', kind: 'money', hint: 'Penyaluran' },
  { key: 'saving', label: 'Saving', kind: 'money', hint: 'Simpanan' },
  { key: 'repayment', label: 'Repayment', kind: 'money', hint: 'Pengembalian pokok' },
  { key: 'margin', label: 'Margin', kind: 'money', hint: 'Interest collection' },
  { key: 'memberPAR', label: 'Member PAR', kind: 'count', hint: 'Jumlah member PAR' },
  { key: 'outstandingPAR', label: 'Outstanding PAR', kind: 'money', hint: 'Saldo outstanding PAR' }
];
const AN_MONTHS = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
function analisaLabelDate(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
  return m ? Number(m[3]) + ' ' + AN_MONTHS[Number(m[2]) - 1] + ' ' + m[1] : iso || '—';
}
// Arah perbandingan wajib dari periode yang lebih lama menuju periode terpilih yang lebih baru.
// Kunci periode memakai ISO yyyy-MM-dd sehingga perbandingan string bersifat kronologis.
function analisaIsEarlierPeriod(comparePeriod, selectedPeriod) {
  return /^\d{4}-\d{2}-\d{2}$/.test(comparePeriod || '')
    && /^\d{4}-\d{2}-\d{2}$/.test(selectedPeriod || '')
    && comparePeriod < selectedPeriod;
}
function analisaFormatMetric(metric, value) {
  return metric.kind === 'money' ? rpShort(value || 0) : fmt(value || 0);
}
function analisaAggregate(rows) {
  const out = {};
  ANALISA_METRICS.forEach(m => out[m.key] = 0);
  rows.forEach(r => ANALISA_METRICS.forEach(m => { out[m.key] += Number(r[m.key]) || 0; }));
  return out;
}
function analisaRowsFor(period) {
  const d = state.analisaCabang, f = state.analisaFilters;
  if (!d || !period) return [];
  return d.records.filter(r => r.period === period && (f.branch === '*' || r.key === f.branch));
}
function analisaUpdatePeriodOptions(keepPeriod) {
  const d = state.analisaCabang, f = state.analisaFilters;
  if (!d) return;
  const filtered = (d.periods || []).filter(p => {
    const y = p.slice(0, 4), m = String(Number(p.slice(5, 7)));
    const branchHasPeriod = f.branch === '*' || d.records.some(r => r.period === p && r.key === f.branch);
    return branchHasPeriod && (f.year === '*' || f.year === y) && (f.month === '*' || f.month === m);
  });
  const period = keepPeriod && filtered.includes(f.period) ? f.period : (filtered[0] || '');
  f.period = period;
  const options = [h('option', { value: '', text: filtered.length ? 'Pilih periode' : 'Tidak ada periode' })]
    .concat(filtered.map(p => h('option', { value: p, text: analisaLabelDate(p) })));
  $('an-period').replaceChildren(...options); $('an-period').value = period;
  analisaUpdateCompareOptions();
}
function analisaUpdateCompareOptions() {
  const d = state.analisaCabang, f = state.analisaFilters;
  // Hanya periode yang benar-benar lebih lama dari Periode Analisa terpilih yang boleh menjadi pembanding.
  const comparable = d ? (d.periods || []).filter(p => analisaIsEarlierPeriod(p, f.period)
    && (f.branch === '*' || d.records.some(r => r.period === p && r.key === f.branch)))
    .sort((a, b) => b.localeCompare(a)) : [];
  const options = comparable.length
    ? [h('option', { value: '', text: 'Tidak dibandingkan' })]
    : [h('option', { value: '', text: f.period ? 'Tidak ada tanggal yang lebih lama' : 'Pilih periode analisa dahulu' })];
  comparable.forEach(p => options.push(h('option', { value: p, text: analisaLabelDate(p) })));
  $('an-compare').replaceChildren(...options);
  $('an-compare').disabled = !f.period || !comparable.length;
  // Default pembanding = tanggal terdekat yang lebih lama pada cabang/cakupan yang sama.
  const earlier = comparable[0] || '';
  if (!analisaIsEarlierPeriod(f.compare, f.period) || !comparable.includes(f.compare)) f.compare = earlier;
  $('an-compare').value = f.compare;
}
function analisaInitFilters(data) {
  const f = state.analisaFilters;
  const years = [...new Set((data.periods || []).map(p => p.slice(0, 4)))].sort().reverse();
  const branches = data.branches || [];
  const branchCurrent = branches.some(b => b.toLocaleUpperCase('id-ID') === f.branch) ? f.branch : '*';
  f.branch = branchCurrent;
  $('an-branch').replaceChildren(h('option', { value: '*', text: 'Semua Cabang' }), ...branches.map(b => h('option', { value: b.toLocaleUpperCase('id-ID'), text: b })));
  $('an-branch').value = f.branch;
  $('an-year').replaceChildren(h('option', { value: '*', text: 'Semua Tahun' }), ...years.map(y => h('option', { value: y, text: y })));
  if (f.year !== '*' && !years.includes(f.year)) f.year = '*';
  $('an-year').value = f.year;
  $('an-month').replaceChildren(h('option', { value: '*', text: 'Semua Bulan' }), ...AN_MONTHS.map((m, i) => h('option', { value: String(i + 1), text: m })));
  $('an-month').value = f.month;
  analisaUpdatePeriodOptions(true);
}
function analisaRenderCards(current, compare, comparePeriod) {
  const host = $('an-kpis');
  put(host, ANALISA_METRICS.map(metric => {
    const value = current[metric.key] || 0, base = compare ? (compare[metric.key] || 0) : null;
    let changeText = 'Belum dibandingkan', changeClass = 'neutral', title = '';
    if (base !== null) {
      if (base === 0 && value === 0) { changeText = '0,0% · tetap'; }
      else if (base === 0) { changeText = 'Baru · dasar 0'; changeClass = 'up'; }
      else {
        const pct = (value - base) / Math.abs(base) * 100;
        const sign = pct > 0 ? '+' : '';
        changeText = sign + nf1.format(pct) + '% · ' + (pct > 0 ? 'naik' : pct < 0 ? 'turun' : 'tetap');
        changeClass = pct > 0 ? 'up' : pct < 0 ? 'down' : 'neutral';
      }
      title = 'Periode pembanding ' + analisaLabelDate(comparePeriod) + ': ' + analisaFormatMetric(metric, base);
    }
    return h('article', { class: 'an-kpi-card' },
      h('span', { class: 'an-kpi-label', text: metric.label }),
      h('strong', { class: 'an-kpi-value', text: analisaFormatMetric(metric, value), title: metric.kind === 'money' ? rp(value) : fmt(value) }),
      h('span', { class: 'an-kpi-hint', text: metric.hint }),
      h('div', { class: 'an-change ' + changeClass, title }, h('b', { text: changeText }),
        h('small', { text: comparePeriod ? 'vs ' + analisaLabelDate(comparePeriod) : 'Pilih tanggal pembanding' })));
  }));
}
function analisaRenderTable(rows) {
  const headers = ['Nama Cabang','Member','Outstanding','Disbursement','Saving','Repayment','Margin','Member PAR','Outstanding PAR'];
  const table = $('an-table');
  const head = h('thead', null, h('tr', null, headers.map((v,i) => h('th', { class: i ? 'num' : '', text: v }))));
  if (!rows.length) {
    table.replaceChildren(head, h('tbody', null, h('tr', null, h('td', { colspan: headers.length, class: 'an-empty-cell', text: 'Tidak ada data pada periode dan filter cabang yang dipilih.' }))));
    return;
  }
  rows = rows.slice().sort((a,b) => collator.compare(a.branch,b.branch));
  const tbody = h('tbody', null, rows.map(r => h('tr', null,
    h('th', { scope: 'row', text: r.branch }),
    h('td', { class: 'num', text: fmt(r.member) }),
    h('td', { class: 'num', text: rp(r.outstanding) }),
    h('td', { class: 'num', text: rp(r.disbursement) }),
    h('td', { class: 'num', text: rp(r.saving) }),
    h('td', { class: 'num', text: rp(r.repayment) }),
    h('td', { class: 'num', text: rp(r.margin) }),
    h('td', { class: 'num', text: fmt(r.memberPAR) }),
    h('td', { class: 'num', text: rp(r.outstandingPAR) }))));
  const total = analisaAggregate(rows);
  const foot = h('tfoot', null, h('tr', null,
    h('th', { text: 'TOTAL · ' + fmt(rows.length) + ' cabang' }),
    h('td', { class: 'num', text: fmt(total.member) }),
    h('td', { class: 'num', text: rp(total.outstanding) }),
    h('td', { class: 'num', text: rp(total.disbursement) }),
    h('td', { class: 'num', text: rp(total.saving) }),
    h('td', { class: 'num', text: rp(total.repayment) }),
    h('td', { class: 'num', text: rp(total.margin) }),
    h('td', { class: 'num', text: fmt(total.memberPAR) }),
    h('td', { class: 'num', text: rp(total.outstandingPAR) })));
  table.replaceChildren(head, tbody, foot);
}
function analisaRenderDashboard() {
  const d = state.analisaCabang, f = state.analisaFilters;
  if (!d) return;
  // Pertahanan tambahan: jangan pernah menghitung/menampilkan perbandingan dengan tanggal yang sama atau lebih baru.
  if (f.compare && !analisaIsEarlierPeriod(f.compare, f.period)) {
    f.compare = '';
    analisaUpdateCompareOptions();
  }
  $('an-scope').textContent = d.scope || 'Sesuai Wilayah Kerja';
  $('an-meta').textContent = d.periods && d.periods.length ? 'Data sheet Analisa · ' + d.periods.length + ' periode tersedia · ' + (d.records || []).length + ' baris sesuai cakupan akses.' : (d.message || 'Belum ada data Analisa.');
  $('an-sync').textContent = d.generatedAt ? 'Terakhir diperiksa: ' + fmtStamp(d.generatedAt) + (d.timezone ? ' · ' + d.timezone : '') : '';
  const period = f.period;
  const curRows = analisaRowsFor(period), compareRows = analisaRowsFor(f.compare);
  const current = analisaAggregate(curRows), compare = f.compare ? analisaAggregate(compareRows) : null;
  analisaRenderCards(current, compare, f.compare);
  analisaRenderTable(curRows);
  $('an-table-caption').textContent = (f.branch === '*' ? 'Seluruh cabang dalam cakupan akses' : 'Cabang: ' + ($('an-branch').selectedOptions[0] ? $('an-branch').selectedOptions[0].textContent : f.branch))
    + ' · Periode ' + (period ? analisaLabelDate(period) : 'belum dipilih')
    + (f.compare ? ' · dibandingkan dengan ' + analisaLabelDate(f.compare) : ' · belum ada periode pembanding');
  const notices = [];
  if (d.scopeMessage) notices.push(d.scopeMessage);
  if (d.missingPeriod) notices.push(fmt(d.missingPeriod) + ' baris di sheet Analisa diabaikan karena Periode Analisa kosong/tidak valid.');
  if (period && !curRows.length) notices.push('Tidak ditemukan data untuk periode ' + analisaLabelDate(period) + '.');
  $('an-notice').hidden = !notices.length; $('an-notice').textContent = notices.join(' ');
  $('an-export').disabled = !period || !curRows.length;
}
async function loadAnalisaCabang(options) {
  if (state.analisaLoading || !state.user) return;
  const automatic = !!(options && options.automatic);
  state.analisaLoading = true;
  $('an-refresh').disabled = true; $('an-export').disabled = true;
  if (!automatic) $('an-sync').textContent = 'Mengambil data dari sheet Analisa…';
  setNote('an-msg', '');
  try {
    if (automatic && !reserveAutoSync()) return;
    const r = await api('webAnalisaCabang', { token: (session.get() || {}).token }, { retry: true });
    if (r.code === 'SESSION_REQUIRED') return logout(r.error);
    if (!r.ok) {
      state.analisaCabang = null;
      $('an-notice').hidden = false; $('an-notice').textContent = r.error || 'Gagal mengambil data Analisa.';
      $('an-meta').textContent = 'Analisa Cabang gagal dimuat.';
      return;
    }
    state.analisaCabang = r; state.analisaLoadedAt = Date.now();
    analisaInitFilters(r); analisaRenderDashboard();
    if (automatic) $('an-sync').textContent += ' · sinkron otomatis';
  } catch (err) {
    $('an-notice').hidden = false; $('an-notice').textContent = err && err.message ? err.message : 'Jaringan tidak dapat dihubungi.';
  } finally {
    state.analisaLoading = false; $('an-refresh').disabled = false;
    $('an-export').disabled = !state.analisaCabang || !state.analisaFilters.period || !analisaRowsFor(state.analisaFilters.period).length;
  }
}
function analisaExportSheet(title, subtitle, headers, rows, widths, numericCols) {
  const ws = XLSX.utils.aoa_to_sheet([[title],[subtitle],[],headers].concat(rows));
  ws['!merges'] = [{ s: { r:0,c:0 }, e:{ r:0,c:headers.length-1 } }, { s:{r:1,c:0},e:{r:1,c:headers.length-1} }];
  ws['!cols'] = widths.map(w => ({ wch:w })); ws['!freeze'] = { xSplit:0, ySplit:4 };
  const style = (cell, fill, color, bold, size) => { if (ws[cell]) ws[cell].s = { font:{name:'Aptos',sz:size||10,bold:!!bold,color:{rgb:color||'18352B'}}, fill:{patternType:'solid',fgColor:{rgb:fill}}, alignment:{vertical:'center',wrapText:true}, border:{bottom:{style:'thin',color:{rgb:'D5E5DC'}}} }; };
  style('A1','064E3B','FFFFFF',true,15); style('A2','F3FAF6','405D50',false,10);
  headers.forEach((_,c) => style(XLSX.utils.encode_cell({r:3,c}),'0A7A5A','FFFFFF',true,10));
  rows.forEach((row,i) => row.forEach((v,c) => { const cell = XLSX.utils.encode_cell({r:4+i,c}); if (ws[cell] && numericCols.includes(c)) ws[cell].z = '#,##0.00;[Red](#,##0.00);-'; }));
  if (rows.length) ws['!autofilter'] = { ref:XLSX.utils.encode_range({s:{r:3,c:0},e:{r:3+rows.length,c:headers.length-1}}) };
  return ws;
}
async function exportAnalisaCabang() {
  const d = state.analisaCabang, f = state.analisaFilters;
  if (!d || !f.period) return;
  const comparePeriod = analisaIsEarlierPeriod(f.compare, f.period) ? f.compare : '';
  const btn = $('an-export'); btn.disabled = true; setNote('an-msg','Menyiapkan laporan Excel…');
  try {
    await Promise.all([loadXLSX(), nextFrame()]);
    const currentRows = analisaRowsFor(f.period).slice().sort((a,b) => collator.compare(a.branch,b.branch));
    const compareRows = comparePeriod ? analisaRowsFor(comparePeriod) : [];
    const cur = analisaAggregate(currentRows), old = comparePeriod ? analisaAggregate(compareRows) : null;
    const summaryRows = ANALISA_METRICS.map(m => {
      const now = cur[m.key] || 0, before = old ? old[m.key] || 0 : '';
      const delta = old ? now - before : '';
      const pct = old ? (before === 0 ? (now === 0 ? 0 : 'N/A') : (now - before) / Math.abs(before) * 100) : '';
      return [m.label, now, before, delta, pct];
    });
    const detailRows = currentRows.map(r => [r.branch,r.member,r.outstanding,r.disbursement,r.saving,r.repayment,r.margin,r.memberPAR,r.outstandingPAR]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, analisaExportSheet('RINGKASAN ANALISA CABANG','Periode '+analisaLabelDate(f.period)+' · Pembanding '+(comparePeriod?analisaLabelDate(comparePeriod):'tidak dipilih')+' · '+d.scope,
      ['Indikator','Periode Terpilih','Periode Pembanding','Selisih','Perubahan (%)'], summaryRows,[28,22,22,22,18],[1,2,3,4]), 'Ringkasan');
    XLSX.utils.book_append_sheet(wb, analisaExportSheet('RINCIAN ANALISA PER CABANG','Periode '+analisaLabelDate(f.period)+' · '+d.scope,
      ['Nama Cabang','Member','Outstanding','Disbursement','Saving','Repayment','Margin','Member PAR','Outstanding PAR'], detailRows,[34,15,20,20,20,20,20,15,22],[1,2,3,4,5,6,8]), 'Per Cabang');
    XLSX.utils.book_append_sheet(wb, analisaExportSheet('PETUNJUK','Sumber: sheet Analisa. Persentase perubahan = (Periode Terpilih - Periode Pembanding) / nilai absolut Periode Pembanding × 100%.',
      ['Istilah','Keterangan'],[['Periode Analisa','Tanggal diambil dari nama file Analisa dd-mm-yyyy.'],['Cabang','Semua Cabang atau satu Nama Cabang.'],['Cakupan','Administrator melihat semua cabang; pengguna lain mengikuti Wilayah_Kerja.'],['Nilai pembanding 0','Perubahan ditampilkan N/A bila nilai periode pembanding nol dan nilai terbaru bukan nol.']],[25,100],[]), 'Petunjuk');
    XLSX.writeFile(wb, 'Analisa_Cabang_'+f.period+'.xlsx');
    setNote('an-msg','Laporan Excel berhasil dibuat (3 sheet).','good');
  } catch (err) { setNote('an-msg',err && err.message ? err.message : 'Gagal membuat laporan Excel.','bad'); }
  finally { btn.disabled = false; }
}

/* ---------------------------------------------------------------------
   Setting: Penyesuaian Margin (ADMINISTRATOR dan USER 0; server memeriksa ulang)
   --------------------------------------------------------------------- */
async function saveMargin() {
  const raw = $('mval').value.trim().replace(',', '.'), v = Number(raw);
  if (raw === '' || !isFinite(v) || v < 0 || v > 100) { $('mval').classList.add('bad'); return setNote('setmsg', 'Isi angka 0 sampai 100.', 'bad'); }
  $('mval').classList.remove('bad');
  const ok = await askConfirm('Simpan Penyesuaian Margin?', 'Semua perhitungan pelunasan akan memakai Total Balance + ' + nf1.format(v) + '%. Data dashboard dihitung ulang untuk semua pengguna.', 'Simpan');
  if (!ok) return;
  $('msave').disabled = true; setNote('setmsg', 'Menyimpan dan menghitung ulang…');
  const r = await api('webSetMargin', { token: (session.get() || {}).token, value: String(v) });
  $('msave').disabled = false;
  if (r.code === 'SESSION_REQUIRED') return logout(r.error);
  if (!r.ok) return setNote('setmsg', r.error || 'Gagal menyimpan.', 'bad');
  setNote('setmsg', 'Tersimpan: Penyesuaian Margin ' + nf1.format(r.margin) + '%.', 'good');
  loadDashFresh();
}

/* ---------------------------------------------------------------------
   Setting: Include: Simpanan Wajib (ADMINISTRATOR dan USER 0; server memeriksa ulang)
   Satu kotak centang per kategori. Dicentang = Simp. Wajib ikut dihitung di kategori itu, tidak dicentang = tidak.
   --------------------------------------------------------------------- */
const wajibChecked = () => CATS.filter(c => document.querySelector('#wjlist input[value="' + c.k + '"]').checked).map(c => c.k).join('');
const wajibNames = codes => CATS.filter(c => codes.indexOf(c.k) > -1).map(c => c.title).join(', ') || 'tidak ada kategori';

function buildWajibForm() {
  put($('wjlist'), CATS.map(c => h('label', { class: 'chk' },
    h('input', { type: 'checkbox', value: c.k }),
    h('span', null, h('b', { text: c.title }), h('small', { text: c.rule })))));
}
function syncWajibForm() {   // tombol Simpan aktif hanya bila pilihan berbeda dari yang tersimpan
  $('wjsave').disabled = !state.data || wajibChecked() === state.data.wajib;
}
function paintWajib() {
  if (!state.data) return;
  const cur = state.data.wajib;
  $('wjcur').textContent = wajibNames(cur);
  if (!state.wjDirty) document.querySelectorAll('#wjlist input').forEach(i => { i.checked = cur.indexOf(i.value) > -1; });   // jangan menimpa pilihan yang sedang diubah
  syncWajibForm();
}
async function saveWajib() {
  if (!state.data) return setNote('wjmsg', 'Data dashboard masih dimuat. Coba lagi sebentar lagi.', 'bad');
  const codes = wajibChecked();
  if (codes === state.data.wajib) return setNote('wjmsg', 'Tidak ada perubahan.');
  const ok = await askConfirm('Simpan Include: Simpanan Wajib?',
    (codes ? 'Simp. Wajib ikut dihitung di: ' + wajibNames(codes) + '. Kategori lainnya tanpa Simp. Wajib.' : 'Simp. Wajib tidak dihitung di kategori mana pun.')
    + ' Data dashboard dihitung ulang untuk semua pengguna.', 'Simpan');
  if (!ok) return;
  $('wjsave').disabled = true; setNote('wjmsg', 'Menyimpan dan menghitung ulang…');
  const r = await api('webSetWajib', { token: (session.get() || {}).token, cats: codes ? codes.split('').join(',') : '-' });
  if (r.code === 'SESSION_REQUIRED') return logout(r.error);
  if (!r.ok) { syncWajibForm(); return setNote('wjmsg', r.error || 'Gagal menyimpan.', 'bad'); }
  state.wjDirty = false;
  setNote('wjmsg', 'Tersimpan: ' + (r.wajib ? 'Simp. Wajib dihitung di ' + wajibNames(r.wajib) + '.' : 'Simp. Wajib tidak dihitung di kategori mana pun.'), 'good');
  loadDashFresh();
}

/* ---------------------------------------------------------------------
   Setting: Wilayah Kerja (ADMINISTRATOR dan USER 0; akses server diperiksa ulang)
   --------------------------------------------------------------------- */
function setSettingPage(which) {
  const isWilayah = which === 'wilayah';
  document.querySelectorAll('[data-setpage]').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.setpage === which)));
  $('set-global').hidden = isWilayah;
  $('set-wilayah').hidden = !isWilayah;
  if (isWilayah && (!state.wilayah || (!wilayahIsDirty() && Date.now() - state.wilayahLoadedAt > 60000))) loadWilayah();
}

function wilayahUserByNik(nik) {
  return state.wilayah && state.wilayah.users.find(u => u.nik === nik);
}
function wilayahSort(a, b) { return collator.compare(a, b); }
function wilayahCurrentBranches(nik) {
  return state.wilayah && state.wilayah.assignments[nik] ? state.wilayah.assignments[nik].slice() : [];
}
function wilayahDraftFor(nik) {
  if (nik === state.wilayahSelected) return state.wilayahDraft.slice();
  if (state.wilayahDrafts && state.wilayahDrafts[nik]) return state.wilayahDrafts[nik].slice();
  return wilayahCurrentBranches(nik);
}
function wilayahUserDirty(nik) {
  const a = wilayahCurrentBranches(nik).slice().sort(wilayahSort);
  const b = wilayahDraftFor(nik).slice().sort(wilayahSort);
  return a.length !== b.length || a.some((v, i) => v !== b[i]);
}
function wilayahPendingUsers() {
  return state.wilayah ? state.wilayah.users.filter(u => wilayahUserDirty(u.nik)) : [];
}
function wilayahIsDirty() { return wilayahPendingUsers().length > 0; }

async function loadWilayah() {
  if (state.wilayahLoading) return;
  state.wilayahLoading = true;
  setNote('wilayah-msg', 'Memuat daftar user dan cabang…');
  $('wilayah-save').disabled = true;
  const r = await api('webGetWilayah', { token: (session.get() || {}).token }, { retry: true });
  state.wilayahLoading = false;
  if (r.code === 'SESSION_REQUIRED') return logout(r.error);
  if (!r.ok) { setNote('wilayah-msg', r.error || 'Gagal memuat Wilayah Kerja.', 'bad'); return; }
  state.wilayah = {
    users: Array.isArray(r.users) ? r.users : [],
    branches: Array.isArray(r.branches) ? r.branches : [],
    assignments: r.assignments || {}, source: r.source || 'Data_Upload'
  };
  state.wilayahLoadedAt = Date.now();
  state.wilayahDrafts = Object.create(null);
  if (!wilayahUserByNik(state.wilayahSelected)) state.wilayahSelected = state.wilayah.users.length ? state.wilayah.users[0].nik : '';
  state.wilayahDraft = wilayahCurrentBranches(state.wilayahSelected);
  $('wilayah-user-count').textContent = fmt(state.wilayah.users.length) + ' user';
  setNote('wilayah-msg', 'Sumber cabang: ' + state.wilayah.source + '. Perubahan penugasan langsung membatasi data yang dikirim server.', 'good');
  renderWilayahUsers(); renderWilayahEditor();
}

function renderWilayahUsers() {
  const box = $('wilayah-users');
  if (!state.wilayah) return box.replaceChildren();
  const q = $('wilayah-user-q').value.trim().toLowerCase();
  const users = state.wilayah.users.filter(u => (u.nik + ' ' + u.nama + ' ' + u.tingkat).toLowerCase().includes(q));
  box.replaceChildren(...users.map(u => {
    const count = wilayahDraftFor(u.nik).length;
    return h('button', { type: 'button', class: 'wilayah-user' + (u.nik === state.wilayahSelected ? ' active' : ''),
      'data-nik': u.nik, 'aria-pressed': String(u.nik === state.wilayahSelected), onclick: () => selectWilayahUser(u.nik) },
      h('span', { class: 'wilayah-user-text' }, h('b', { text: u.nama || '(Nama kosong)' }),
        h('small', { text: u.nik + ' · ' + u.tingkat })),
      h('span', { class: 'wilayah-user-count', text: count + ' cabang' }));
  }));
  if (!users.length) box.append(h('p', { class: 'dim wilayah-empty', text: 'User tidak ditemukan.' }));
}

function selectWilayahUser(nik) {
  if (wilayahIsDirty()) {
    // Perubahan draft tetap disimpan di memori per user agar berpindah user tidak menghilangkannya.
    if (!state.wilayahDrafts) state.wilayahDrafts = Object.create(null);
    state.wilayahDrafts[state.wilayahSelected] = state.wilayahDraft.slice();
  }
  state.wilayahSelected = nik;
  state.wilayahDraft = state.wilayahDrafts && state.wilayahDrafts[nik]
    ? state.wilayahDrafts[nik].slice() : wilayahCurrentBranches(nik);
  renderWilayahUsers(); renderWilayahEditor();
}

function updateWilayahDraft() {
  const count = state.wilayahDraft.length, total = state.wilayah ? state.wilayah.branches.length : 0;
  $('wilayah-selected-count').textContent = count + ' / ' + total + ' cabang';
  $('wilayah-save').disabled = !wilayahIsDirty();
  if (state.wilayahDrafts && state.wilayahSelected) state.wilayahDrafts[state.wilayahSelected] = state.wilayahDraft.slice();
}

function renderWilayahEditor() {
  const user = wilayahUserByNik(state.wilayahSelected);
  $('wilayah-prompt').hidden = !!user;
  $('wilayah-edit').hidden = !user;
  if (!user) return;
  $('wilayah-selected-name').textContent = user.nama || '(Nama kosong)';
  $('wilayah-selected-meta').textContent = user.nik + ' · ' + user.tingkat;
  const q = $('wilayah-branch-q').value.trim().toLowerCase();
  const branches = state.wilayah.branches.filter(b => b.name.toLowerCase().includes(q));
  $('wilayah-branches').replaceChildren(...branches.map(b => {
    const label = h('label', { class: 'wilayah-branch' + (state.wilayahDraft.indexOf(b.name) > -1 ? ' selected' : '') });
    const input = h('input', { type: 'checkbox', value: b.name });
    input.checked = state.wilayahDraft.indexOf(b.name) > -1;
    input.addEventListener('change', () => {
      if (input.checked) { if (state.wilayahDraft.indexOf(b.name) < 0) state.wilayahDraft.push(b.name); }
      else state.wilayahDraft = state.wilayahDraft.filter(x => x !== b.name);
      label.classList.toggle('selected', input.checked);
      updateWilayahDraft();
    });
    label.append(input, h('span', null, h('b', { text: b.name }),
      h('small', { text: b.active ? 'Ada pada data unggahan' : 'Tidak ada pada data unggahan saat ini' })));
    return label;
  }));
  if (!branches.length) $('wilayah-branches').append(h('p', { class: 'dim wilayah-empty', text: 'Cabang tidak ditemukan.' }));
  updateWilayahDraft();
}

function setWilayahAll(on) {
  if (!state.wilayah) return;
  state.wilayahDraft = on ? state.wilayah.branches.map(b => b.name) : [];
  renderWilayahEditor();
}

async function saveWilayah() {
  const pending = wilayahPendingUsers();
  if (!state.wilayah || !pending.length) return;
  const summary = pending.map(u => u.nama + ' (' + u.nik + '): ' + wilayahDraftFor(u.nik).length + ' cabang').join('\n');
  const ok = await askConfirm('Simpan Wilayah Kerja?',
    'Perubahan untuk ' + pending.length + ' user akan disimpan:\n' + summary + '\nCabang yang tidak ditugaskan tidak akan dikirim ke akun tersebut.', 'Simpan');
  if (!ok) return;
  $('wilayah-save').disabled = true;
  setNote('wilayah-msg', 'Menyimpan pemetaan…');
  const assignments = state.wilayah.users.map(u => ({ nik: u.nik, branches: wilayahDraftFor(u.nik) }));
  const r = await api('webSaveWilayah', { token: (session.get() || {}).token, assignments: JSON.stringify(assignments) });
  if (r.code === 'SESSION_REQUIRED') return logout(r.error);
  if (!r.ok) { updateWilayahDraft(); return setNote('wilayah-msg', r.error || 'Gagal menyimpan Wilayah Kerja.', 'bad'); }
  state.wilayahDrafts = Object.create(null);
  setNote('wilayah-msg', 'Tersimpan. ' + fmt(r.assignments) + ' penugasan cabang untuk ' + fmt(r.users) + ' user.', 'good');
  await loadWilayah();
  loadDashFresh();
}

/* ---------------------------------------------------------------------
   12. Mulai
   --------------------------------------------------------------------- */
function bind() {
  // Masuk
  $('form').addEventListener('submit', submitLogin);
  ['nik', 'pw'].forEach(i => $(i).addEventListener('input', () => $(i).classList.remove('bad')));
  $('eye').onclick = () => {
    const show = $('pw').type === 'password';
    $('pw').type = show ? 'text' : 'password';
    $('eye').textContent = show ? 'Sembunyi' : 'Lihat';
    $('eye').setAttribute('aria-pressed', String(show));
    $('eye').setAttribute('aria-label', show ? 'Sembunyikan password' : 'Tampilkan password');
  };
  $('out').onclick = () => logout();

  // Tab (dengan panah kiri/kanan)
  TABS.forEach(k => { $('t-' + k).onclick = () => selectTab(k); });
  document.querySelectorAll('[data-dashpage]').forEach(b => b.addEventListener('click', () => selectDashboardPage(b.dataset.dashpage)));
  $('app').querySelector('.tabs').addEventListener('keydown', e => {
    const vis = TABS.filter(k => !$('t-' + k).hidden), i = vis.findIndex(k => $('t-' + k).getAttribute('aria-selected') === 'true');
    const to = e.key === 'ArrowRight' ? vis[(i + 1) % vis.length] : e.key === 'ArrowLeft' ? vis[(i - 1 + vis.length) % vis.length]
      : e.key === 'Home' ? vis[0] : e.key === 'End' ? vis[vis.length - 1] : null;
    if (to) { e.preventDefault(); selectTab(to, true); }
  });

  // Dashboard: Rekening Bergerak
  $('rb-ref').onclick = () => loadRekeningBergerak();
  $('rb-xlsx').onclick = exportRekeningBergerak;
  document.querySelectorAll('[data-rb-view]').forEach(b => b.addEventListener('click', () => selectRekeningTable(b.dataset.rbView)));
  $('rb-branch-q').addEventListener('input', debounce(() => { state.rbSearch.branch = $('rb-branch-q').value; refreshRBTable('branch'); }, 120));
  $('rb-coordinator-q').addEventListener('input', debounce(() => { state.rbSearch.coordinator = $('rb-coordinator-q').value; refreshRBTable('coordinator'); }, 120));
  $('rb-branch-tbl').addEventListener('click', e => {
    const btn = e.target.closest('[data-rb-sort]'); if (!btn) return;
    const sort = state.rbSort.branch;
    sort.dir = sort.key === btn.dataset.rbSort ? sort.dir * -1 : 1; sort.key = btn.dataset.rbSort;
    refreshRBTable('branch');
  });
  $('rb-coordinator-tbl').addEventListener('click', e => {
    const btn = e.target.closest('[data-rb-sort]'); if (!btn) return;
    const sort = state.rbSort.coordinator;
    sort.dir = sort.key === btn.dataset.rbSort ? sort.dir * -1 : 1; sort.key = btn.dataset.rbSort;
    refreshRBTable('coordinator');
  });
  selectRekeningTable('branch');
  $('rb-setting-form').addEventListener('submit', e => { e.preventDefault(); saveRekeningBergerakSetting(); });
  $('rb-setting-date').addEventListener('input', () => { $('rb-setting-date').classList.remove('bad'); setNote('rb-setting-msg', ''); });

  // Dashboard Progres Harian
  $('ph-refresh').onclick = () => loadProgresHarian();
  $('ph-export').onclick = exportProgresHarian;

  // Dashboard Analisa Cabang
  $('an-refresh').onclick = () => loadAnalisaCabang();
  $('an-export').onclick = exportAnalisaCabang;
  $('an-branch').onchange = () => { state.analisaFilters.branch = $('an-branch').value; analisaUpdatePeriodOptions(false); analisaRenderDashboard(); };
  $('an-year').onchange = () => { state.analisaFilters.year = $('an-year').value; analisaUpdatePeriodOptions(false); analisaRenderDashboard(); };
  $('an-month').onchange = () => { state.analisaFilters.month = $('an-month').value; analisaUpdatePeriodOptions(false); analisaRenderDashboard(); };
  $('an-period').onchange = () => { state.analisaFilters.period = $('an-period').value; state.analisaFilters.compare = ''; analisaUpdateCompareOptions(); analisaRenderDashboard(); };
  $('an-compare').onchange = () => {
    const selected = $('an-compare').value, current = state.analisaFilters.period;
    if (selected && !analisaIsEarlierPeriod(selected, current)) {
      state.analisaFilters.compare = '';
      analisaUpdateCompareOptions();
      setNote('an-msg', 'Tanggal pembanding harus lebih lama daripada Periode Analisa yang dipilih.', 'bad');
    } else {
      state.analisaFilters.compare = selected;
      setNote('an-msg', '');
    }
    analisaRenderDashboard();
  };

  // Dashboard Tool PAR (dashboard lama)
  $('dref').onclick = () => loadDash();
  $('ddl').onclick = exportAll;
  $('dxl').onclick = exportView;
  $('dq').addEventListener('input', debounce(() => { state.q = $('dq').value; state.page = 0; drawList(); }, 150));
  $('dbranch').onchange = () => { state.branch = $('dbranch').value; state.staf = ''; state.page = 0; fillStafOptions(); drawList(); };
  $('dstaf').onchange = () => { state.staf = $('dstaf').value; state.page = 0; drawList(); };
  $('dsort').onchange = () => { const v = $('dsort').value; if (!v) { state.sort = { label: '', dir: 1 }; drawList(); } else { const [l, d] = v.split('|'); sortBy(l, +d); } };
  $('dreset').onclick = resetFilters;
  $('dfirst').onclick = () => go(0);
  $('dprev').onclick = () => go(state.page - 1);
  $('dnext').onclick = () => go(state.page + 1);
  $('dlast').onclick = () => go(1e9);
  $('dsize').replaceChildren(...PAGE_SIZES.map(n => h('option', { value: n, text: String(n) })));
  $('dsize').value = String(state.size);
  $('dsize').onchange = () => { state.size = +$('dsize').value; state.page = 0; drawList(); };
  $('dtbl').addEventListener('click', e => {
    const th = e.target.closest('.th');
    if (th) return sortBy(th.dataset.l);
    const tr = e.target.closest('tbody tr');
    if (tr) openDetail(state.view[+tr.dataset.i]);
  });
  $('detail').addEventListener('click', e => { if (e.target === $('detail')) $('detail').close(); });   // klik di luar kotak = tutup

  // Ringkasan
  document.querySelectorAll('.seg button').forEach(b => { b.onclick = () => { state.sumBy = b.dataset.g; state.sumSort = { key: 'n', dir: -1 }; drawSummary(); }; });
  $('sq').addEventListener('input', debounce(() => { state.sumQ = $('sq').value; drawSummary(); }, 150));
  $('stbl').addEventListener('click', e => {
    const th = e.target.closest('.th');
    if (th) {
      const k = th.dataset.s, s = state.sumSort;
      state.sumSort = { key: k, dir: s.key === k ? -s.dir : (k === 'name' || k === 'branch' ? 1 : -1) };
      return drawSummary();
    }
    const a = e.target.closest('[data-drill]');
    if (a) drill(a.dataset.drill, a.dataset.v, a.dataset.b, a.dataset.c);
  });

  // Setting
  $('setform').addEventListener('submit', e => { e.preventDefault(); saveMargin(); });
  $('mval').addEventListener('input', () => $('mval').classList.remove('bad'));
  buildWajibForm();
  $('wjform').addEventListener('submit', e => { e.preventDefault(); saveWajib(); });
  $('wjlist').addEventListener('change', () => { state.wjDirty = !!state.data && wajibChecked() !== state.data.wajib; setNote('wjmsg', ''); syncWajibForm(); });
  document.querySelectorAll('[data-setpage]').forEach(b => b.addEventListener('click', () => setSettingPage(b.dataset.setpage)));
  $('wilayah-user-q').addEventListener('input', debounce(renderWilayahUsers, 120));
  $('wilayah-branch-q').addEventListener('input', debounce(renderWilayahEditor, 120));
  $('wilayah-all').onclick = () => setWilayahAll(true);
  $('wilayah-none').onclick = () => setWilayahAll(false);
  $('wilayah-save').onclick = saveWilayah;

  // Unggah
  $('file').onchange = e => pickFile(e.target.files[0]);
  const drop = $('drop');
  ['dragover', 'dragenter'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => pickFile(e.dataTransfer.files[0]));
  $('upgo').onclick = sendUpload;

  // Pintasan: "/" langsung ke kolom cari; tab dibuka kembali setelah lama = cek pembaruan data
  document.addEventListener('keydown', e => {
    if (e.key === '/' && state.user && !$('p-dashboard').hidden && state.dashboardPage === 'dash' && !$('dbody').hidden && !e.target.closest('input,select,textarea,dialog')) { e.preventDefault(); $('dq').focus(); }
  });
  document.addEventListener('visibilitychange', () => {
    // Begitu aplikasi kembali aktif, periksa versi/ruang lingkup di server tanpa menunggu interval.
    // Pemicu ini menggunakan penghitung yang sama dengan interval berkala.
    if (!document.hidden && state.user) {
      if (!$('p-dashboard').hidden && state.dashboardPage === 'move') {
        if (!state.rekeningLoading) loadRekeningBergerak({ automatic: true });
      } else if (!$('p-dashboard').hidden && state.dashboardPage === 'progress') {
        if (!state.progressLoading) loadProgresHarian({ automatic: true });
      } else if (!$('p-dashboard').hidden && state.dashboardPage === 'analisa') {
        if (!state.analisaLoading) loadAnalisaCabang({ automatic: true });
      } else if (state.data && !state.loading) loadDash({ automatic: true });
    }
  });
  // Jika sesi dihapus dari tab lain (mis. batas 3 sinkron otomatis tercapai), hentikan tab ini juga.
  window.addEventListener('storage', e => {
    if (e.key !== SESSION_KEY || !state.user) return;
    let next = null;
    try { next = e.newValue ? JSON.parse(e.newValue) : null; } catch (err) { next = null; }
    if (!next || next.token !== state.sessionToken) {
      logout('Sesi telah berakhir di tab lain. Silakan login kembali.', true);
    }
  });
}

async function init() {
  bind();
  try { indexedDB.deleteDatabase(CACHE_DB_OLD); } catch (e) { /* abaikan */ }   // salinan data versi lama tidak dipakai lagi
  const s = session.get();
  if (!s || !s.token || !(s.expires > Date.now())) { session.clear(); return showView('login'); }
  if (!Number.isInteger(Number(s.autoSyncCount)) || Number(s.autoSyncCount) < 0) {
    s.autoSyncCount = 0; session.set(s); // kompatibilitas sesi yang dibuat sebelum batas sinkron otomatis ditambahkan
  }
  if (s.user) { enterApp(s.user); return; }                // pengguna tersimpan: langsung tampil, server memverifikasi sambil memuat data
  const r = await api('webMe', { token: s.token }, { retry: true });   // sesi dari versi lama (belum menyimpan pengguna)
  if (r.ok) { session.set(Object.assign(s, { user: r.user })); return enterApp(r.user); }
  if (r.code === 'SESSION_REQUIRED') session.clear();
  showView('login', r.error);
}
init();
})();
