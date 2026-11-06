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
const SESSION_KEY = 'dash_sesi_v2'; // v2: sesi lama diinvalidasi satu kali untuk menerapkan alur ganti password awal.
const LEGACY_SESSION_KEY = 'dash_sesi_v1';
const CACHE_DB = 'dash_cache_v4';   // v4: cache lokal direset agar data lama tidak dipakai setelah reset database operasional
const CACHE_DB_OLD = 'dash_cache_v3';   // cache sebelum reset; dihapus saat halaman dibuka
const CACHE_DB_OLDER = 'dash_cache_v2'; // cache lama; dihapus saat halaman dibuka
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
let initialPasswordChangeTicket = ''; // hanya disimpan di memori selama halaman ganti password terbuka.
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
  imageDownloadLevels: [], imageDownloadConfigLoaded: false, imageDownloadDirty: false, settingPage: 'global',
  data: null, byCat: {}, branches: new Map(), groups: {}, clientIdCounts: new Map(),
  cat: 'a', q: '', branch: '', staf: '', sort: { label: '', dir: 1 }, page: 0, size: 50, view: [],
  sumBy: 'branch', sumSort: { key: 'n', dir: -1 }, sumQ: '', wjDirty: false,
  wilayah: null, wilayahSelected: '', wilayahDraft: [], wilayahDrafts: Object.create(null), wilayahLoading: false, wilayahLoadedAt: 0,
  rekening: null, rekeningLoading: false, rekeningLoadedAt: 0, dashboardPage: 'move',
  progress: null, progressLoading: false, progressLoadedAt: 0, progressReloadPending: false,
  progressFilters: { branch: '', monthKey: '', progressDate: '', coordinatorNik: '' }, progressFilterRevision: 0,
  progressRecap: 'new', progressDetailPage: 1, progressDetailSeq: 0, progressDetailResult: null,
  progressDetailCache: new Map(), progressDetailPending: new Map(),
  analisaCabang: null, analisaLoading: false, analisaLoadedAt: 0, analisaFilters: { branch: '*', period: '', compare: '' },
  newsProgress: null, newsProgressLoadedAt: 0, newsLoading: false, newsPromise: null, newsLoadSeq: 0, newsReady: false,
  rbView: 'branch', rbSearch: { branch: '', coordinator: '' },
  rbSort: { branch: { key: 'name', dir: 1 }, coordinator: { key: 'name', dir: 1 } },
};

function showView(name, msg) {
  $('splash').hidden = name !== 'splash';
  $('login').hidden = name !== 'login';
  $('password-change').hidden = name !== 'password-change';
  $('app').hidden = name !== 'app';
  if (name === 'login') { showErr(msg); $(msg ? 'pw' : 'nik').focus(); }
  if (name === 'password-change') { showPasswordChangeErr(msg || ''); $('force-new-password').focus(); }
}
function showErr(msg) { $('err').hidden = !msg; $('err').textContent = msg || ''; }

const canSetting = u => !!u && (u.tingkat === 'ADMINISTRATOR' || u.tingkat === 'USER 0');   // menu Setting
const TABS = ['dashboard', 'set', 'up', 'versions'];
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
    if (state.user && canDownloadImage()) prepareRegionalNewsData();
    else if (state.user && !state.data && !state.loading) loadDash();
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
  paintImageDownloadControls();
  paintImageDownloadSettings();
  $('t-set').hidden = !canSetting(user);
  if (!canSetting(user) && !$('p-set').hidden) selectTab('dashboard');
  $('t-up').hidden = user.tingkat !== 'ADMINISTRATOR';   // hanya tampilan; penegakan sebenarnya di server
  const mainNav = document.querySelector('#app>.tabs');
  if (mainNav) mainNav.style.setProperty('--nav-count', String(TABS.filter(k => !$('t-' + k).hidden).length));
}

const IMAGE_DOWNLOAD_USER_LEVELS = ['USER 0', 'USER 1', 'USER 2', 'USER 3'];
function canDownloadImage() {
  if (!state.user) return false;
  if (state.user.tingkat === 'ADMINISTRATOR') return true;
  return state.imageDownloadConfigLoaded && state.imageDownloadLevels.indexOf(state.user.tingkat) > -1;
}
function paintImageDownloadControls() {
  const allowed = canDownloadImage();
  document.querySelectorAll('.image-download-btn').forEach(button => { button.hidden = !allowed; });
  const admin = !!state.user && state.user.tingkat === 'ADMINISTRATOR';
  const nav = $('setnav-enable-download');
  if (nav) nav.hidden = !admin;
  if (!admin && state.settingPage === 'enable-download') setSettingPage('global');
  if (allowed && state.dashboardPage === 'sum' && state.user && !state.newsLoading) {
    Promise.resolve().then(() => prepareRegionalNewsData());
  }
}
function imageDownloadChecked() {
  return IMAGE_DOWNLOAD_USER_LEVELS.filter(level => {
    const input = document.querySelector('#image-download-level-list input[value="' + level + '"]');
    return !!(input && input.checked);
  });
}
function paintImageDownloadSettings() {
  if (!$('image-download-level-list')) return;
  const current = state.imageDownloadLevels || [];
  if (!state.imageDownloadDirty) {
    document.querySelectorAll('#image-download-level-list input').forEach(input => { input.checked = current.indexOf(input.value) > -1; });
  }
  $('image-download-current').textContent = current.length ? 'Administrator, ' + current.join(', ') : 'Administrator saja';
  const same = imageDownloadChecked().join('|') === current.join('|');
  $('image-download-save').disabled = !state.imageDownloadConfigLoaded || same || !state.user || state.user.tingkat !== 'ADMINISTRATOR';
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
  initialPasswordChangeTicket = '';
  state.loadSeq++; state.loading = false; state.user = null; state.sessionToken = ''; state.data = null; state.byCat = {}; state.groups = {}; state.clientIdCounts = new Map(); state.wjDirty = false;
  state.imageDownloadLevels = []; state.imageDownloadConfigLoaded = false; state.imageDownloadDirty = false; state.settingPage = 'global';
  session.clear();
  const cacheClear = cache.clear();
  state.wilayah = null; state.wilayahDraft = []; state.wilayahDrafts = Object.create(null); state.wilayahSelected = ''; setNote('wjmsg', '');
  state.rekening = null; state.rekeningLoadedAt = 0; state.rekeningLoading = false;
  state.progress = null; state.progressLoadedAt = 0; state.progressLoading = false; state.progressReloadPending = false;
  state.progressRecap = 'new'; state.progressDetailPage = 1; state.progressDetailSeq++; state.progressDetailResult = null;
  state.progressDetailCache.clear(); state.progressDetailPending.clear();
  state.progressFilters = { branch: '', monthKey: '', progressDate: '', coordinatorNik: '' }; state.progressFilterRevision++;
  state.analisaCabang = null; state.analisaLoadedAt = 0; state.analisaLoading = false;
  invalidateRegionalNewsData();
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

  // Akun dengan password awal hanya menerima tiket terbatas; belum memiliki sesi Dashboard.
  if (r.mustChangePassword === true) {
    if (!r.changeToken) return showErr('Server tidak memberikan sesi penggantian password. Hubungi Administrator.');
    initialPasswordChangeTicket = r.changeToken;
    $('pw').value = '';
    $('force-new-password').value = '';
    $('force-confirm-password').value = '';
    showPasswordChangeErr('');
    showView('password-change');
    return;
  }

  if (!r.token || !r.user) return showErr('Balasan login tidak lengkap. Silakan coba kembali.');
  session.set({ token: r.token, expires: r.expires, user: r.user, autoSyncCount: 0 });
  $('pw').value = '';
  enterApp(r.user);
}

function showPasswordChangeErr(msg) {
  const el = $('force-password-err');
  if (!el) return;
  el.hidden = !msg;
  el.textContent = msg || '';
}

async function submitInitialPasswordChange(ev) {
  ev.preventDefault();
  const newPass = $('force-new-password').value;
  const confirmPass = $('force-confirm-password').value;
  $('force-new-password').classList.toggle('bad', !newPass || newPass.length < 8);
  $('force-confirm-password').classList.toggle('bad', !confirmPass || newPass !== confirmPass);
  if (!initialPasswordChangeTicket) {
    showView('login', 'Sesi penggantian password tidak tersedia. Silakan login kembali.');
    return;
  }
  if (newPass.length < 8) return showPasswordChangeErr('Password baru minimal 8 karakter.');
  if (newPass.length > 200) return showPasswordChangeErr('Password baru maksimal 200 karakter.');
  if (!confirmPass || newPass !== confirmPass) return showPasswordChangeErr('Konfirmasi password baru tidak sama.');

  showPasswordChangeErr('');
  const btn = $('force-password-go');
  btn.disabled = true;
  btn.replaceChildren(h('span', { class: 'spin' }), h('span', { text: 'Menyimpan…' }));
  const r = await api('webChangeInitialPassword', {
    changeToken: initialPasswordChangeTicket,
    newPassword: newPass,
    confirmPassword: confirmPass
  });
  btn.disabled = false;
  btn.textContent = 'Simpan Password Baru';

  if (!r.ok) {
    if (r.code === 'PASSWORD_CHANGE_REQUIRED') {
      initialPasswordChangeTicket = '';
      return showView('login', r.error || 'Sesi penggantian password berakhir. Silakan login kembali.');
    }
    return showPasswordChangeErr(r.error || 'Password baru gagal disimpan. Silakan coba kembali.');
  }
  if (!r.token || !r.user) return showPasswordChangeErr('Password mungkin sudah tersimpan, tetapi sesi login tidak lengkap. Silakan login menggunakan password baru.');

  initialPasswordChangeTicket = '';
  $('force-new-password').value = '';
  $('force-confirm-password').value = '';
  session.set({ token: r.token, expires: r.expires, user: r.user, autoSyncCount: 0 });
  enterApp(r.user);
  toast('Password berhasil diperbarui. Selamat datang, ' + (r.user.nama || r.user.nik) + '.');
}

/* ---------------------------------------------------------------------
   6. Dashboard
   --------------------------------------------------------------------- */
function enterApp(user) {
  state.user = user;
  state.sessionToken = (session.get() || {}).token || '';
  state.imageDownloadLevels = []; state.imageDownloadConfigLoaded = false; state.imageDownloadDirty = false; state.settingPage = 'global';
  paintUser(user);
  state.cat = 'a'; state.q = ''; state.branch = ''; state.staf = ''; state.page = 0; state.sort = { label: '', dir: 1 };
  resetUpload(); setNote('setmsg', ''); setNote('wjmsg', ''); setNote('rb-setting-msg', ''); state.wjDirty = false;
  state.rekening = null; state.rekeningLoadedAt = 0;
  state.progress = null; state.progressLoadedAt = 0; state.progressLoading = false; state.progressReloadPending = false;
  state.progressRecap = 'new'; state.progressDetailPage = 1; state.progressDetailSeq++; state.progressDetailResult = null;
  state.progressDetailCache.clear(); state.progressDetailPending.clear();
  state.progressFilters = { branch: '', monthKey: '', progressDate: '', coordinatorNik: '' }; state.progressFilterRevision++;
  $('ph-filter-month').value = ''; $('ph-filter-branch').value = ''; $('ph-filter-date').value = ''; $('ph-filter-user').value = '';
  state.analisaCabang = null; state.analisaLoadedAt = 0; state.analisaLoading = false;
  state.analisaFilters = { branch: '*', period: '', compare: '' };
  invalidateRegionalNewsData();
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
  $('dmeta').textContent = 'Memuat data…'; $('ddl').disabled = true; $('ddl-img').disabled = true; $('dref').disabled = false; setSync(''); setNote('dmsg', '');
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
  state.imageDownloadConfigLoaded = Array.isArray(r.imageDownloadLevels);
  state.imageDownloadLevels = state.imageDownloadConfigLoaded
    ? IMAGE_DOWNLOAD_USER_LEVELS.filter(level => r.imageDownloadLevels.indexOf(level) > -1) : [];
  paintImageDownloadControls();
  paintImageDownloadSettings();
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
  $('ddl').disabled = !d.scanned; $('ddl-img').disabled = !d.scanned;
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
  document.querySelectorAll('#p-sum .seg button[data-g]').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.g === kind)));

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
  const baseline = !!c.initialBaseline;
  const allChecked = Number(c.movementCheckedLoanCount || 0);
  put($('updiffstats'),
    chip(fmt(c.uploadCount || 0) + ' LOAN NO pada Excel'),
    chip(fmt(c.currentCount || 0) + ' LOAN NO pada Data_Upload'),
    baseline ? chip('BASELINE AWAL · bukan PAR Baru', true) : chip(fmt(allChecked) + ' LOAN NO unik diperiksa untuk Total Balance'),
    !baseline ? chip(fmt(c.matchedCount || 0) + ' Loan No cocok · Total Balance dibandingkan') : null,
    !baseline ? chip(fmt(c.firstSeenCount || 0) + ' Loan No hanya di Excel · perubahan 0') : null,
    !baseline ? chip(fmt(c.missingFromNewUploadCount || 0) + ' Loan No hanya di Data_Upload · riwayat dipertahankan') : null,
    chip(fmt(c.newCount || 0) + ' PAR Baru', !baseline && (c.newCount || 0) > 0),
    chip(fmt(c.doneCount || 0) + ' PAR Selesai', !baseline && (c.doneCount || 0) > 0),
    baseline ? chip(fmt(c.baselineCount || c.uploadCount || 0) + ' Loan No disimpan sebagai data awal') : null,
    c.progressDate ? chip('Tanggal Progress: ' + c.progressDate) : null);
  $('updiff-new-samples').textContent = baseline
    ? 'Tidak dihitung pada unggahan baseline awal.'
    : ((c.newSamples && c.newSamples.length) ? c.newSamples.join(' · ') : 'Tidak ditemukan LOAN NO baru.');
  $('updiff-done-samples').textContent = baseline
    ? 'Belum ada data lama untuk dibandingkan.'
    : ((c.doneSamples && c.doneSamples.length) ? c.doneSamples.join(' · ') : 'Tidak ditemukan LOAN NO selesai.');
  // Tabel audit hanya menampilkan Total Balance; seluruh audit selisih simpanan dinonaktifkan.
  const auditWrap = $('updiff-audit-wrap'), auditTbl = $('updiff-audit');
  const audit = !baseline && Array.isArray(c.movementAudit) ? c.movementAudit : [];
  if (auditWrap && auditTbl) {
    auditWrap.hidden = !audit.length;
    if (audit.length) {
      const num = v => h('td', { class: 'num', text: fmt(v) });
      const heads = ['Nominal', 'Data_Upload lama', 'Excel baru', 'Selisih (lama − baru)', 'Loan turun', 'Loan naik', 'Loan tetap'];
      auditTbl.replaceChildren(
        h('thead', null, h('tr', null, heads.map((t, i) => h('th', { scope: 'col', class: i ? 'num' : '', text: t })))),
        h('tbody', null, audit.map(a => h('tr', null, [h('td', { text: a.label }), num(a.prev), num(a.now), num(a.delta), num(a.down), num(a.up), num(a.same)])))
      );
    }
  }
  const grid = $('updiff-grid');
  if (grid) grid.hidden = baseline;
  const help = $('updiff-help');
  if (help) help.textContent = baseline
    ? 'Data_Upload belum memiliki data. File ini akan disimpan sebagai BASELINE AWAL. Seluruh Loan No menjadi data pembanding; tidak ada baris yang ditambahkan ke PAR Baru, PAR Selesai, Rekening Bergerak, atau Total Balance Penurunan. Pembandingan baru dimulai pada unggahan PAR berikutnya.'
    : 'Pemeriksaan membandingkan Total Balance untuk Loan No yang cocok pada Data_Upload lama dan Excel baru, tanpa memfilter PAR Baru/Selesai. Proses mencari selisih semua jenis simpanan sedang dinonaktifkan. Loan No yang hanya ada di salah satu snapshot tetap dicatat dengan perubahan Total Balance 0. Data_Upload lama belum diubah. Periksa hasil di atas, lalu konfirmasi untuk menyimpan arsip dan mengganti Data_Upload.';
  $('updiff').hidden = false;
}

async function pickFile(file) {
  if (!file) return;
  resetUpload(); $('upinfo').hidden = false; $('upgo').disabled = true; setNote('upmsg', 'Memeriksa nama dan membaca file…');
  await nextFrame();
  try {
    const info = parseUploadFilename(file.name);
    pending = info.type === 'ANALISA' ? await readAnalisaExcel(file, info) : await readExcel(file);
    if (info.type === 'PAR') {
      pending.type = 'PAR';
      pending.clientKey = [file.name, file.size, file.lastModified].join('|');
    }
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
  const total = Number(pending && pending.totalRows || sent || 0);
  const stagedNote = sent
    ? (Number(sent) >= total
      ? ' Staging menerima seluruh ' + fmt(sent) + ' dari ' + fmt(total) + ' baris. Data_Upload lama tetap aman. Tekan ulang Periksa LOAN NO untuk melanjutkan pemeriksaan pada sesi yang sama; jangan mulai upload baru.'
      : ' Staging menerima ' + fmt(sent) + ' dari ' + fmt(total) + ' baris. Data_Upload lama tetap aman; tekan ulang untuk melanjutkan pengiriman yang tersisa.')
    : '';
  setNote('upmsg', (r.error || 'Gagal memproses data.') + stagedNote, 'bad');
}

async function stageAndCheck() {
  if (!pending || (!Array.isArray(pending.rows) && !pending.uploadId)) return;
  const total = Number(pending.totalRows || (pending.rows && pending.rows.length) || 0);
  if (!total) return;
  const tok = (session.get() || {}).token;
  let id = pending.uploadId || '';
  let sent = 0;

  $('upgo').disabled = true; upBar(0);
  // Endpoint Start akan mengembalikan sesi lama yang cocok (termasuk sesi v20.16 tanpa clientKey),
  // sehingga browser dapat memulihkan jumlah staging yang benar setelah timeout atau reload.
  setNote('upmsg', id ? 'Memulihkan sesi staging sebelumnya…' : 'Menyiapkan area pemeriksaan…');
  const start = await api('webUploadStart', {
    token: tok, total, file: pending.name, clientKey: pending.clientKey || ''
  });
  if (!start.ok) return upFail(start, 0);
  if (id && start.uploadId !== id) {
    return upFail({ error: 'Sesi staging berubah. Pilih ulang file yang sama untuk melanjutkan sesi yang benar.' }, 0);
  }
  id = start.uploadId;
  sent = Number(start.rows || 0);
  pending.uploadId = id; pending.totalRows = total; pending.sentRows = sent;
  upBar(sent / total);

  // Lanjutkan sisa potongan pada staging yang sama. Jika jawaban request sebelumnya terputus
  // setelah server menerima potongan, endpoint akan mengembalikan jumlah baris yang sudah diterima.
  while (sent < total) {
    const seq = Math.floor(sent / UP_CHUNK);
    const part = pending.rows.slice(sent, sent + UP_CHUNK);
    let good = false, res;
    for (let t = 0; t < 3 && !good; t++) {
      res = await api('webUploadChunk', {
        token: tok, uploadId: id, seq, rows: JSON.stringify(part)
      });
      good = !!res.ok;
      if (!good) {
        if (res.code === 'SESSION_REQUIRED') break;
        await sleep(1500 * (t + 1));
      }
    }
    if (!good) {
      pending.uploadId = id; pending.totalRows = total; pending.sentRows = sent;
      return upFail(res, sent);
    }
    sent = Math.max(sent + part.length, Number(res.rows || 0));
    pending.uploadId = id; pending.totalRows = total; pending.sentRows = sent;
    upBar(sent / total);
    setNote('upmsg', 'Menyiapkan pemeriksaan ' + fmt(sent) + ' dari ' + fmt(total) + ' baris…');
  }

  pending.uploadId = id; pending.totalRows = total; pending.sentRows = total;
  setNote('upmsg', 'Membandingkan seluruh LOAN NO dengan Data_Upload…');
  // Pemeriksaan server dijalankan per batch. Setiap permintaan menyimpan checkpoint sehingga
  // tombol ulang melanjutkan offset terakhir, bukan mengunggah ulang 20 ribu baris.
  let r = await api('webUploadCheck', { token: tok, uploadId: id, total }, { timeoutMs: 120000 });
  let guard = 0;
  while (r.ok && r.checking && guard++ < 200) {
    const pct = Math.max(0, Math.min(1, Number(r.progress || 0)));
    upBar(pct);
    setNote('upmsg', (r.message || 'Pemeriksaan Total Balance berjalan…') + ' ' + fmt(r.checkedRows || 0) + ' dari ' + fmt(r.totalRows || total) + ' baris.');
    await nextFrame();
    r = await api('webUploadCheck', { token: tok, uploadId: id, total }, { timeoutMs: 120000 });
  }
  if (!r.ok) {
    pending.uploadId = id; pending.totalRows = total; pending.sentRows = total;
    return upFail(r, total);
  }
  if (r.checking || !r.checked) {
    pending.uploadId = id; pending.totalRows = total; pending.sentRows = total;
    return upFail({ error: 'Pemeriksaan masih berjalan. Tekan ulang Periksa LOAN NO untuk melanjutkan batch berikutnya.' }, total);
  }
  pending.uploadId = id; pending.totalRows = total; pending.sentRows = total;
  pending.checked = true; pending.compare = r.compare || {};
  pending.rows = null; // file sudah aman di sheet staging; bebaskan memori browser
  renderUploadCompare(pending.compare);
  upBar(null);
  $('upgo').textContent = 'Konfirmasi & Upload ke Data_Upload'; $('upgo').disabled = false;
  setNote('upmsg', pending.compare.initialBaseline
    ? 'Pemeriksaan selesai. Data_Upload belum memiliki data; file ini akan dijadikan baseline awal dan tidak dihitung sebagai PAR Baru/Selesai maupun Rekening Bergerak.'
    : 'Pemeriksaan selesai. ' + fmt(pending.compare.movementCheckedLoanCount || 0) + ' Loan No unik diperiksa untuk Total Balance; ' + fmt(pending.compare.matchedCount || 0) + ' cocok, ' + fmt(pending.compare.firstSeenCount || 0) + ' hanya pada Excel, dan ' + fmt(pending.compare.missingFromNewUploadCount || 0) + ' hanya pada Data_Upload. Selisih simpanan tidak diperiksa. Data_Upload lama masih utuh dan belum diganti.', 'good');
}

async function commitUpload() {
  if (!pending || !pending.checked || !pending.uploadId) return;
  const c = pending.compare || {};
  const baseline = !!c.initialBaseline;
  const confirmText = baseline
    ? 'Data_Upload saat ini hanya memiliki header. File berisi ' + fmt(c.baselineCount || c.uploadCount || pending.totalRows) + ' Loan No ini akan disimpan sebagai BASELINE AWAL. Tidak ada data yang ditambahkan ke PAR Baru, PAR Selesai, Rekening Bergerak, atau Total Balance Penurunan. Unggahan PAR berikutnya akan dibandingkan dengan baseline ini. Lanjutkan menyimpan ' + fmt(pending.totalRows) + ' baris ke Data_Upload?'
    : 'Pemeriksaan Total Balance mencakup ' + fmt(c.movementCheckedLoanCount || 0) + ' Loan No unik dari gabungan Data_Upload dan Excel baru (tanpa filter PAR). Cocok: ' + fmt(c.matchedCount || 0) + '; hanya di Excel: ' + fmt(c.firstSeenCount || 0) + ' (baseline Total Balance); hanya di Data_Upload: ' + fmt(c.missingFromNewUploadCount || 0) + ' (riwayat Total Balance tetap, perubahan 0). Hanya Total Balance yang dibandingkan dengan rumus Data_Upload lama dikurangi Excel baru; seluruh selisih simpanan dinonaktifkan. PAR Baru: ' + fmt(c.newCount || 0) + ' baris. PAR Selesai: ' + fmt(c.doneCount || 0) + ' baris. Tanggal Progress: ' + (c.progressDate || '-') + '. Setelah semua hasil dicatat dan divalidasi, Data_Upload akan diganti dengan ' + fmt(pending.totalRows) + ' baris dari Excel.';
  const ok = await askConfirm(baseline ? 'Simpan data awal (baseline)?' : 'Konfirmasi hasil pemeriksaan?', confirmText,
    baseline ? 'Simpan sebagai baseline' : 'Simpan arsip & upload');
  if (!ok) return;
  $('upgo').disabled = true; upBar(null); setNote('upmsg', baseline
    ? 'Menyimpan file sebagai baseline awal dan memperbarui Data_Upload…'
    : 'Menyimpan PAR Baru dan PAR Selesai, lalu memperbarui Data_Upload…');
  const tok = (session.get() || {}).token;
  const r = await api('webUploadCommit', { token: tok, uploadId: pending.uploadId }, { timeoutMs: 300000 });
  if (!r.ok) {
    upBar(null); $('upgo').disabled = false;
    $('upgo').textContent = 'Ulangi Konfirmasi & Upload';
    return setNote('upmsg', r.error || 'Penyimpanan belum selesai. Tekan ulang untuk melanjutkan proses.', 'bad');
  }
  upBar(1);
  setNote('upmsg', r.initialBaseline || baseline
    ? 'Baseline awal berhasil disimpan. Data_Upload: ' + fmt(r.rows) + ' baris. PAR Baru: 0 · PAR Selesai: 0 · Rekening Bergerak: 0. File ini sekarang menjadi pembanding untuk seluruh unggahan berikutnya.'
    : 'Selesai. Data_Upload: ' + fmt(r.rows) + ' baris · PAR Baru: ' + fmt(r.newCount || c.newCount || 0) + ' baris · PAR Selesai: ' + fmt(r.doneCount || c.doneCount || 0) + ' baris · Tanggal Progress: ' + (r.progressDate || c.progressDate || '-') + '.', 'good');
  toast(r.initialBaseline || baseline ? 'Baseline awal berhasil disimpan.' : 'Pemeriksaan dan unggah Excel berhasil. Dashboard dimuat ulang.');
  // Unggahan mengubah arsip dan Data_Upload; batalkan data sub-dashboard yang tersimpan di memori.
  state.progress = null; state.progressLoadedAt = 0;
  invalidateRegionalNewsData();
  state.rekening = null; state.rekeningLoadedAt = 0;
  pending = null; $('file').value = ''; $('upgo').disabled = true;
  loadDash();
  if (state.dashboardPage === 'sum' && canDownloadImage()) prepareRegionalNewsData();
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
  invalidateRegionalNewsData();
  if (!$('p-dashboard').hidden && state.dashboardPage === 'analisa') loadAnalisaCabang();
  else if (!$('p-dashboard').hidden && state.dashboardPage === 'sum' && canDownloadImage()) prepareRegionalNewsData();
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
   11A. Ekspor Ringkasan Dashboard ke Gambar PNG ukuran F4 (portrait)
   Dibuat secara lokal dari data dashboard yang sudah diterima browser.
   --------------------------------------------------------------------- */
const IMG_F4_W = 1654;       // 210 mm pada 200 dpi
const IMG_F4_H = 2598;       // 330 mm pada 200 dpi
const IMG_F4_M = 76;
const IMG_F4_CW = IMG_F4_W - IMG_F4_M * 2;

function imgRoundRect(ctx, x, y, w, h, r, fill, stroke) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
  if (fill) { ctx.fillStyle = fill; ctx.fill(); }
  if (stroke) { ctx.strokeStyle = stroke; ctx.lineWidth = 1.5; ctx.stroke(); }
}

function imgFitText(ctx, value, maxWidth) {
  const text = String(value == null ? '' : value);
  if (ctx.measureText(text).width <= maxWidth) return text;
  let out = text;
  while (out.length > 1 && ctx.measureText(out + '…').width > maxWidth) out = out.slice(0, -1);
  return out + '…';
}

function imgWrapText(ctx, value, x, y, maxWidth, lineHeight, maxLines, color) {
  const text = String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
  if (!text) return y;
  const words = text.split(' '), lines = [];
  let line = '';
  words.forEach(word => {
    const test = line ? line + ' ' + word : word;
    if (line && ctx.measureText(test).width > maxWidth) { lines.push(line); line = word; }
    else line = test;
  });
  if (line) lines.push(line);
  const take = Math.max(1, maxLines || 2);
  if (lines.length > take) {
    lines.length = take;
    lines[take - 1] = imgFitText(ctx, lines[take - 1] + ' ' + words.slice(lines.join(' ').split(' ').length).join(' '), maxWidth);
  }
  ctx.fillStyle = color || '#53665D';
  lines.forEach((lineText, i) => ctx.fillText(imgFitText(ctx, lineText, maxWidth), x, y + i * lineHeight));
  return y + lines.length * lineHeight;
}

function imgPrepare(title, subtitle) {
  const canvas = document.createElement('canvas');
  canvas.width = IMG_F4_W; canvas.height = IMG_F4_H;
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('Browser tidak dapat menyiapkan kanvas gambar.');
  ctx.fillStyle = '#F3F7F4'; ctx.fillRect(0, 0, IMG_F4_W, IMG_F4_H);
  // Header utama
  ctx.fillStyle = '#064E3B'; ctx.fillRect(0, 0, IMG_F4_W, 214);
  ctx.fillStyle = '#D9A928'; ctx.fillRect(IMG_F4_M, 48, 12, 36);
  ctx.font = '700 21px Arial, sans-serif'; ctx.fillStyle = '#BFE6D4';
  ctx.fillText('DASHBOARD OPERASIONAL SATU PINTU', IMG_F4_M + 32, 71);
  ctx.font = '700 48px Arial, sans-serif'; ctx.fillStyle = '#FFFFFF';
  ctx.fillText(imgFitText(ctx, title, IMG_F4_CW), IMG_F4_M, 137);
  ctx.font = '18px Arial, sans-serif'; ctx.fillStyle = '#E0F2E9';
  ctx.fillText(imgFitText(ctx, subtitle || '', IMG_F4_CW), IMG_F4_M, 177);
  return { canvas, ctx, y: 244 };
}

function imgSection(ctx, title, subtitle, y, x, width) {
  x = x == null ? IMG_F4_M : x; width = width || IMG_F4_CW;
  ctx.font = '700 23px Arial, sans-serif'; ctx.fillStyle = '#123D2F';
  ctx.fillText(imgFitText(ctx, title, width), x, y + 25);
  if (subtitle) {
    ctx.font = '15px Arial, sans-serif'; ctx.fillStyle = '#6B7E74';
    ctx.fillText(imgFitText(ctx, subtitle, width), x, y + 51);
  }
  ctx.fillStyle = '#D9A928'; ctx.fillRect(x, y + 63, 58, 5);
  return y + 82;
}

function imgCards(ctx, cards, y, options) {
  options = options || {};
  const cols = Math.max(1, options.cols || 2), gap = options.gap == null ? 18 : options.gap;
  const cardH = options.cardH || 148, x = options.x == null ? IMG_F4_M : options.x;
  const width = options.width || IMG_F4_CW, cardW = (width - gap * (cols - 1)) / cols;
  cards.forEach((card, i) => {
    const row = Math.floor(i / cols), col = i % cols;
    const cx = x + col * (cardW + gap), cy = y + row * (cardH + gap);
    imgRoundRect(ctx, cx, cy, cardW, cardH, 16, i % 2 ? '#FFFDF6' : '#FFFFFF', '#DCE8E0');
    ctx.fillStyle = i % 2 ? '#D9A928' : '#0A7A5A';
    ctx.fillRect(cx + 18, cy + 18, 6, 27);
    ctx.font = '700 16px Arial, sans-serif'; ctx.fillStyle = '#61746A';
    ctx.fillText(imgFitText(ctx, card.label || '', cardW - 54), cx + 36, cy + 38);
    ctx.font = '700 ' + (options.valueFont || 29) + 'px Arial, sans-serif'; ctx.fillStyle = '#123D2F';
    ctx.fillText(imgFitText(ctx, card.value == null ? '—' : card.value, cardW - 40), cx + 20, cy + 82);
    ctx.font = '14px Arial, sans-serif';
    imgWrapText(ctx, card.note || '', cx + 20, cy + 111, cardW - 40, 18, 2, '#6B7E74');
  });
  const rows = Math.ceil(cards.length / cols);
  return y + rows * cardH + Math.max(0, rows - 1) * gap;
}

function imgTable(ctx, y, title, headers, rows, options) {
  options = options || {};
  const x = options.x == null ? IMG_F4_M : options.x;
  const width = options.width || IMG_F4_CW;
  const rowH = options.rowH || 42;
  const fontSize = options.fontSize || 15;
  const headerH = options.headerH || 48;
  const maxRows = options.maxRows == null ? rows.length : options.maxRows;
  const subset = rows.slice(0, maxRows);
  if (title) y = imgSection(ctx, title, options.subtitle || '', y, x, width);
  const ratios = options.widths && options.widths.length === headers.length
    ? options.widths : headers.map(() => 1);
  const ratioTotal = ratios.reduce((a, b) => a + b, 0);
  const widths = ratios.map(v => width * v / ratioTotal);
  let cx = x;
  ctx.save();
  imgRoundRect(ctx, x, y, width, headerH, 9, '#087A5A');
  headers.forEach((header, i) => {
    ctx.font = '700 ' + (options.headerFont || fontSize) + 'px Arial, sans-serif';
    ctx.fillStyle = '#FFFFFF';
    const right = options.rightCols && options.rightCols.includes(i);
    const tx = right ? cx + widths[i] - 12 : cx + 12;
    ctx.textAlign = right ? 'right' : 'left';
    ctx.fillText(imgFitText(ctx, header, widths[i] - 24), tx, y + headerH / 2 + 6);
    cx += widths[i];
  });
  ctx.textAlign = 'left';
  let ry = y + headerH;
  if (!subset.length) {
    ctx.fillStyle = '#FFFFFF'; ctx.fillRect(x, ry, width, rowH);
    ctx.font = '15px Arial, sans-serif'; ctx.fillStyle = '#6B7E74';
    ctx.fillText(imgFitText(ctx, options.emptyText || 'Tidak ada data.', width - 24), x + 12, ry + rowH / 2 + 5);
    ry += rowH;
  } else {
    subset.forEach((row, ri) => {
      const cells = Array.isArray(row) ? row : row.cells || [];
      const totalRow = !!(row && row.total);
      ctx.fillStyle = totalRow ? '#E3F3EA' : (ri % 2 ? '#F2F7F4' : '#FFFFFF');
      ctx.fillRect(x, ry, width, rowH);
      ctx.strokeStyle = '#E1EAE4'; ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(x, ry + rowH); ctx.lineTo(x + width, ry + rowH); ctx.stroke();
      let cellX = x;
      cells.forEach((value, i) => {
        const cw = widths[i] || 0;
        const right = options.rightCols && options.rightCols.includes(i);
        ctx.font = (totalRow ? '700 ' : '') + fontSize + 'px Arial, sans-serif';
        ctx.fillStyle = totalRow ? '#064E3B' : '#253D32';
        ctx.textAlign = right ? 'right' : 'left';
        const tx = right ? cellX + cw - 10 : cellX + 10;
        ctx.fillText(imgFitText(ctx, value == null ? '' : value, Math.max(8, cw - 20)), tx, ry + rowH / 2 + fontSize * 0.34);
        cellX += cw;
      });
      ry += rowH;
    });
  }
  ctx.restore();
  return ry + 8;
}

function imgFooter(ctx, note) {
  const y = IMG_F4_H - 84;
  ctx.fillStyle = '#D7E4DC'; ctx.fillRect(IMG_F4_M, y, IMG_F4_CW, 2);
  ctx.font = '14px Arial, sans-serif'; ctx.fillStyle = '#61746A';
  ctx.fillText(imgFitText(ctx, note || 'Ringkasan dibuat dari data Dashboard Operasional Satu Pintu.', IMG_F4_CW), IMG_F4_M, y + 27);
  ctx.font = '13px Arial, sans-serif'; ctx.fillStyle = '#819188';
  const stamp = new Date().toLocaleString('id-ID', { dateStyle: 'medium', timeStyle: 'short' });
  ctx.fillText('Dibuat ' + stamp + ' · Format F4 portrait · 210 × 330 mm', IMG_F4_M, y + 52);
}

function imgCompactMoney(value) { return rpShort(Number(value) || 0); }

function imgPngCrc32(bytes) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) {
    crc ^= bytes[i];
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xEDB88320 : 0);
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

// Canvas PNG secara default menyimpan resolusi 96 DPI. Sisipkan/perbarui pHYs
// agar file Progres Harian mempertahankan ukuran cetak F4 landscape pada 300 DPI.
async function imgPngSetDpi(blob, dpi) {
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const signature = [137, 80, 78, 71, 13, 10, 26, 10];
  if (bytes.length < 33 || !signature.every((v, i) => bytes[i] === v)) {
    throw new Error('Format PNG tidak valid; metadata resolusi tidak dapat dipasang.');
  }
  const chunks = [];
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = 8, hasPhys = false;
  while (offset + 12 <= bytes.length) {
    const length = view.getUint32(offset, false);
    const end = offset + 12 + length;
    if (end > bytes.length) throw new Error('Struktur PNG tidak lengkap.');
    const type = String.fromCharCode(bytes[offset + 4], bytes[offset + 5], bytes[offset + 6], bytes[offset + 7]);
    chunks.push({ type, bytes: bytes.slice(offset, end) });
    if (type === 'pHYs') hasPhys = true;
    offset = end;
    if (type === 'IEND') break;
  }
  const pixelsPerMeter = Math.round(Number(dpi) / 0.0254);
  const data = new Uint8Array(9), dataView = new DataView(data.buffer);
  dataView.setUint32(0, pixelsPerMeter, false);
  dataView.setUint32(4, pixelsPerMeter, false);
  data[8] = 1; // unit = meter
  const crcInput = new Uint8Array(13);
  crcInput.set([112, 72, 89, 115], 0); // pHYs
  crcInput.set(data, 4);
  const chunk = new Uint8Array(21), chunkView = new DataView(chunk.buffer);
  chunkView.setUint32(0, 9, false);
  chunk.set(crcInput.subarray(0, 4), 4);
  chunk.set(data, 8);
  chunkView.setUint32(17, imgPngCrc32(crcInput), false);

  const output = [bytes.slice(0, 8)];
  let inserted = false;
  chunks.forEach(item => {
    if (item.type === 'pHYs') {
      if (!inserted) output.push(chunk);
      inserted = true;
      return;
    }
    output.push(item.bytes);
    if (item.type === 'IHDR' && !hasPhys && !inserted) {
      output.push(chunk);
      inserted = true;
    }
  });
  if (!inserted) throw new Error('Metadata resolusi PNG gagal dipasang.');
  return new Blob(output, { type: 'image/png' });
}

function imgDownload(canvas, filename, dpi) {
  return new Promise((resolve, reject) => {
    canvas.toBlob(blob => {
      if (!blob) return reject(new Error('Gambar gagal dibuat oleh browser.'));
      (async () => {
        const output = dpi ? await imgPngSetDpi(blob, dpi) : blob;
        const url = URL.createObjectURL(output), link = document.createElement('a');
        link.href = url; link.download = filename; document.body.appendChild(link); link.click(); link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1500);
        resolve();
      })().catch(reject);
    }, 'image/png');
  });
}

function imgBuildRekeningBergerak() {
  const data = state.rekening;
  if (!data || data.settingRequired || !data.settingDate) throw new Error('Data Rekening Bergerak belum siap atau tanggal N belum diatur.');
  const totals = (data.branchRows || []).reduce((a, r) => {
    a.total += Number(r.total) || 0; a.active += Number(r.active) || 0; return a;
  }, { total: 0, active: 0 });
  const percent = totals.total ? totals.active / totals.total * 100 : 0;
  const kind = state.rbView === 'coordinator' ? 'coordinator' : 'branch';
  const raw = kind === 'branch' ? (data.branchRows || []) : (data.coordinatorRows || []);
  const query = String(state.rbSearch[kind] || '').trim().toLocaleLowerCase('id-ID');
  const sort = state.rbSort[kind] || { key: 'name', dir: 1 };
  const visible = raw.filter(r => !query || String(kind === 'branch' ? r.branch : r.name || '').toLocaleLowerCase('id-ID').includes(query))
    .slice().sort((a, b) => {
      const av = sort.key === 'name' ? (kind === 'branch' ? a.branch || '' : a.name || '') : Number(a[sort.key]) || 0;
      const bv = sort.key === 'name' ? (kind === 'branch' ? b.branch || '' : b.name || '') : Number(b[sort.key]) || 0;
      return (typeof av === 'string' ? collator.compare(av, bv) : av - bv) * (sort.dir || 1);
    });
  const filteredTotals = visible.reduce((a, r) => { a.total += Number(r.total) || 0; a.active += Number(r.active) || 0; return a; }, { total: 0, active: 0 });
  const sub = 'Tanggal N: ' + rbDateLabel(data.settingDate) + ' · Tahun acuan: ' + (data.year || '—') + ' · ' + (data.scope || 'Sesuai Wilayah Kerja');
  const report = imgPrepare('Rekening Bergerak', sub), ctx = report.ctx;
  let y = report.y;
  y = imgCards(ctx, [
    { label: 'TOTAL REKENING PAR', value: fmt(data.totalRows == null ? totals.total : data.totalRows), note: 'Seluruh rekening pada cakupan akses' },
    { label: 'REKENING BERGERAK', value: fmt(totals.active), note: 'Transaksi pada atau setelah tanggal N' },
    { label: 'PERSENTASE BERGERAK', value: nf1.format(percent) + '%', note: 'Rekening bergerak dibanding total PAR' },
    { label: 'TANGGAL TIDAK TERISI', value: fmt(data.missingDateCount || 0), note: 'Tetap masuk Total Rekening PAR' }
  ], y, { cols: 2, cardH: 150, valueFont: 31 });
  y += 28;
  const rowLabel = kind === 'branch' ? 'Cabang' : 'Branch Coordinator';
  const rows = visible.slice(0, 34).map(r => [
    kind === 'branch' ? (r.branch || '(tanpa nama)') : (r.name || '(tanpa nama)'),
    fmt(r.total), fmt(r.active), fmt(r.deviation), nf1.format(r.percent) + '%'
  ]);
  rows.push({ total: true, cells: [query ? 'TOTAL HASIL FILTER' : 'TOTAL', fmt(filteredTotals.total), fmt(filteredTotals.active), fmt(filteredTotals.active - filteredTotals.total), nf1.format(filteredTotals.total ? filteredTotals.active / filteredTotals.total * 100 : 0) + '%'] });
  y = imgTable(ctx, y, 'RINCIAN ' + (kind === 'branch' ? 'PER CABANG' : 'PER BRANCH COORDINATOR'),
    [rowLabel, 'Total PAR', 'Bergerak', 'Deviasi', 'Bergerak (%)'], rows,
    { widths: [3.1, 1.25, 1.25, 1.15, 1.2], rowH: 39, fontSize: 15, maxRows: 35,
      subtitle: 'Menampilkan ' + fmt(visible.length) + ' dari ' + fmt(raw.length) + ' baris yang cocok dengan pencarian aktif.', rightCols: [1, 2, 3, 4] });
  if (visible.length > 34) {
    ctx.font = '14px Arial, sans-serif'; ctx.fillStyle = '#6B7E74';
    ctx.fillText('Tabel gambar dibatasi 34 baris; file Excel memuat seluruh cabang/koordinator.', IMG_F4_M, Math.min(y + 12, IMG_F4_H - 112));
  }
  const warning = $('rb-warning').hidden ? '' : $('rb-warning').textContent;
  imgFooter(ctx, warning || 'Rumus Deviasi = Rekening Bergerak − Total Rekening PAR. Persentase = Rekening Bergerak ÷ Total PAR × 100%.');
  return report.canvas;
}

function imgBuildToolPAR() {
  const d = state.data;
  if (!d || !d.scanned) throw new Error('Data Tool PAR belum siap.');
  const meta = 'Data diperiksa: ' + fmt(d.scanned) + ' pinjaman · Filter: ' + (state.branch || 'Semua cabang') + ' · ' + (state.staf || 'Semua staf');
  const report = imgPrepare('Tool PAR', meta), ctx = report.ctx;
  let y = report.y;
  y = imgCards(ctx, [
    { label: 'PINJAMAN DICEK', value: fmt(d.scanned), note: 'Total data yang diproses' },
    { label: 'MASUK KATEGORI', value: fmt(d.rows.length), note: nf1.format(d.rows.length / d.scanned * 100) + '% dari pinjaman dicek' },
    { label: 'TOTAL BALANCE', value: rpShort(d.stats ? d.stats.balance : 0), note: d.stats ? rp(d.stats.balance) : '—' },
    { label: 'TOTAL SIMPANAN', value: rpShort(d.stats ? d.stats.savings : 0), note: d.stats ? rp(d.stats.savings) : '—' }
  ], y, { cols: 2, cardH: 150, valueFont: 30 });
  y += 25;
  const categoryRows = CATS.map(c => {
    const rows = state.byCat[c.k] || [];
    const balance = rows.reduce((sum, r) => sum + (Number(r[F.BALANCE]) || 0), 0);
    return [c.title, fmt(rows.length), rpShort(balance), nf1.format(d.scanned ? rows.length / d.scanned * 100 : 0) + '%'];
  });
  y = imgTable(ctx, y, 'RINGKASAN PER KATEGORI',
    ['Kategori', 'Jumlah', 'Total Balance', 'Dari Dicek'], categoryRows,
    { subtitle: 'Jumlah pinjaman dan Total Balance per kategori PAR.',
      widths: [3.5, 1, 1.7, 1.1], rowH: 43, fontSize: 15, rightCols: [1, 2, 3] });
  y += 16;
  const cols = [
    { label: 'Branch Name', get: r => r[F.BRANCH] },
    { label: 'Loan No', get: r => r[F.LOAN] },
    { label: 'Client ID', get: r => r[F.CLIENT] },
    { label: 'Client Name', get: r => r[F.NAME] },
    { label: 'Product Name', get: r => r[F.PRODUCT] },
    { label: 'Total Balance', get: r => rp(r[F.BALANCE]) },
    { label: 'Kode/Staf', get: r => r[F.STAF] }
  ];
  const from = state.page * state.size, currentRows = state.view.slice(from, from + state.size).slice(0, 16);
  const loanRows = currentRows.map(r => cols.map(c => c.get(r)));
  y = imgTable(ctx, y, 'DAFTAR ' + String(catOf(state.cat).title).toLocaleUpperCase('id-ID'),
    cols.map(c => c.label), loanRows,
    { widths: [1.15, 1.05, 1.0, 1.55, 1.35, 1.15, 0.9], rowH: 43, fontSize: 12,
      subtitle: 'Menampilkan ' + fmt(from + 1) + '–' + fmt(from + currentRows.length) + ' dari ' + fmt(state.view.length) + ' data hasil filter dan urutan layar.',
      rightCols: [5], emptyText: 'Tidak ada data untuk kategori dan filter ini.' });
  imgFooter(ctx, 'Daftar pada gambar mengikuti halaman aktif dan dibatasi 16 baris. Unduh Excel untuk seluruh data kategori.');
  return report.canvas;
}

function imgDrawLineChart(ctx, x, y, width, height, daily, keyA, keyB, money, selectedDate, title) {
  imgRoundRect(ctx, x, y, width, height, 14, '#FFFFFF', '#DCE8E0');
  ctx.font = '700 17px Arial, sans-serif'; ctx.fillStyle = '#244638';
  ctx.fillText(imgFitText(ctx, title, width - 32), x + 16, y + 29);
  const plotX = x + 54, plotY = y + 58, plotW = width - 78, plotH = height - 105;
  const data = (daily || []).filter(d => d && d.date);
  if (!data.length) {
    ctx.font = '14px Arial, sans-serif'; ctx.fillStyle = '#7A8B82'; ctx.fillText('Belum ada data.', x + 18, y + 92); return;
  }
  const val = (r, key) => Number(r[key]) || 0;
  const max = Math.max(1, ...data.map(r => Math.max(val(r, keyA), val(r, keyB))));
  ctx.strokeStyle = '#E7EFEA'; ctx.lineWidth = 1;
  for (let i = 0; i <= 3; i++) {
    const gy = plotY + plotH - plotH * i / 3;
    ctx.beginPath(); ctx.moveTo(plotX, gy); ctx.lineTo(plotX + plotW, gy); ctx.stroke();
  }
  const px = i => data.length > 1 ? plotX + plotW * i / (data.length - 1) : plotX + plotW / 2;
  const py = v => plotY + plotH - (v / max) * plotH;
  const colors = ['#DC3545', '#198754'];
  [keyA, keyB].forEach((key, si) => {
    ctx.strokeStyle = colors[si]; ctx.fillStyle = colors[si]; ctx.lineWidth = 3;
    ctx.beginPath();
    data.forEach((r, i) => { const xx = px(i), yy = py(val(r, key)); if (i) ctx.lineTo(xx, yy); else ctx.moveTo(xx, yy); });
    ctx.stroke();
    data.forEach((r, i) => {
      const xx = px(i), yy = py(val(r, key));
      ctx.beginPath(); ctx.arc(xx, yy, 4, 0, Math.PI * 2); ctx.fill();
    });
  });
  if (selectedDate) {
    const idx = data.findIndex(r => r.date === selectedDate);
    if (idx >= 0) {
      ctx.strokeStyle = '#60776B'; ctx.setLineDash([5, 4]);
      ctx.beginPath(); ctx.moveTo(px(idx), plotY); ctx.lineTo(px(idx), plotY + plotH); ctx.stroke(); ctx.setLineDash([]);
    }
  }
  ctx.font = '12px Arial, sans-serif'; ctx.fillStyle = '#778980';
  const labelIndexes = [...new Set([0, Math.floor((data.length - 1) / 2), data.length - 1])];
  labelIndexes.forEach(i => { const day = String(data[i].date).slice(8, 10); ctx.textAlign = 'center'; ctx.fillText(day, px(i), plotY + plotH + 22); });
  ctx.textAlign = 'left';
  ctx.fillStyle = colors[0]; ctx.fillRect(x + 18, y + height - 22, 11, 7);
  ctx.font = '12px Arial, sans-serif'; ctx.fillStyle = '#53665D'; ctx.fillText('PAR Baru', x + 35, y + height - 15);
  ctx.fillStyle = colors[1]; ctx.fillRect(x + 125, y + height - 22, 11, 7); ctx.fillStyle = '#53665D'; ctx.fillText('PAR Selesai', x + 142, y + height - 15);
  ctx.textAlign = 'left';
}

/* ---------------------------------------------------------------------
   Ekspor khusus Progres Harian: F4 landscape, 300 DPI, layout 40:30:30.
   Hanya mengubah komposisi gambar; sumber angka tetap state.progress.
   --------------------------------------------------------------------- */
const PH_IMG_W = 3898;   // F4 landscape 330 mm @ 300 DPI
const PH_IMG_H = 2480;   // F4 landscape 210 mm @ 300 DPI
const PH_IMG_M = 96;
const PH_IMG_G = 36;

function phImgFont(ctx, size, weight) {
  ctx.font = (weight || '400') + ' ' + size + 'px Arial, sans-serif';
}

function phImgFit(ctx, value, maxWidth, size, weight, minSize) {
  const text = String(value == null ? '' : value);
  let fontSize = size || 32;
  const floor = minSize == null ? Math.max(20, Math.round(fontSize * 0.68)) : minSize;
  while (fontSize > floor) {
    phImgFont(ctx, fontSize, weight || '400');
    if (ctx.measureText(text).width <= maxWidth) break;
    fontSize -= 1;
  }
  phImgFont(ctx, fontSize, weight || '400');
  return imgFitText(ctx, text, maxWidth);
}

function phImgText(ctx, value, x, y, maxWidth, size, color, weight, align, minSize) {
  ctx.save();
  ctx.fillStyle = color || '#263D33';
  ctx.textAlign = align || 'left';
  const text = phImgFit(ctx, value, maxWidth, size || 30, weight || '400', minSize);
  ctx.fillText(text, x, y);
  ctx.restore();
}

function phImgPanel(ctx, x, y, w, h, fill) {
  imgRoundRect(ctx, x, y, w, h, 28, fill || '#FFFFFF', '#DCE8E0');
}

function phImgKpiCard(ctx, x, y, w, h, kind, amount, count) {
  const isNew = kind === 'new';
  const accent = isNew ? '#0A7A5A' : '#D09A20';
  const soft = isNew ? '#E6F4EC' : '#FFF4D8';
  const ink = isNew ? '#075B43' : '#815B0B';
  const label = isNew ? 'PAR BARU · PENINGKATAN' : 'PAR SELESAI · PENYELESAIAN';
  const mark = isNew ? '↑' : '✓';
  phImgPanel(ctx, x, y, w, h);
  ctx.save();
  ctx.fillStyle = accent;
  imgRoundRect(ctx, x + 28, y + 30, 12, h - 60, 6, accent);
  imgRoundRect(ctx, x + w - 106, y + 30, 72, 72, 20, soft);
  phImgFont(ctx, 52, '700'); ctx.fillStyle = ink; ctx.textAlign = 'center';
  ctx.fillText(mark, x + w - 70, y + 82);
  phImgText(ctx, label, x + 68, y + 83, w - 190, 37, ink, '700', 'left', 27);
  phImgText(ctx, amount, x + 54, y + 275, w - 105, 78, '#123D2F', '700', 'left', 43);
  phImgText(ctx, isNew ? 'Out. PAR bertambah' : 'Out. PAR terselesaikan', x + 56, y + 345, w - 110, 31, '#64766D', '400', 'left', 25);
  ctx.strokeStyle = '#E3ECE6'; ctx.lineWidth = 3;
  ctx.beginPath(); ctx.moveTo(x + 54, y + 405); ctx.lineTo(x + w - 48, y + 405); ctx.stroke();
  phImgText(ctx, fmt(count), x + 56, y + 490, w - 110, 54, '#183C2F', '700', 'left', 35);
  phImgText(ctx, 'Client PAR · jumlah Loan No', x + 56, y + 550, w - 110, 28, '#6B7E74', '400', 'left', 22);
  ctx.fillStyle = soft;
  imgRoundRect(ctx, x + 52, y + h - 94, w - 104, 54, 16, soft);
  phImgText(ctx, isNew ? 'PENINGKATAN' : 'PENYELESAIAN', x + 76, y + h - 58, w - 152, 25, ink, '700', 'left', 21);
  ctx.restore();
}

function phImgNetCard(ctx, x, y, w, h, d) {
  phImgPanel(ctx, x, y, w, h, '#F5F9F6');
  ctx.save();
  ctx.fillStyle = '#D9A928'; imgRoundRect(ctx, x + 34, y + 34, 12, 70, 6, '#D9A928');
  phImgText(ctx, 'SELISIH & PERUBAHAN OUTSTANDING', x + 76, y + 82, w - 130, 49, '#123D2F', '700', 'left', 35);
  phImgText(ctx, 'PAR Baru dibanding PAR Selesai serta perubahan saldo pada Loan No yang tetap muncul.', x + 76, y + 132, w - 130, 29, '#6B7E74', '400', 'left', 22);

  const reduction = d.balanceReduction || { amount: 0, loans: 0, note: 'Riwayat angsuran belum tersedia.' };
  const pad = 38, gap = 18, innerX = x + pad, innerW = w - pad * 2;
  const top = y + 184;
  const blockH = Math.max(190, Math.floor((h - 220 - gap * 2) / 3));
  const selectedMonth = d.filters && d.filters.monthKey;
  const selectedDate = d.filters && d.filters.progressDate;
  const reductionPeriod = selectedDate ? 'Tanggal ' + progresFilterDateLabel(selectedDate)
    : selectedMonth ? (d.monthLabel || selectedMonth) + ' · semua tanggal' : 'Seluruh riwayat';
  const metrics = [
    {
      label: 'SELISIH OUT. PAR', value: (Number(d.netBalance) > 0 ? '+' : '') + rpShort(d.netBalance),
      note: Number(d.netBalance) > 0 ? 'PAR Baru lebih besar daripada PAR Selesai' : Number(d.netBalance) < 0 ? 'PAR Selesai lebih besar daripada PAR Baru' : 'PAR Baru dan PAR Selesai seimbang',
      color: Number(d.netBalance) > 0 ? '#A66A08' : Number(d.netBalance) < 0 ? '#087A55' : '#52695D', sub: 'TOTAL BALANCE'
    },
    {
      label: 'SELISIH CLIENT PAR', value: (Number(d.netCount) > 0 ? '+' : '') + fmt(d.netCount),
      note: Number(d.netCount) > 0 ? 'Loan No baru lebih banyak daripada yang selesai' : Number(d.netCount) < 0 ? 'Loan No selesai lebih banyak daripada yang baru' : 'Jumlah Loan No seimbang',
      color: Number(d.netCount) > 0 ? '#A66A08' : Number(d.netCount) < 0 ? '#087A55' : '#52695D', sub: 'JUMLAH LOAN NO'
    },
    {
      label: 'PENURUNAN OUTSTANDING PAR BY ANGSURAN', value: rpShort(reduction.amount || 0),
      note: fmt(reduction.loans || 0) + ' Loan No dengan penurunan bersih positif · ' + reductionPeriod,
      color: Number(reduction.amount) > 0 ? '#087A55' : Number(reduction.amount) < 0 ? '#A66A08' : '#52695D',
      sub: 'DATA: REKENING_BERGERAK_ANGSURAN'
    }
  ];
  metrics.forEach((m, i) => {
    const by = top + i * (blockH + gap);
    phImgPanel(ctx, innerX, by, innerW, blockH, '#FFFFFF');
    phImgText(ctx, m.label, innerX + 28, by + 46, innerW - 56, 38, '#52695D', '700', 'left', 23);
    phImgText(ctx, m.value, innerX + 28, by + 126, innerW - 56, 64, m.color, '700', 'left', 43);
    phImgText(ctx, m.sub, innerX + 30, by + 169, innerW - 60, 23, '#84948B', '700', 'left', 18);
    ctx.strokeStyle = '#E6EEE8'; ctx.lineWidth = 2;
    ctx.beginPath(); ctx.moveTo(innerX + 28, by + 196); ctx.lineTo(innerX + innerW - 28, by + 196); ctx.stroke();
    phImgText(ctx, m.note, innerX + 30, by + Math.min(blockH - 22, 238), innerW - 60, 34, '#52695D', '400', 'left', 18);
  });
  ctx.restore();
}
function phImgRankPanel(ctx, x, y, w, h, kind, rows) {
  const isNew = kind === 'new';
  // PAR Baru = gradasi merah; PAR Selesai = gradasi hijau.
  // Opasitas tiap peringkat mengikuti permintaan: 50%, 30%, 10%.
  const rgb = isNew ? '220,53,69' : '25,135,84';
  const accent = isNew ? '#DC3545' : '#198754';
  const ink = isNew ? '#9F1D2D' : '#11643F';
  const title = isNew ? '3 Cabang dengan Kenaikan Tertinggi' : '3 Cabang dengan Penyelesaian Tertinggi';
  const desc = isNew ? 'Diurutkan berdasarkan akumulasi Out. PAR dari arsip PAR Baru.' : 'Diurutkan berdasarkan akumulasi Out. PAR dari arsip PAR Selesai.';
  phImgPanel(ctx, x, y, w, h);
  phImgText(ctx, title, x + 38, y + 74, w - 76, 47, '#123D2F', '700', 'left', 32);
  phImgText(ctx, desc, x + 40, y + 124, w - 80, 27, '#6B7E74', '400', 'left', 21);
  const pillW = 220;
  ctx.fillStyle = `rgba(${rgb},0.12)`;
  imgRoundRect(ctx, x + w - pillW - 34, y + 151, pillW, 52, 16, `rgba(${rgb},0.12)`);
  phImgText(ctx, isNew ? 'PAR BARU' : 'PAR SELESAI', x + w - pillW / 2 - 34, y + 187, pillW - 24, 25, ink, '700', 'center', 20);

  const list = Array.isArray(rows) ? rows.slice(0, 3) : [];
  const listTop = y + 230, rowGap = 20;
  const rowH = Math.floor((h - 270 - rowGap * 2) / 3);
  if (!list.length) {
    phImgText(ctx, 'Belum ada progres untuk periode ini.', x + 40, listTop + 100, w - 80, 34, '#7A8B82', '400', 'left', 24);
    return;
  }
  const opacityByRank = [0.50, 0.30, 0.10];
  list.forEach((r, i) => {
    const ry = listTop + i * (rowH + rowGap);
    const tint = `rgba(${rgb},${opacityByRank[i]})`;
    const border = `rgba(${rgb},${Math.min(0.88, opacityByRank[i] + 0.28)})`;
    const cardX = x + 30, cardW = w - 60;
    // Seluruh kartu cabang diberi warna sesuai intensitas peringkat.
    imgRoundRect(ctx, cardX, ry, cardW, rowH, 22, tint, border);
    // Aksen kiri memperjelas kelompok ranking tanpa mengganggu keterbacaan.
    imgRoundRect(ctx, cardX + 12, ry + 18, 12, rowH - 36, 6, accent);

    const badgeX = x + 99, badgeY = ry + rowH / 2;
    ctx.fillStyle = accent;
    ctx.beginPath(); ctx.arc(badgeX, badgeY, 42, 0, Math.PI * 2); ctx.fill();
    phImgText(ctx, String(i + 1).padStart(2, '0'), badgeX, badgeY + 17, 76, 39, '#FFFFFF', '700', 'center', 28);

    const branch = r.branch || '-';
    const count = isNew ? r.newCount : r.doneCount;
    const amount = isNew ? r.newBalance : r.doneBalance;
    phImgText(ctx, branch, x + 170, badgeY - 8, w - 500, 43, '#183C2F', '700', 'left', 27);
    phImgText(ctx, fmt(count) + ' Client PAR (Loan No)', x + 172, badgeY + 48, w - 500, 28, '#465D52', '400', 'left', 21);
    phImgText(ctx, rpShort(amount || 0), x + w - 62, badgeY - 5, 330, 39, ink, '700', 'right', 25);
    phImgText(ctx, 'Out. PAR', x + w - 62, badgeY + 42, 330, 24, '#465D52', '400', 'right', 19);
  });
}

function phImgChartPanel(ctx, x, y, w, h, daily, keyA, keyB, isMoney, selectedDate) {
  const title = isMoney ? 'Peningkatan vs Penyelesaian Out. PAR' : 'Peningkatan vs Penyelesaian Client';
  const desc = isMoney ? (selectedDate ? 'Nilai Total Balance seluruh tanggal; tanggal pilihan ditandai.' : 'Nilai Total Balance per tanggal progress.')
    : (selectedDate ? 'Jumlah Loan No seluruh tanggal; tanggal pilihan ditandai.' : 'Jumlah Loan No per tanggal progress.');
  phImgPanel(ctx, x, y, w, h);
  phImgText(ctx, title, x + 38, y + 75, w - 76, 48, '#123D2F', '700', 'left', 31);
  phImgText(ctx, desc, x + 40, y + 128, w - 80, 28, '#6B7E74', '400', 'left', 22);

  const data = (Array.isArray(daily) ? daily : []).filter(r => r && r.date);
  const plotX = x + 232, plotY = y + 235, plotW = w - 292, plotH = h - 450;
  const baseY = plotY + plotH;
  const colorNew = '#DC3545', colorDone = '#198754';
  if (!data.length) {
    phImgText(ctx, 'Belum ada data grafik untuk periode ini.', x + 42, plotY + 160, w - 84, 32, '#7A8B82', '400', 'left', 24);
  } else {
    const value = (r, key) => Number(r[key]) || 0;
    const max = Math.max(1, ...data.map(r => Math.max(value(r, keyA), value(r, keyB)))) * 1.12;
    const px = i => data.length > 1 ? plotX + plotW * i / (data.length - 1) : plotX + plotW / 2;
    const py = v => baseY - (v / max) * plotH;
    ctx.save();
    ctx.strokeStyle = '#E5EDE8'; ctx.lineWidth = 2;
    phImgFont(ctx, 27, '400'); ctx.fillStyle = '#718279'; ctx.textAlign = 'right';
    for (let i = 0; i <= 4; i++) {
      const ratio = i / 4, gy = baseY - plotH * ratio;
      ctx.beginPath(); ctx.moveTo(plotX, gy); ctx.lineTo(plotX + plotW, gy); ctx.stroke();
      const tick = max * ratio;
      const label = isMoney ? rpShort(tick) : fmt(Math.round(tick));
      ctx.fillText(phImgFit(ctx, label, plotX - x - 30, 27, '400', 19), plotX - 28, gy + 9);
    }
    ctx.strokeStyle = '#BBCBC1'; ctx.lineWidth = 3;
    ctx.beginPath(); ctx.moveTo(plotX, plotY); ctx.lineTo(plotX, baseY); ctx.lineTo(plotX + plotW, baseY); ctx.stroke();

    [[keyA, colorNew, 'circle'], [keyB, colorDone, 'diamond']].forEach(series => {
      const [key, color, marker] = series;
      ctx.strokeStyle = color; ctx.fillStyle = color; ctx.lineWidth = 9; ctx.lineJoin = 'round'; ctx.lineCap = 'round';
      ctx.beginPath();
      data.forEach((r, i) => { const xx = px(i), yy = py(value(r, key)); if (i) ctx.lineTo(xx, yy); else ctx.moveTo(xx, yy); });
      ctx.stroke();
      data.forEach((r, i) => {
        const xx = px(i), yy = py(value(r, key));
        ctx.beginPath();
        if (marker === 'circle') ctx.arc(xx, yy, 15, 0, Math.PI * 2);
        else { ctx.moveTo(xx, yy - 18); ctx.lineTo(xx + 18, yy); ctx.lineTo(xx, yy + 18); ctx.lineTo(xx - 18, yy); ctx.closePath(); }
        ctx.fill(); ctx.strokeStyle = '#FFFFFF'; ctx.lineWidth = 4; ctx.stroke(); ctx.strokeStyle = color; ctx.lineWidth = 9;
      });
    });

    if (selectedDate) {
      const idx = data.findIndex(r => r.date === selectedDate);
      if (idx >= 0) {
        ctx.strokeStyle = '#52695D'; ctx.lineWidth = 3; ctx.setLineDash([12, 10]);
        ctx.beginPath(); ctx.moveTo(px(idx), plotY); ctx.lineTo(px(idx), baseY); ctx.stroke(); ctx.setLineDash([]);
      }
    }
    const labelCount = Math.min(5, data.length);
    const labelIndexes = [...new Set(Array.from({ length: labelCount }, (_, i) => labelCount === 1 ? 0 : Math.round(i * (data.length - 1) / (labelCount - 1))))];
    ctx.font = '700 29px Arial, sans-serif'; ctx.fillStyle = '#718279'; ctx.textAlign = 'center';
    labelIndexes.forEach(i => ctx.fillText(String(data[i].date).slice(8, 10), px(i), baseY + 52));
    ctx.restore();
  }
  const legY = y + h - 72;
  ctx.save();
  ctx.strokeStyle = colorNew; ctx.fillStyle = colorNew; ctx.lineWidth = 8;
  ctx.beginPath(); ctx.moveTo(x + 54, legY - 8); ctx.lineTo(x + 112, legY - 8); ctx.stroke();
  ctx.beginPath(); ctx.arc(x + 83, legY - 8, 12, 0, Math.PI * 2); ctx.fill();
  phImgText(ctx, 'PAR Baru', x + 132, legY + 2, 210, 29, '#53665D', '700', 'left', 23);
  ctx.strokeStyle = colorDone; ctx.fillStyle = colorDone;
  ctx.beginPath(); ctx.moveTo(x + 390, legY - 8); ctx.lineTo(x + 448, legY - 8); ctx.stroke();
  ctx.beginPath(); ctx.moveTo(x + 419, legY - 23); ctx.lineTo(x + 434, legY - 8); ctx.lineTo(x + 419, legY + 7); ctx.lineTo(x + 404, legY - 8); ctx.closePath(); ctx.fill();
  phImgText(ctx, 'PAR Selesai', x + 468, legY + 2, 260, 29, '#53665D', '700', 'left', 23);
  ctx.restore();
}

function imgBuildProgresHarian() {
  const d = state.progress;
  if (!d || !d.totals) throw new Error('Data Progres Harian belum siap.');
  const filters = d.filters || {};
  const filterText = [
    'Bulan: ' + ((filters.monthKey && d.monthLabel) || 'Semua Bulan'),
    'Cabang: ' + (filters.branch || 'Semua Cabang'),
    'Tanggal: ' + (filters.progressDate ? progresFilterDateLabel(filters.progressDate) : 'Semua Tanggal'),
    'Branch Coordinator: ' + (filters.coordinatorName || 'Semua')
  ].join('   ·   ');
  const meta = ((filters.monthKey && d.monthLabel) || 'Semua Periode') + '   ·   s.d. ' + fmtStamp(d.asOf) + '   ·   ' + (d.scope || 'Sesuai Wilayah Kerja');
  const canvas = document.createElement('canvas');
  canvas.width = PH_IMG_W; canvas.height = PH_IMG_H;
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('Browser tidak dapat menyiapkan kanvas gambar Progres Harian.');
  ctx.fillStyle = '#F3F7F4'; ctx.fillRect(0, 0, PH_IMG_W, PH_IMG_H);

  // Header landscape dan filter aktif.
  ctx.fillStyle = '#064E3B'; ctx.fillRect(0, 0, PH_IMG_W, 292);
  ctx.fillStyle = '#D9A928'; imgRoundRect(ctx, PH_IMG_M, 54, 14, 54, 7, '#D9A928');
  phImgText(ctx, 'DASHBOARD OPERASIONAL SATU PINTU', PH_IMG_M + 42, 91, PH_IMG_W - PH_IMG_M * 2 - 80, 31, '#BFE6D4', '700', 'left', 25);
  phImgText(ctx, 'PROGRES HARIAN', PH_IMG_M, 188, PH_IMG_W - PH_IMG_M * 2, 82, '#FFFFFF', '700', 'left', 58);
  phImgText(ctx, meta, PH_IMG_M, 247, PH_IMG_W - PH_IMG_M * 2, 31, '#E0F2E9', '400', 'left', 24);
  imgRoundRect(ctx, PH_IMG_M, 318, PH_IMG_W - PH_IMG_M * 2, 72, 18, '#E7F0EA', '#D5E3DA');
  phImgText(ctx, filterText + (d.filterDescription ? '   ·   ' + d.filterDescription : ''), PH_IMG_M + 28, 365, PH_IMG_W - PH_IMG_M * 2 - 56, 31, '#52695D', '600', 'left', 22);

  // Rasio kolom tepat 40:30:30 setelah margin dan gap.
  const contentX = PH_IMG_M, contentY = 420;
  const contentW = PH_IMG_W - PH_IMG_M * 2 - PH_IMG_G * 2;
  const ratioBase = contentW / 10;
  const leftW = Math.round(ratioBase * 4);
  const midW = Math.round(ratioBase * 3);
  const rightW = contentW - leftW - midW;
  const leftX = contentX;
  const midX = leftX + leftW + PH_IMG_G;
  const rightX = midX + midW + PH_IMG_G;
  const contentH = 1900;

  // KIRI 40%: dua kartu utama di atas, ringkasan selisih selebar kolom di bawah.
  const leftInnerGap = 24;
  const kpiW = Math.floor((leftW - leftInnerGap) / 2);
  const kpiH = 700;
  phImgKpiCard(ctx, leftX, contentY, kpiW, kpiH, 'new', rpShort(d.totals.newBalance), d.totals.newCount);
  phImgKpiCard(ctx, leftX + kpiW + leftInnerGap, contentY, leftW - kpiW - leftInnerGap, kpiH, 'done', rpShort(d.totals.doneBalance), d.totals.doneCount);
  phImgNetCard(ctx, leftX, contentY + kpiH + PH_IMG_G, leftW, contentH - kpiH - PH_IMG_G, d);

  // TENGAH 30%: dua kartu ranking cabang, sesuai urutan dan isi UI.
  const stackGap = 36, stackH = Math.floor((contentH - stackGap) / 2);
  phImgRankPanel(ctx, midX, contentY, midW, stackH, 'new', d.topNew || []);
  phImgRankPanel(ctx, midX, contentY + stackH + stackGap, midW, contentH - stackH - stackGap, 'done', d.topDone || []);

  // KANAN 30%: dua grafik line-with-marker dengan skema merah/hijau.
  const chartDaily = Array.isArray(d.dailyChart) ? d.dailyChart : (d.daily || []);
  const selectedDate = filters.progressDate || '';
  phImgChartPanel(ctx, rightX, contentY, rightW, stackH, chartDaily, 'newBalance', 'doneBalance', true, selectedDate);
  phImgChartPanel(ctx, rightX, contentY + stackH + stackGap, rightW, contentH - stackH - stackGap, chartDaily, 'newCount', 'doneCount', false, selectedDate);

  // Footer dibuat di luar tiga kolom agar tidak mengurangi area isi.
  ctx.fillStyle = '#D7E4DC'; ctx.fillRect(PH_IMG_M, PH_IMG_H - 124, PH_IMG_W - PH_IMG_M * 2, 3);
  phImgText(ctx, 'Ringkasan mengikuti filter dan cakupan Wilayah Kerja yang aktif. Satuan Out. PAR adalah Total Balance; Client PAR dihitung berdasarkan jumlah Loan No.', PH_IMG_M, PH_IMG_H - 78, PH_IMG_W - PH_IMG_M * 2, 24, '#61746A', '400', 'left', 20);
  const stamp = new Date().toLocaleString('id-ID', { dateStyle: 'medium', timeStyle: 'short' });
  phImgText(ctx, 'Dibuat ' + stamp + '  ·  F4 Landscape 330 × 210 mm  ·  300 DPI  ·  Layout 40:30:30', PH_IMG_M, PH_IMG_H - 34, PH_IMG_W - PH_IMG_M * 2, 23, '#819188', '400', 'left', 20);
  return canvas;
}

function imgBuildAnalisaCabang() {
  const d = state.analisaCabang, f = state.analisaFilters;
  if (!d || !f.period) throw new Error('Pilih Periode Analisa setelah data selesai dimuat.');
  const currentRows = analisaRowsFor(f.period).slice().sort((a, b) => collator.compare(a.branch, b.branch));
  if (!currentRows.length) throw new Error('Tidak ada data pada periode dan filter cabang yang dipilih.');
  const comparePeriod = analisaIsEarlierPeriod(f.compare, f.period) ? f.compare : '';
  const current = analisaAggregate(currentRows), oldRows = comparePeriod ? analisaRowsFor(comparePeriod) : [];
  const old = comparePeriod ? analisaAggregate(oldRows) : null;
  const branchLabel = f.branch === '*' ? 'Semua Cabang' : (($('an-branch').selectedOptions[0] || {}).textContent || f.branch);
  const meta = 'Periode: ' + analisaLabelDate(f.period) + ' · Pembanding: ' + (comparePeriod ? analisaLabelDate(comparePeriod) : 'Tidak dipilih') + ' · ' + branchLabel + ' · ' + (d.scope || 'Sesuai Wilayah Kerja');
  const report = imgPrepare('Analisa Cabang', meta), ctx = report.ctx;
  let y = report.y;
  const cards = ANALISA_METRICS.map(metric => {
    const value = Number(current[metric.key]) || 0;
    let note = metric.hint;
    if (old) {
      const before = Number(old[metric.key]) || 0, delta = value - before;
      const pct = before === 0 ? (value === 0 ? '0,0%' : 'N/A · dasar 0') : ((delta > 0 ? '+' : '') + nf1.format(delta / Math.abs(before) * 100) + '%');
      note = 'Pembanding: ' + (metric.kind === 'money' ? rpShort(before) : fmt(before)) + ' · Δ ' + analisaFormatDelta(metric, delta) + ' (' + pct + ')';
    }
    return { label: metric.label.toLocaleUpperCase('id-ID'), value: analisaFormatMetric(metric, value), note };
  });
  y = imgCards(ctx, cards, y, { cols: 2, cardH: 144, gap: 14, valueFont: 25 });
  y += 22;
  const metrics = ANALISA_METRICS.map(m => [m.label, rpOrCount(m, current[m.key] || 0), old ? rpOrCount(m, old[m.key] || 0) : '—', old ? analisaFormatDelta(m, (current[m.key] || 0) - (old[m.key] || 0)) : '—']);
  y = imgTable(ctx, y, 'PERBANDINGAN INDIKATOR', ['Indikator', 'Terpilih', 'Pembanding', 'Perubahan'], metrics,
    { widths: [1.8, 1.4, 1.4, 1.3], rowH: 38, fontSize: 13, rightCols: [1, 2, 3] });
  y += 14;
  const total = analisaAggregate(currentRows);
  const rows = currentRows.slice(0, 23).map(r => [r.branch, fmt(r.member), rp(r.outstanding), rp(r.disbursement), rp(r.saving), rp(r.repayment), rp(r.margin), fmt(r.memberPAR), rp(r.outstandingPAR)]);
  rows.push({ total: true, cells: ['TOTAL · ' + fmt(currentRows.length) + ' CABANG', fmt(total.member), rp(total.outstanding), rp(total.disbursement), rp(total.saving), rp(total.repayment), rp(total.margin), fmt(total.memberPAR), rp(total.outstandingPAR)] });
  y = imgTable(ctx, y, 'RINCIAN INDIKATOR PER CABANG', ['Cabang', 'Member', 'Outstanding', 'Disbursement', 'Saving', 'Repayment', 'Margin', 'Member PAR', 'Outstanding PAR'], rows,
    { widths: [1.55, 0.75, 1.15, 1.15, 1.0, 1.0, 0.9, 0.85, 1.25], rowH: 38, fontSize: 11,
      subtitle: 'Menampilkan maksimal 23 cabang ditambah baris total. Angka rupiah ditampilkan lengkap.', rightCols: [1, 2, 3, 4, 5, 6, 7, 8], maxRows: 24 });
  if (currentRows.length > 23) {
    ctx.font = '13px Arial, sans-serif'; ctx.fillStyle = '#6B7E74';
    ctx.fillText('Sebanyak ' + fmt(currentRows.length) + ' cabang tersedia; gambar menampilkan 23 cabang pertama. Excel memuat seluruh cabang.', IMG_F4_M, Math.min(y + 8, IMG_F4_H - 110));
  }
  imgFooter(ctx, 'Perubahan (%) = (Periode Terpilih − Pembanding) ÷ nilai absolut Pembanding × 100%. Nilai pembanding 0 ditampilkan N/A bila relevan.');
  return report.canvas;
}

function rpOrCount(metric, value) { return metric.kind === 'money' ? rp(value) : fmt(value); }

/* ---------------------------------------------------------------------
   Buletin Ringkasan: "Apa kabar Regional J?"
   Sumber tetap memakai endpoint yang sudah menerapkan Wilayah_Kerja.
   Progres diminta tanpa filter tampilan agar mewakili seluruh cakupan user.
   --------------------------------------------------------------------- */
const REGIONAL_NEWS_W = 1440;
const REGIONAL_NEWS_H = 2620;
const REGIONAL_NEWS_M = 72;

async function regionalNewsWaitForModule(isLoading, label) {
  const deadline = Date.now() + API_TIMEOUT_MS + 15000;
  while (isLoading() && Date.now() < deadline) await sleep(160);
  if (isLoading()) throw new Error(label + ' belum selesai dimuat. Coba buka kembali Ringkasan setelah sinkronisasi selesai.');
}

async function regionalNewsEnsureModule(label, isReady, isLoading, loadedAt, loader) {
  if (isLoading()) await regionalNewsWaitForModule(isLoading, label);
  if (isReady() && Date.now() - loadedAt() <= AUTO_SYNC_MS) return true;
  if (!isLoading()) await loader();
  if (isLoading()) await regionalNewsWaitForModule(isLoading, label);
  if (!isReady()) throw new Error(label + ' belum berhasil dimuat. Periksa halaman modul tersebut, lalu buka Ringkasan kembali.');
  return true;
}

async function regionalNewsLoadProgress(seq) {
  const isCurrent = () => seq == null || seq === state.newsLoadSeq;
  if (!isCurrent() || !state.user) throw new Error('Pemuatan buletin dibatalkan karena sesi atau data berubah.');
  const now = Date.now();
  if (state.newsProgress && now - state.newsProgressLoadedAt <= AUTO_SYNC_MS) return true;
  const filters = state.progressFilters || {};
  if (state.progress && !filters.branch && !filters.monthKey && !filters.progressDate && !filters.coordinatorNik
      && now - state.progressLoadedAt <= AUTO_SYNC_MS) {
    if (!isCurrent()) throw new Error('Pemuatan buletin dibatalkan karena data berubah.');
    state.newsProgress = state.progress;
    state.newsProgressLoadedAt = state.progressLoadedAt;
    return true;
  }
  const r = await api('webProgresHarian', {
    token: (session.get() || {}).token,
    branch: '', monthKey: '', progressDate: '', coordinatorNik: ''
  }, { retry: true });
  if (!isCurrent() || !state.user) throw new Error('Pemuatan buletin dibatalkan karena sesi atau data berubah.');
  if (r.code === 'SESSION_REQUIRED') {
    logout(r.error);
    throw new Error('Sesi berakhir. Silakan login kembali.');
  }
  if (!r.ok) throw new Error(r.error || 'Data Progres Harian tidak dapat dimuat untuk buletin.');
  state.newsProgress = r;
  state.newsProgressLoadedAt = Date.now();
  return true;
}

function regionalNewsSourcesReady() {
  return !!(state.rekening && state.data && state.analisaCabang && state.newsProgress);
}
function regionalNewsAllReady() {
  return !!state.newsReady && regionalNewsSourcesReady();
}

function invalidateRegionalNewsData() {
  state.newsLoadSeq++;
  state.newsProgress = null; state.newsProgressLoadedAt = 0;
  state.newsLoading = false; state.newsPromise = null; state.newsReady = false;
  const button = $('sum-news-img');
  if (button) button.disabled = true;
}

async function prepareRegionalNewsData() {
  if (!state.user || !canDownloadImage()) return;
  if (state.newsPromise) return state.newsPromise;
  const seq = ++state.newsLoadSeq;
  const userNik = state.user.nik;
  state.newsLoading = true; state.newsReady = false;
  const button = $('sum-news-img');
  if (button) button.disabled = true;
  setNote('sum-news-msg', 'Mengumpulkan data Rekening Bergerak, Analisa Cabang, Progres Harian, dan Tool PAR…');

  const jobs = [
    ['Rekening Bergerak', () => regionalNewsEnsureModule('Rekening Bergerak', () => !!state.rekening,
      () => state.rekeningLoading, () => state.rekeningLoadedAt, () => loadRekeningBergerak())],
    ['Tool PAR', () => regionalNewsEnsureModule('Tool PAR', () => !!state.data,
      () => state.loading, () => state.lastCheck, () => loadDash())],
    ['Analisa Cabang', () => regionalNewsEnsureModule('Analisa Cabang', () => !!state.analisaCabang,
      () => state.analisaLoading, () => state.analisaLoadedAt, () => loadAnalisaCabang())],
    ['Progres Harian', () => regionalNewsLoadProgress(seq)]
  ];
  state.newsPromise = (async () => {
    const results = await Promise.all(jobs.map(async ([label, run]) => {
      try { await run(); return { label, ok: true }; }
      catch (error) { return { label, ok: false, error: error && error.message ? error.message : String(error) }; }
    }));
    if (seq !== state.newsLoadSeq || !state.user || state.user.nik !== userNik) return;
    const failed = results.filter(x => !x.ok);
    const complete = !failed.length && regionalNewsSourcesReady();
    state.newsReady = complete;
    if (button) button.disabled = !complete;
    if (complete) {
      const stamp = new Date().toLocaleString('id-ID', { dateStyle: 'medium', timeStyle: 'short' });
      setNote('sum-news-msg', 'Buletin siap diunduh. Empat sumber data berhasil disiapkan · diperiksa ' + stamp + '.', 'good');
    } else {
      const details = failed.map(x => x.label + ': ' + x.error);
      if (!failed.length) details.push('Masih ada sumber data yang belum tersedia.');
      setNote('sum-news-msg', 'Buletin belum lengkap. ' + details.join(' '), 'bad');
    }
  })().finally(() => {
    if (seq === state.newsLoadSeq) {
      state.newsLoading = false;
      state.newsPromise = null;
      if (button) button.disabled = !state.newsReady;
    }
  });
  return state.newsPromise;
}

function regionalNewsText(ctx, text, x, y, width, font, color, maxLines, lineHeight) {
  ctx.font = font || '400 30px Arial, sans-serif';
  return imgWrapText(ctx, text, x, y, width, lineHeight || 38, maxLines || 2, color || '#52665C');
}

// Tipografi buletin menggunakan skala tetap. Hanya teks nilai yang sangat panjang
// boleh mengecil sedikit agar tidak keluar dari kartu.
function regionalNewsFitFont(ctx, text, maxWidth, baseSize, minSize, weight) {
  let size = baseSize || 30;
  const min = minSize || size;
  const fontWeight = weight || '400';
  while (size > min && ctx.measureText(String(text == null ? '' : text)).width > maxWidth) size--;
  ctx.font = fontWeight + ' ' + size + 'px Arial, sans-serif';
  return size;
}

function regionalNewsSingleLine(ctx, text, x, y, width, baseSize, minSize, weight, color) {
  const value = String(text == null ? '' : text);
  regionalNewsFitFont(ctx, value, width, baseSize, minSize, weight);
  ctx.fillStyle = color || '#52665C';
  ctx.fillText(imgFitText(ctx, value, width), x, y);
}

function regionalNewsPanel(ctx, y, height, number, title, subtitle, accent) {
  const x = REGIONAL_NEWS_M, width = REGIONAL_NEWS_W - REGIONAL_NEWS_M * 2;
  imgRoundRect(ctx, x, y, width, height, 28, '#FFFFFF', '#DDE8E1');
  imgRoundRect(ctx, x + 28, y + 27, 68, 68, 20, accent || '#0A7A5A');
  ctx.font = '700 28px Arial, sans-serif'; ctx.fillStyle = '#FFFFFF'; ctx.textAlign = 'center';
  ctx.fillText(String(number).padStart(2, '0'), x + 62, y + 72); ctx.textAlign = 'left';
  regionalNewsSingleLine(ctx, title, x + 120, y + 60, width - 156, 44, 38, '700', '#173F31');
  regionalNewsSingleLine(ctx, subtitle || '', x + 120, y + 96, width - 156, 26, 23, '400', '#728178');
  ctx.fillStyle = accent || '#0A7A5A'; ctx.fillRect(x + 30, y + 112, 72, 5);
  return { x: x + 30, y: y + 140, width: width - 60, height: height - 158 };
}

function regionalNewsMetric(ctx, x, y, width, label, value, tone, detail) {
  const height = 150;
  imgRoundRect(ctx, x, y, width, height, 18, '#F4F8F5', '#E0EAE3');
  regionalNewsText(ctx, String(label || '').toLocaleUpperCase('id-ID'), x + 18, y + 29,
    width - 36, '700 24px Arial, sans-serif', '#667A70', 2, 25);
  regionalNewsSingleLine(ctx, value == null ? '—' : value, x + 18, y + 101,
    width - 36, 48, 34, '700', tone || '#123D2F');
  if (detail) regionalNewsSingleLine(ctx, detail, x + 18, y + 134,
    width - 36, 21, 18, '400', '#728178');
}

function regionalNewsAggregateAnalisa() {
  const data = state.analisaCabang || {}, records = Array.isArray(data.records) ? data.records : [];
  const periods = Array.isArray(data.periods) ? data.periods.slice().sort((a, b) => b.localeCompare(a)) : [];
  const period = periods[0] || records.map(r => r.period).filter(Boolean).sort().reverse()[0] || '';
  const currentRows = period ? records.filter(r => r.period === period) : [];
  const comparePeriod = period ? periods.find(p => p < period) || '' : '';
  const compareRows = comparePeriod ? records.filter(r => r.period === comparePeriod) : [];
  return {
    period, comparePeriod, currentRows, current: analisaAggregate(currentRows),
    compare: compareRows.length ? analisaAggregate(compareRows) : null,
    topParBranch: currentRows.slice().sort((a, b) => (Number(b.outstandingPAR) || 0) - (Number(a.outstandingPAR) || 0))[0] || null
  };
}

function imgBuildRegionalNews() {
  if (!canDownloadImage()) throw new Error('Fitur Download Gambar belum diaktifkan untuk tingkat akun ini.');
  if (!regionalNewsAllReady()) throw new Error('Empat sumber data belum lengkap. Tunggu proses pemuatan selesai, lalu coba kembali.');

  const canvas = document.createElement('canvas');
  canvas.width = REGIONAL_NEWS_W; canvas.height = REGIONAL_NEWS_H;
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('Browser tidak dapat menyiapkan gambar buletin.');
  ctx.fillStyle = '#EEF4F0'; ctx.fillRect(0, 0, REGIONAL_NEWS_W, REGIONAL_NEWS_H);

  // Masthead: satu judul dominan, tagline, dan tanggal dengan skala yang konsisten.
  ctx.fillStyle = '#064E3B'; ctx.fillRect(0, 0, REGIONAL_NEWS_W, 300);
  ctx.fillStyle = '#E2B83E'; ctx.fillRect(REGIONAL_NEWS_M, 50, 14, 48);
  regionalNewsSingleLine(ctx, 'BULETIN OPERASIONAL  ·  REGIONAL J', REGIONAL_NEWS_M + 34, 84,
    REGIONAL_NEWS_W - REGIONAL_NEWS_M * 2 - 34, 22, 22, '700', '#BFE6D4');
  regionalNewsSingleLine(ctx, 'Apa kabar Regional J?', REGIONAL_NEWS_M, 184,
    REGIONAL_NEWS_W - REGIONAL_NEWS_M * 2, 76, 68, '700', '#FFFFFF');
  regionalNewsSingleLine(ctx, 'Empat sorotan penting, satu cerita untuk kita semua.', REGIONAL_NEWS_M, 237,
    REGIONAL_NEWS_W - REGIONAL_NEWS_M * 2, 34, 30, '400', '#E2F1E8');
  const issueDate = new Date().toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
  regionalNewsSingleLine(ctx, issueDate, REGIONAL_NEWS_M, 278,
    REGIONAL_NEWS_W - REGIONAL_NEWS_M * 2, 25, 22, '600', '#D6E9DF');

  // Lead paragraph.
  imgRoundRect(ctx, REGIONAL_NEWS_M, 320, REGIONAL_NEWS_W - REGIONAL_NEWS_M * 2, 150, 24, '#FFFFFF', '#DDE8E1');
  regionalNewsSingleLine(ctx, 'KABAR UTAMA', REGIONAL_NEWS_M + 28, 363,
    REGIONAL_NEWS_W - REGIONAL_NEWS_M * 2 - 56, 30, 28, '700', '#0A7654');
  regionalNewsText(ctx,
    'Buletin ini merangkum empat hal utama yang perlu diketahui. Setiap modul tetap menerapkan pembatasan cabang sesuai hak akses akun.',
    REGIONAL_NEWS_M + 28, 408, REGIONAL_NEWS_W - REGIONAL_NEWS_M * 2 - 56,
    '400 30px Arial, sans-serif', '#405B4E', 2, 38);

  // 01. Rekening Bergerak.
  const rb = state.rekening || {}, rbRows = Array.isArray(rb.branchRows) ? rb.branchRows : [];
  const rbTotal = rbRows.reduce((sum, row) => sum + (Number(row.total) || 0), 0);
  const rbActive = rbRows.reduce((sum, row) => sum + (Number(row.active) || 0), 0);
  const rbPct = rbTotal ? rbActive / rbTotal * 100 : 0;
  const rbTop = rbRows.slice().sort((a, b) => (Number(b.active) || 0) - (Number(a.active) || 0))[0];
  const p1 = regionalNewsPanel(ctx, 490, 390, 1, 'Rekening Bergerak', rb.settingDate ? 'Tanggal N: ' + rbDateLabel(rb.settingDate) : 'Tanggal N belum ditetapkan', '#0A7A5A');
  const gap3 = 18, cardW3 = (p1.width - gap3 * 2) / 3;
  if (rb.settingRequired || !rb.settingDate) {
    regionalNewsMetric(ctx, p1.x, p1.y, p1.width, 'Status', 'Belum dihitung', '#9B6A12', 'Tanggal N belum diatur');
    regionalNewsText(ctx, 'Rekening Bergerak belum dapat dihitung karena tanggal batas (Tanggal N) belum ditetapkan oleh Administrator atau User 0.',
      p1.x, p1.y + 174, p1.width, '400 30px Arial, sans-serif', '#53665D', 2, 38);
  } else {
    regionalNewsMetric(ctx, p1.x, p1.y, cardW3, 'Rekening aktif', fmt(rbActive), '#087A5A', 'Bertransaksi sejak Tanggal N');
    regionalNewsMetric(ctx, p1.x + cardW3 + gap3, p1.y, cardW3, 'Total rekening PAR', fmt(rb.totalRows == null ? rbTotal : rb.totalRows), '#173F31', 'Dalam cakupan akses');
    regionalNewsMetric(ctx, p1.x + (cardW3 + gap3) * 2, p1.y, cardW3, 'Porsi bergerak', nf1.format(rbPct) + '%', '#087A5A', 'Dibanding total rekening');
    const line = rbTop && rbTop.active
      ? 'Aktivitas terbanyak tercatat di ' + rbTop.branch + ' dengan ' + fmt(rbTop.active) + ' rekening bergerak.'
      : 'Belum ada rekening yang tercatat bergerak pada cakupan dan tanggal batas ini.';
    regionalNewsText(ctx, line, p1.x, p1.y + 174, p1.width, '600 30px Arial, sans-serif', '#405B4E', 2, 38);
  }

  // 02. Analisa Cabang.
  const an = regionalNewsAggregateAnalisa();
  const p2 = regionalNewsPanel(ctx, 900, 530, 2, 'Analisa Cabang', 'Posisi indikator pada periode analisa terbaru', '#D0A12B');
  const gap2 = 18, cardW2 = (p2.width - gap2) / 2;
  if (an.period && an.currentRows.length) {
    regionalNewsMetric(ctx, p2.x, p2.y, cardW2, 'Member', fmt(an.current.member), '#173F31', 'Jumlah anggota');
    regionalNewsMetric(ctx, p2.x + cardW2 + gap2, p2.y, cardW2, 'Outstanding', rpShort(an.current.outstanding), '#173F31', 'Saldo pinjaman berjalan');
    regionalNewsMetric(ctx, p2.x, p2.y + 160, cardW2, 'Disbursement', rpShort(an.current.disbursement), '#087A5A', 'Penyaluran dana');
    regionalNewsMetric(ctx, p2.x + cardW2 + gap2, p2.y + 160, cardW2, 'Outstanding PAR', rpShort(an.current.outstandingPAR), '#A64B3C', 'Saldo pinjaman PAR');
    let comparison = 'Periode data: ' + analisaLabelDate(an.period) + '. ';
    if (an.compare) {
      const delta = (Number(an.current.outstandingPAR) || 0) - (Number(an.compare.outstandingPAR) || 0);
      comparison += delta < 0
        ? 'Outstanding PAR turun ' + rpShort(Math.abs(delta)) + ' dibanding ' + analisaLabelDate(an.comparePeriod) + '.'
        : delta > 0
          ? 'Outstanding PAR meningkat ' + rpShort(delta) + ' dibanding ' + analisaLabelDate(an.comparePeriod) + '.'
          : 'Outstanding PAR tetap dibanding ' + analisaLabelDate(an.comparePeriod) + '.';
    } else comparison += 'Belum tersedia periode pembanding.';
    if (an.topParBranch) comparison += ' Cabang dengan Outstanding PAR tertinggi: ' + an.topParBranch.branch + '.';
    regionalNewsText(ctx, comparison, p2.x, p2.y + 326, p2.width, '600 29px Arial, sans-serif', '#405B4E', 2, 35);
  } else {
    regionalNewsText(ctx, 'Data Analisa Cabang belum tersedia untuk periode mana pun. Unggah file Analisa melalui menu Unggah Excel agar ringkasan indikator cabang dapat ditampilkan.',
      p2.x, p2.y + 42, p2.width, '400 30px Arial, sans-serif', '#53665D', 3, 38);
  }

  // 03. Progres Harian: snapshot tanpa filter UI, tetap dibatasi Wilayah Kerja server.
  const ph = state.newsProgress || {}, pt = ph.totals || {};
  const newBalance = Number(pt.newBalance) || 0, doneBalance = Number(pt.doneBalance) || 0;
  const newCount = Number(pt.newCount) || 0, doneCount = Number(pt.doneCount) || 0;
  const netBalance = newBalance - doneBalance, netCount = newCount - doneCount;
  const p3 = regionalNewsPanel(ctx, 1450, 430, 3, 'Progres Harian', 'PAR Baru, PAR Selesai, dan perubahan angsuran · seluruh periode', '#0A7A5A');
  const phReduction = ph.balanceReduction || { amount: 0, loans: 0 };
  const phCardGap = 16, phCardW = (p3.width - phCardGap * 2) / 3;
  regionalNewsMetric(ctx, p3.x, p3.y, phCardW, 'PAR Baru · peningkatan', rpShort(newBalance), '#B43D3D', fmt(newCount) + ' Client PAR / Loan No');
  regionalNewsMetric(ctx, p3.x + phCardW + phCardGap, p3.y, phCardW, 'PAR Selesai · penyelesaian', rpShort(doneBalance), '#087A5A', fmt(doneCount) + ' Client PAR / Loan No');
  regionalNewsMetric(ctx, p3.x + (phCardW + phCardGap) * 2, p3.y, phCardW, 'Penurunan Outstanding PAR by Angsuran', rpShort(phReduction.amount || 0),
    Number(phReduction.amount) > 0 ? '#087A5A' : Number(phReduction.amount) < 0 ? '#A66A08' : '#52695D',
    fmt(phReduction.loans || 0) + ' Loan No dengan penurunan bersih positif');
  const balanceStory = netBalance > 0
    ? 'Out. PAR Baru lebih tinggi ' + rpShort(netBalance) + '.'
    : netBalance < 0 ? 'Out. PAR Selesai lebih tinggi ' + rpShort(Math.abs(netBalance)) + '.'
      : 'Nilai Out. PAR Baru dan Selesai seimbang.';
  const countStory = netCount > 0
    ? 'PAR Baru lebih banyak ' + fmt(netCount) + ' Loan No.'
    : netCount < 0 ? 'PAR Selesai lebih banyak ' + fmt(Math.abs(netCount)) + ' Loan No.'
      : 'Jumlah Loan No seimbang.';
  regionalNewsText(ctx, balanceStory + ' ' + countStory, p3.x, p3.y + 176, p3.width,
    '600 29px Arial, sans-serif', '#405B4E', 2, 35);
  const topNew = (ph.topNew || [])[0], topDone = (ph.topDone || [])[0];
  const branchStory = [topNew ? 'Kenaikan tertinggi: ' + topNew.branch : '', topDone ? 'Penyelesaian tertinggi: ' + topDone.branch : '']
    .filter(Boolean).join('  ·  ');
  regionalNewsText(ctx, (branchStory || 'Belum ada cabang dengan progres tercatat.') + '  ·  ' + (ph.monthLabel || 'Periode bulan berjalan'),
    p3.x, p3.y + 250, p3.width, '400 23px Arial, sans-serif', '#728178', 2, 27);

  // 04. Tool PAR.
  const d = state.data || {}, categories = CATS.map(c => ({ name: c.title, count: (state.byCat[c.k] || []).length }))
    .sort((a, b) => b.count - a.count || collator.compare(a.name, b.name));
  const p4 = regionalNewsPanel(ctx, 1900, 470, 4, 'Tool PAR', 'Hasil pemeriksaan pinjaman dan kategori pemantauan', '#D0A12B');
  const parGap = 18, parCardW = (p4.width - parGap * 2) / 3;
  const scanned = Number(d.scanned) || 0, categorized = Array.isArray(d.rows) ? d.rows.length : 0;
  const categoryText = categories.filter(x => x.count > 0).slice(0, 2).map(x => x.name + ' (' + fmt(x.count) + ')').join(' dan ');
  regionalNewsMetric(ctx, p4.x, p4.y, parCardW, 'Pinjaman diperiksa', fmt(scanned), '#173F31', 'Data yang diproses');
  regionalNewsMetric(ctx, p4.x + parCardW + parGap, p4.y, parCardW, 'Masuk kategori', fmt(categorized), '#087A5A', (scanned ? nf1.format(categorized / scanned * 100) : '0') + '% dari yang diperiksa');
  regionalNewsMetric(ctx, p4.x + (parCardW + parGap) * 2, p4.y, parCardW, 'Total saldo pinjaman', rpShort(d.stats ? d.stats.balance : 0), '#173F31', 'Seluruh pinjaman diperiksa');
  regionalNewsText(ctx,
    categoryText ? 'Kategori dengan jumlah terbanyak: ' + categoryText + '.' : 'Belum ada pinjaman yang masuk kategori pada pemeriksaan ini.',
    p4.x, p4.y + 178, p4.width, '600 29px Arial, sans-serif', '#405B4E', 2, 35);
  regionalNewsText(ctx,
    'Total simpanan yang tercatat: ' + rpShort(d.stats ? d.stats.savings : 0) + '. Satu pinjaman dapat masuk lebih dari satu kategori, sehingga jumlah kategori tidak dijumlahkan sebagai pinjaman unik.',
    p4.x, p4.y + 260, p4.width, '400 23px Arial, sans-serif', '#728178', 2, 30);

  // Footer: catatan cakupan tetap terbaca namun jelas sebagai informasi sekunder.
  ctx.fillStyle = '#D8E4DC'; ctx.fillRect(REGIONAL_NEWS_M, 2400, REGIONAL_NEWS_W - REGIONAL_NEWS_M * 2, 3);
  const mainScope = state.user && state.user.tingkat === 'ADMINISTRATOR' ? 'Semua cabang' : 'Sesuai Wilayah Kerja akun';
  const progressScope = ph.scope || mainScope;
  regionalNewsSingleLine(ctx, 'Cakupan Rekening Bergerak, Analisa Cabang, dan Tool PAR: ' + mainScope,
    REGIONAL_NEWS_M, 2440, REGIONAL_NEWS_W - REGIONAL_NEWS_M * 2, 20, 18, '600', '#405B4E');
  regionalNewsSingleLine(ctx, 'Cakupan Progres Harian: ' + progressScope,
    REGIONAL_NEWS_M, 2480, REGIONAL_NEWS_W - REGIONAL_NEWS_M * 2, 20, 18, '600', '#405B4E');
  regionalNewsSingleLine(ctx, 'Dibuat ' + new Date().toLocaleString('id-ID', { dateStyle: 'medium', timeStyle: 'short' })
    + '  ·  Sumber: Rekening Bergerak, Analisa Cabang, PAR Baru/PAR Selesai, Rekening_Bergerak_Angsuran, dan Data_Upload.',
    REGIONAL_NEWS_M, 2520, REGIONAL_NEWS_W - REGIONAL_NEWS_M * 2, 20, 18, '400', '#728178');
  regionalNewsSingleLine(ctx, 'Ringkasan informasi. Untuk rincian dan daftar pinjaman, buka Dashboard masing-masing.',
    REGIONAL_NEWS_M, 2560, REGIONAL_NEWS_W - REGIONAL_NEWS_M * 2, 20, 18, '400', '#87958D');
  return canvas;
}

async function exportRegionalNewsImage() {
  if (!canDownloadImage()) {
    toast('Fitur Download Gambar belum diaktifkan untuk tingkat akun ini.');
    return;
  }
  const button = $('sum-news-img');
  if (!button) return;
  button.disabled = true;
  try {
    await prepareRegionalNewsData();
    if (!regionalNewsAllReady()) throw new Error('Data belum lengkap. Periksa keterangan pemuatan di halaman Ringkasan.');
    await nextFrame();
    const canvas = imgBuildRegionalNews();
    await imgDownload(canvas, 'Apa_Kabar_Regional_J_' + today() + '.png', 0);
    setNote('sum-news-msg', 'Gambar "Apa kabar Regional J?" berhasil diunduh dan siap dibagikan.', 'good');
  } catch (error) {
    setNote('sum-news-msg', error && error.message ? error.message : 'Gagal membuat gambar kabar Regional J.', 'bad');
  } finally {
    button.disabled = !regionalNewsAllReady();
  }
}

async function exportDashboardImage(which) {
  if (!canDownloadImage()) {
    toast('Fitur Download Gambar belum diaktifkan untuk tingkat akun ini.');
    return;
  }
  const config = {
    move: { button: 'rb-img', note: 'rb-msg', build: imgBuildRekeningBergerak, name: () => 'Ringkasan_Rekening_Bergerak_' + (state.rekening && state.rekening.year || today()) + '.png' },
    dash: { button: 'ddl-img', note: 'dmsg', build: imgBuildToolPAR, name: () => 'Ringkasan_Tool_PAR_' + today() + '.png' },
    progress: { button: 'ph-img', note: 'ph-msg', build: imgBuildProgresHarian, name: () => 'Ringkasan_Progres_Harian_' + (state.progress && state.progress.monthKey || today()) + '.png' },
    analisa: { button: 'an-img', note: 'an-msg', build: imgBuildAnalisaCabang, name: () => 'Ringkasan_Analisa_Cabang_' + (state.analisaFilters && state.analisaFilters.period || today()) + '.png' }
  }[which];
  if (!config) return;
  const btn = $(config.button);
  if (!btn) return;
  btn.disabled = true;
  try {
    setNote(config.note, which === 'progress' ? 'Menyusun gambar F4 landscape 300 DPI…' : 'Menyusun gambar ringkasan F4…');
    await nextFrame();
    const canvas = config.build();
    await imgDownload(canvas, config.name(), which === 'progress' ? 300 : 0);
    setNote(config.note, which === 'progress' ? 'Gambar PNG F4 landscape berhasil diunduh.' : 'Gambar PNG ukuran F4 berhasil diunduh.', 'good');
  } catch (err) {
    setNote(config.note, err && err.message ? err.message : 'Gagal membuat gambar ringkasan.', 'bad');
  } finally {
    const canEnable = which === 'move' ? !!(state.rekening && !state.rekening.settingRequired && state.rekening.settingDate)
      : which === 'dash' ? !!(state.data && state.data.scanned)
      : which === 'progress' ? !!state.progress
      : !!(state.analisaCabang && state.analisaFilters.period && analisaRowsFor(state.analisaFilters.period).length);
    btn.disabled = !canEnable;
  }
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
    $('rb-xlsx').disabled = true; $('rb-img').disabled = true;
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
  $('rb-xlsx').disabled = false; $('rb-img').disabled = false;
}

async function loadRekeningBergerak(options) {
  if (state.rekeningLoading || !state.user) return;
  state.rekeningLoading = true;
  $('rb-ref').disabled = true;
  $('rb-xlsx').disabled = true; $('rb-img').disabled = true;
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
    $('rb-xlsx').disabled = true; $('rb-img').disabled = true;
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
    invalidateRegionalNewsData();
    await loadRekeningBergerak();
    if (state.dashboardPage === 'sum' && canDownloadImage()) prepareRegionalNewsData();
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
    $('rb-xlsx').disabled = false; $('rb-img').disabled = false;
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

/* Grafik garis harian (PAR Baru vs PAR Selesai).
   Lebar mengikuti kotak grafik sehingga seluruh tanggal terlihat tanpa digeser. Sentuh atau arahkan
   penunjuk ke grafik untuk melihat nilai per tanggal. Isi dipasang lewat DOM/textContent. */
const PH_CHART_COLORS = { new: '#DC3545', done: '#198754' };

function progresNiceScale(maxValue, money) {
  if (!(maxValue > 0)) return { max: 4, ticks: 4 };
  const steps = money ? [1, 2, 2.5, 5, 10] : [1, 2, 5, 10];
  let best = null;
  [4, 5].forEach(ticks => {
    const raw = maxValue / ticks, mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const norm = raw / mag, pick = steps.find(s => norm <= s + 1e-9) || 10;
    let step = pick * mag;
    if (!money && step < 1) step = 1;
    const max = step * ticks;
    if (!best || max < best.max) best = { max, ticks };
  });
  return best;
}

function progresDayChart(hostId, daily, firstKey, secondKey, money, selectedDate) {
  const host = $(hostId);
  const data = Array.isArray(daily) ? daily : [];
  if (!host) return;
  host._phChart = { data, firstKey, secondKey, money, selectedDate: selectedDate || '' };
  host._phW = 0;
  progresDrawChart(host);
  if (!host._phObserver && 'ResizeObserver' in window) {
    host._phObserver = new ResizeObserver(debounce(() => progresDrawChart(host), 90));
    host._phObserver.observe(host);
  }
}

function progresDrawChart(host) {
  const cfg = host._phChart;
  if (!cfg) return;
  const { data, firstKey, secondKey, money, selectedDate } = cfg;
  if (!data.length) { put(host, h('div', { class: 'ph-chart-empty', text: 'Belum ada data pada bulan berjalan.' })); host._phW = 0; return; }
  const width = Math.round(host.clientWidth) || 0;
  if (host._phW === width && host.querySelector('svg')) return;       // lebar sama: tidak perlu menggambar ulang
  host._phW = width;
  const W = Math.max(260, width || 560), H = 240;
  const left = 54, right = 14, top = 14, bottom = 30, plotW = W - left - right, plotH = H - top - bottom;
  const n = data.length, pad = n > 1 ? 8 : 0;
  const val = (d, k) => Number(d[k]) || 0;
  const maxValue = Math.max(0, ...data.map(d => Math.max(val(d, firstKey), val(d, secondKey))));
  const scale = progresNiceScale(maxValue, money);
  const xAt = i => n > 1 ? left + pad + i * (plotW - pad * 2) / (n - 1) : left + plotW / 2;
  const yAt = v => top + plotH - Math.max(0, v) / scale.max * plotH;
  const stepPx = n > 1 ? (plotW - pad * 2) / (n - 1) : plotW;
  const markerR = Math.max(2.2, Math.min(4, stepPx * 0.26));
  const ns = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(ns, 'svg');
  svg.setAttribute('viewBox', '0 0 ' + W + ' ' + H);
  svg.setAttribute('width', W); svg.setAttribute('height', H);
  svg.setAttribute('class', 'ph-chart-svg');
  svg.setAttribute('role', 'img');
  svg.setAttribute('aria-label', money ? 'Grafik garis Out. PAR baru dibanding selesai per tanggal' : 'Grafik garis jumlah Client PAR baru dibanding selesai per tanggal');
  const add = (parent, tag, attrs, text) => {
    const e = document.createElementNS(ns, tag);
    Object.keys(attrs || {}).forEach(k => e.setAttribute(k, attrs[k]));
    if (text != null) e.textContent = text;
    parent.appendChild(e); return e;
  };
  const selIndex = selectedDate ? data.findIndex(d => d.date === selectedDate) : -1;
  if (selIndex > -1) {
    const bw = Math.max(14, stepPx);
    add(svg, 'rect', { x: xAt(selIndex) - bw / 2, y: top, width: bw, height: plotH, rx: 4, fill: '#e8f5ee' });
  }
  for (let t = 0; t <= scale.ticks; t++) {
    const y = top + plotH - plotH * t / scale.ticks;
    add(svg, 'line', { x1: left, y1: y, x2: W - right, y2: y, stroke: t === 0 ? '#9bb5a8' : '#e5eee9', 'stroke-width': t === 0 ? 1.2 : 1 });
    add(svg, 'text', { x: left - 8, y: y + 3.5, fill: '#66796f', 'font-size': 10, 'text-anchor': 'end' }, progresChartCompact(scale.max * t / scale.ticks, money));
  }
  // label tanggal: jarak minimal antar label dijaga agar tidak bertumpuk
  const every = Math.max(1, Math.ceil(30 / Math.max(stepPx, 1)));
  let lastLabelX = -Infinity;
  data.forEach((d, i) => {
    const x = xAt(i), isLast = i === n - 1;
    if (!(i % every === 0 || isLast) || x - lastLabelX < 22) return;
    lastLabelX = x;
    add(svg, 'text', { x, y: H - 10, fill: '#66796f', 'font-size': 10, 'text-anchor': 'middle' }, String(Number(String(d.date).slice(8, 10))));
  });
  const series = [{ key: firstKey, color: PH_CHART_COLORS.new, shape: 'circle' }, { key: secondKey, color: PH_CHART_COLORS.done, shape: 'diamond' }];
  const points = s => data.map((d, i) => xAt(i).toFixed(1) + ',' + yAt(val(d, s.key)).toFixed(1));
  const markerAt = (parent, s, x, y, r, attrs) => {
    if (s.shape === 'circle') return add(parent, 'circle', Object.assign({ cx: x, cy: y, r, fill: s.color, stroke: '#fff', 'stroke-width': 1.5 }, attrs));
    const q = r * 1.35;
    return add(parent, 'path', Object.assign({ d: 'M' + x + ' ' + (y - q) + 'L' + (x + q) + ' ' + y + 'L' + x + ' ' + (y + q) + 'L' + (x - q) + ' ' + y + 'Z', fill: s.color, stroke: '#fff', 'stroke-width': 1.5, 'stroke-linejoin': 'round' }, attrs));
  };
  series.forEach(s => {
    if (n > 1) add(svg, 'polyline', { points: points(s).join(' '), fill: 'none', stroke: s.color, 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' });
    data.forEach((d, i) => markerAt(svg, s, xAt(i), yAt(val(d, s.key)), markerR));
  });
  // lapisan interaktif: garis panduan, penanda aktif, dan kotak informasi
  const guide = add(svg, 'line', { y1: top, y2: top + plotH, stroke: '#64897a', 'stroke-width': 1, 'stroke-dasharray': '3 3', visibility: 'hidden' });
  const actives = series.map(s => markerAt(svg, s, 0, 0, markerR + 1.8, { visibility: 'hidden' }));
  const tipTitle = h('b', { class: 'ph-tip-date' });
  const tipNew = h('span', { class: 'ph-tip-val' }), tipDone = h('span', { class: 'ph-tip-val' }), tipNet = h('span', { class: 'ph-tip-val' });
  const tip = h('div', { class: 'ph-chart-tip', hidden: true, role: 'status' }, tipTitle,
    h('div', { class: 'ph-tip-row' }, h('i', { class: 'new' }), h('span', { text: 'PAR Baru' }), tipNew),
    h('div', { class: 'ph-tip-row' }, h('i', { class: 'done' }), h('span', { text: 'PAR Selesai' }), tipDone),
    h('div', { class: 'ph-tip-row net' }, h('span', { class: 'ph-tip-gap' }), h('span', { text: 'Selisih' }), tipNet));
  const showValue = v => money ? rp(v) : fmt(v) + ' Loan No';
  const hide = () => { tip.hidden = true; guide.setAttribute('visibility', 'hidden'); actives.forEach(a => a.setAttribute('visibility', 'hidden')); };
  const show = i => {
    const d = data[i], x = xAt(i), a = val(d, firstKey), b = val(d, secondKey), net = a - b;
    guide.setAttribute('x1', x); guide.setAttribute('x2', x); guide.setAttribute('visibility', 'visible');
    [a, b].forEach((v, j) => {
      const s = series[j], el = actives[j], y = yAt(v);
      if (s.shape === 'circle') { el.setAttribute('cx', x); el.setAttribute('cy', y); }
      else { const q = (markerR + 1.8) * 1.35; el.setAttribute('d', 'M' + x + ' ' + (y - q) + 'L' + (x + q) + ' ' + y + 'L' + x + ' ' + (y + q) + 'L' + (x - q) + ' ' + y + 'Z'); }
      el.setAttribute('visibility', 'visible');
    });
    tipTitle.textContent = String(d.date).split('-').reverse().join('-');
    tipNew.textContent = showValue(a); tipDone.textContent = showValue(b);
    tipNet.textContent = (net > 0 ? '+' : '') + showValue(net);
    tip.hidden = false;
    const tw = tip.offsetWidth || 170;
    tip.style.left = Math.max(0, Math.min(W - tw, x + 12 + tw > W ? x - 12 - tw : x + 12)) + 'px';
  };
  const indexFrom = e => {
    const r = svg.getBoundingClientRect(), x = (e.clientX - r.left) * (W / (r.width || W));
    return Math.max(0, Math.min(n - 1, n > 1 ? Math.round((x - left - pad) / stepPx) : 0));
  };
  svg.addEventListener('pointermove', e => { if (e.pointerType !== 'touch' || e.buttons) show(indexFrom(e)); });
  svg.addEventListener('pointerdown', e => show(indexFrom(e)));
  svg.addEventListener('pointerleave', e => { if (e.pointerType === 'mouse') hide(); });
  host.style.setProperty('--ph-chart-w', W + 'px');
  put(host, svg, tip);
  if (!progresDrawChart.docBound) {
    progresDrawChart.docBound = true;      // sentuh di luar grafik menutup kotak informasi
    document.addEventListener('pointerdown', e => {
      document.querySelectorAll('.ph-chart-host .ph-chart-tip').forEach(t => { if (!t.parentNode.contains(e.target)) { t.hidden = true; } });
      document.querySelectorAll('.ph-chart-host').forEach(hh => { if (!hh.contains(e.target)) {
        const g = hh.querySelector('line[stroke-dasharray]'); if (g) g.setAttribute('visibility', 'hidden');
        hh.querySelectorAll('.ph-chart-svg [visibility="visible"]').forEach(a => a.setAttribute('visibility', 'hidden'));
      } });
    });
  }
}

function progresRenderRanking(hostId, rows, kind) {
  const host = $(hostId);
  if (!rows || !rows.length) {
    put(host, h('div', { class: 'ph-empty-row', text: 'Belum ada progres pada filter yang dipilih.' })); return;
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
    : 'Seluruh data dengan Loan No pada cakupan filter terpilih memiliki kategori TOP UP yang sesuai dengan tiga kategori di atas.';
}

function selectProgresRecap(kind) {
  const active = kind === 'done' ? 'done' : 'new';
  const showNew = active === 'new';
  const newPanel = $('ph-new-recap-panel'), donePanel = $('ph-done-recap-panel');
  if (!newPanel || !donePanel) return;
  const changed = state.progressRecap !== active;
  state.progressRecap = active;
  newPanel.hidden = !showNew;
  donePanel.hidden = showNew;
  document.querySelectorAll('[data-ph-recap]').forEach(button => {
    const selected = button.dataset.phRecap === active;
    button.classList.toggle('is-active', selected);
    button.setAttribute('aria-pressed', selected ? 'true' : 'false');
    button.setAttribute('aria-selected', selected ? 'true' : 'false');
  });
  const note = $('ph-recap-context');
  if (note) note.textContent = showNew
    ? 'Menampilkan daftar Loan No yang masuk ke arsip PAR Baru.'
    : 'Menampilkan daftar Loan No yang masuk ke arsip PAR Selesai.';
  if (changed) state.progressDetailPage = 1;
  if (state.progress) loadProgresDetail(state.progressDetailPage || 1);
}

function progresRenderLoanDetails(result, options) {
  const kind = result.kind === 'done' ? 'done' : 'new';
  const table = $(kind === 'new' ? 'ph-new-recap-table' : 'ph-done-recap-table');
  const label = kind === 'new' ? 'PAR Baru' : 'PAR Selesai';
  if (!table) return;
  const headers = ['No.', 'Tanggal Progress', 'Cabang', 'Loan No', 'TOP UP', 'Total Balance'];
  const thead = h('thead', {}, h('tr', {}, ...headers.map(x => h('th', { text: x }))));
  const tbody = h('tbody', {});
  const records = result.records || [];
  if (!records.length) {
    tbody.append(h('tr', {}, h('td', { colspan: 6, class: 'ph-recap-empty-cell', text: 'Tidak ada daftar Loan No ' + label + ' pada filter ini.' })));
  } else {
    records.forEach((r, i) => {
      const number = (result.page - 1) * result.pageSize + i + 1;
      tbody.append(h('tr', {},
        h('td', { class: 'ph-recap-no', text: fmt(number) }),
        h('td', { text: progresFilterDateLabel(r.progressDate || '') }),
        h('th', { scope: 'row', text: r.branch || '-' }),
        h('td', { class: 'ph-loan-no', text: r.loanNo || '-' }),
        h('td', { text: r.topUp || '-' }),
        h('td', { class: 'ph-detail-amount', text: rp(r.totalBalance || 0) })
      ));
    });
  }
  put(table, thead, tbody);
  if (!options || options.updatePager !== false) {
    const total = Number(result.totalRecords) || 0;
    const page = Math.max(1, Number(result.page) || 1);
    const pageSize = Math.max(1, Number(result.pageSize) || 50);
    const pages = Math.max(1, Number(result.totalPages) || 1);
    const from = total ? ((page - 1) * pageSize + 1) : 0;
    const to = Math.min(total, page * pageSize);
    $('ph-detail-pager-info').textContent = total
      ? 'Menampilkan ' + fmt(from) + '–' + fmt(to) + ' dari ' + fmt(total) + ' Loan No ' + label + '.'
      : 'Tidak ada Loan No ' + label + ' untuk filter yang dipilih.';
    $('ph-detail-page-label').textContent = 'Halaman ' + fmt(page) + ' / ' + fmt(pages);
    $('ph-detail-prev').disabled = page <= 1;
    $('ph-detail-next').disabled = page >= pages;
    $('ph-detail-note').textContent = 'Daftar diurutkan berdasarkan cabang dan Loan No. Data dibatasi sesuai filter serta hak akses Wilayah Kerja.';
  }
}

function progresSetDetailLoading(kind, message) {
  const table = $(kind === 'done' ? 'ph-done-recap-table' : 'ph-new-recap-table');
  if (!table) return;
  const title = kind === 'done' ? 'PAR Selesai' : 'PAR Baru';
  put(table,
    h('thead', {}, h('tr', {}, h('th', { text: 'No.' }), h('th', { text: 'Tanggal Progress' }), h('th', { text: 'Cabang' }), h('th', { text: 'Loan No' }), h('th', { text: 'TOP UP' }), h('th', { text: 'Total Balance' }))),
    h('tbody', {}, h('tr', {}, h('td', { colspan: 6, class: 'ph-recap-empty-cell', text: message || 'Memuat daftar ' + title + '…' }))));
}

function progresDetailCacheKey(kind, page, filters, loadedAt) {
  const f = filters || {};
  return [loadedAt || 0, kind, f.branch || '', f.monthKey || '', f.progressDate || '', f.coordinatorNik || '', page, 50].join('|');
}

function progresFetchDetail(kind, page, filters, loadedAt) {
  const key = progresDetailCacheKey(kind, page, filters, loadedAt);
  if (state.progressDetailCache.has(key)) return Promise.resolve(state.progressDetailCache.get(key));
  if (state.progressDetailPending.has(key)) return state.progressDetailPending.get(key);

  const request = api('webProgresHarianDetail', {
    token: (session.get() || {}).token,
    kind: kind,
    branch: filters.branch || '',
    monthKey: filters.monthKey || '',
    progressDate: filters.progressDate || '',
    coordinatorNik: filters.coordinatorNik || '',
    page: page,
    pageSize: 50
  }, { retry: true }).then(result => {
    if (result && result.ok && state.progressLoadedAt === loadedAt) {
      state.progressDetailCache.set(key, result);
    }
    return result;
  }).finally(() => {
    if (state.progressDetailPending.get(key) === request) state.progressDetailPending.delete(key);
  });

  state.progressDetailPending.set(key, request);
  return request;
}

function progresPrefetchDetail(kind, page) {
  if (!state.user || !state.progress || !state.progressLoadedAt) return;
  const filters = Object.assign({}, state.progressFilters || { branch: '', monthKey: '', progressDate: '', coordinatorNik: '' });
  const loadedAt = state.progressLoadedAt;
  const key = progresDetailCacheKey(kind, page, filters, loadedAt);
  if (state.progressDetailCache.has(key) || state.progressDetailPending.has(key)) return;

  // Siapkan tabel tab yang belum aktif di belakang layar. Pager tab aktif tidak disentuh.
  progresFetchDetail(kind, page, filters, loadedAt).then(result => {
    if (!result || result.code === 'SESSION_REQUIRED' || !result.ok) return;
    if (state.progressLoadedAt !== loadedAt) return;
    progresRenderLoanDetails(result, { updatePager: false });
  }).catch(() => {
    // Prefetch bersifat tambahan; kegagalan tidak mengganggu tampilan tab aktif.
  });
}

async function loadProgresDetail(page) {
  if (!state.user || !state.progress) return;
  const kind = state.progressRecap === 'done' ? 'done' : 'new';
  const requestId = ++state.progressDetailSeq;
  const nextPage = Math.max(1, Number(page) || 1);
  state.progressDetailPage = nextPage;
  const filters = Object.assign({}, state.progressFilters || { branch: '', monthKey: '', progressDate: '', coordinatorNik: '' });
  const loadedAt = state.progressLoadedAt;
  const key = progresDetailCacheKey(kind, nextPage, filters, loadedAt);
  const cached = state.progressDetailCache.get(key);

  if (cached) {
    state.progressDetailResult = cached;
    state.progressDetailPage = cached.page || nextPage;
    progresRenderLoanDetails(cached);
    return;
  }

  progresSetDetailLoading(kind, 'Memuat daftar Loan No…');
  $('ph-detail-prev').disabled = true;
  $('ph-detail-next').disabled = true;
  $('ph-detail-pager-info').textContent = 'Memuat daftar ' + (kind === 'new' ? 'PAR Baru' : 'PAR Selesai') + '…';
  try {
    const result = await progresFetchDetail(kind, nextPage, filters, loadedAt);
    if (requestId !== state.progressDetailSeq || kind !== state.progressRecap || loadedAt !== state.progressLoadedAt) return;
    if (result.code === 'SESSION_REQUIRED') return logout(result.error);
    if (!result.ok) throw new Error(result.error || 'Gagal mengambil daftar Loan No.');
    state.progressDetailResult = result;
    state.progressDetailPage = result.page || nextPage;
    progresRenderLoanDetails(result);
    $('ph-detail-note').textContent = result.filterDescription
      ? 'Filter: ' + result.filterDescription + ' · ' + (result.scope || 'Sesuai Wilayah Kerja')
      : 'Daftar dibatasi sesuai filter dan hak akses Wilayah Kerja.';

    // Siapkan daftar tab pasangannya setelah tabel aktif siap. Perpindahan tab berikutnya
    // cukup membaca cache lokal; tidak perlu menunggu pemindaian sheet di server lagi.
    if (nextPage === 1) progresPrefetchDetail(kind === 'new' ? 'done' : 'new', 1);
  } catch (err) {
    if (requestId !== state.progressDetailSeq || kind !== state.progressRecap || loadedAt !== state.progressLoadedAt) return;
    progresSetDetailLoading(kind, err && err.message ? err.message : 'Gagal memuat daftar Loan No.');
    $('ph-detail-pager-info').textContent = 'Daftar Loan No belum dapat dimuat.';
    $('ph-detail-note').textContent = err && err.message ? err.message : 'Periksa koneksi lalu coba lagi.';
  }
}

function progresRenderBranchRecap(hostId, noteId, rows, kind, totals, isSingleDay) {
  const table = $(hostId);
  const countKey = kind === 'new' ? 'newCount' : 'doneCount';
  const balanceKey = kind === 'new' ? 'newBalance' : 'doneBalance';
  const list = (rows || []).slice().filter(r => (Number(r[countKey]) || 0) || (Number(r[balanceKey]) || 0))
    .sort((a, b) => (Number(b[balanceKey]) - Number(a[balanceKey]))
      || (Number(b[countKey]) - Number(a[countKey]))
      || String(a.branch || '').localeCompare(String(b.branch || ''), 'id', { sensitivity: 'base', numeric: true }));
  const thead = h('thead', {}, h('tr', {},
    h('th', { text: 'No.' }), h('th', { text: 'Cabang' }),
    h('th', { text: 'Client PAR · Loan No' }), h('th', { text: 'Out. PAR · Total Balance' })));
  const tbody = h('tbody', {});
  if (!list.length) {
    tbody.append(h('tr', {}, h('td', { colspan: 4, class: 'ph-recap-empty-cell', text: 'Tidak ada data PAR ' + (kind === 'new' ? 'Baru' : 'Selesai') + ' pada filter ini.' })));
  } else {
    list.forEach((r, i) => tbody.append(h('tr', {},
      h('td', { class: 'ph-recap-no', text: fmt(i + 1) }),
      h('th', { scope: 'row', text: r.branch || '-' }),
      h('td', { text: fmt(r[countKey] || 0) }),
      h('td', { text: rp(r[balanceKey] || 0) }))));
  }
  const totalCount = totals ? Number(totals[countKey]) || 0 : list.reduce((sum, r) => sum + (Number(r[countKey]) || 0), 0);
  const totalBalance = totals ? Number(totals[balanceKey]) || 0 : list.reduce((sum, r) => sum + (Number(r[balanceKey]) || 0), 0);
  const tfoot = h('tfoot', {}, h('tr', {},
    h('th', { colspan: 2, text: 'TOTAL ' + fmt(list.length) + ' CABANG' }),
    h('td', { text: fmt(totalCount) }), h('td', { text: rp(totalBalance) })));
  put(table, thead, tbody, tfoot);
  const period = isSingleDay ? 'tanggal progress terpilih' : 'periode terpilih';
  $(noteId).textContent = 'Diurutkan berdasarkan Out. PAR tertinggi. Menampilkan ' + fmt(list.length) + ' cabang pada ' + period + ' sesuai filter dan Wilayah Kerja.';
}

function progresFilterDateLabel(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
  return m ? m[1] + ' ' + (MONTHS_ID[Number(m[2]) - 1] || m[2]) + ' ' + m[3] : iso || '';
}

function progresFillSelect(id, firstLabel, rows, valueOf, labelOf, selected) {
  const select = $(id), options = [h('option', { value: '', text: firstLabel })];
  (rows || []).forEach(row => options.push(h('option', { value: valueOf(row), text: labelOf(row) })));
  select.replaceChildren(...options);
  const exists = Array.from(select.options).some(o => o.value === (selected || ''));
  select.value = exists ? (selected || '') : '';
  select.disabled = !(rows && rows.length);
  return select.value;
}

function paintProgresFilterControls(data) {
  const opts = data.filterOptions || {};
  const filters = data.filters || {};
  const month = progresFillSelect('ph-filter-month', 'Semua Bulan', opts.months || [], x => x.key, x => x.label, filters.monthKey || '');
  const branch = progresFillSelect('ph-filter-branch', 'Semua Cabang', opts.branches || [], x => x.name, x => x.name, filters.branch || '');
  const date = progresFillSelect('ph-filter-date', 'Semua Tanggal', opts.dates || [], x => x, progresFilterDateLabel, filters.progressDate || '');
  const coordinator = progresFillSelect('ph-filter-user', 'Semua Branch Coordinator', opts.coordinators || [], x => x.nik, x => x.name, filters.coordinatorNik || '');
  state.progressFilters = { branch: branch, monthKey: month, progressDate: date, coordinatorNik: coordinator };
  const userIsBC = state.user && String(state.user.jabatan || '').trim().toUpperCase().replace(/\s+/g, ' ') === 'BRANCH COORDINATOR';
  $('ph-filter-note').textContent = userIsBC
    ? 'Pilihan filter dibatasi pada cabang dalam Wilayah Kerja Anda. Data di luar wilayah tidak dikirim ke browser.'
    : 'Pilihan filter Branch Coordinator dan cabang mengikuti Wilayah Kerja akun yang sedang login.';
}

function changeProgresFilter(field, value) {
  state.progressFilterRevision++;
  state.progressFilters = state.progressFilters || { branch: '', monthKey: '', progressDate: '', coordinatorNik: '' };
  state.progressFilters[field] = value || '';
  if (field === 'coordinatorNik') {
    state.progressFilters.branch = '';
    state.progressFilters.monthKey = '';
    state.progressFilters.progressDate = '';
  } else if (field === 'monthKey') {
    state.progressFilters.progressDate = '';
  } else if (field === 'branch') {
    state.progressFilters.monthKey = '';
    state.progressFilters.progressDate = '';
  }
  loadProgresHarian({ filterChange: true });
}

function paintProgresHarian(data, preservePendingFilters) {
  if (!preservePendingFilters) paintProgresFilterControls(data);
  const selectedMonth = data.filters && data.filters.monthKey || '';
  const periodLabel = selectedMonth ? data.monthLabel : 'Semua Periode';
  $('ph-meta').textContent = (selectedMonth ? 'Periode ' + data.monthLabel : 'Seluruh riwayat') + ' · s.d. ' + fmtStamp(data.asOf)
    + ' · sumber: PAR Baru dan PAR Selesai' + (data.filterDescription ? ' · filter: ' + data.filterDescription : '');
  $('ph-scope').textContent = data.scope || 'Sesuai Wilayah Kerja';
  const isSingleDay = !!(data.filters && data.filters.progressDate);
  const periodPhrase = selectedMonth ? data.monthLabel : 'seluruh periode';
  $('ph-net-balance').previousElementSibling.textContent = isSingleDay ? 'Selisih Out. PAR pada tanggal terpilih' : 'Selisih Out. PAR · ' + periodPhrase;
  $('ph-net-count').previousElementSibling.textContent = isSingleDay ? 'Selisih Client PAR pada tanggal terpilih' : 'Selisih Client PAR · ' + periodPhrase;
  const topUpEyebrow = document.querySelector('#p-progress .ph-topup-panel .eyebrow');
  if (topUpEyebrow) topUpEyebrow.textContent = isSingleDay ? 'REKAP TANGGAL TERPILIH' : 'KOMULATIF · ' + periodLabel.toUpperCase();
  const chartCaptions = document.querySelectorAll('#p-progress .ph-chart-panel .ph-panel-head p');
  if (chartCaptions[0]) chartCaptions[0].textContent = isSingleDay ? 'Pergerakan pada ' + periodPhrase + '; tanggal terpilih ditandai.' : 'Nilai Total Balance per tanggal · ' + periodPhrase + '.';
  if (chartCaptions[1]) chartCaptions[1].textContent = isSingleDay ? 'Jumlah Loan No pada ' + periodPhrase + '; tanggal terpilih ditandai.' : 'Jumlah Loan No per tanggal · ' + periodPhrase + '.';
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
  const reduction = data.balanceReduction || { amount: 0, loans: 0, note: 'Riwayat perubahan angsuran belum tersedia.' };
  $('ph-movement-balance').textContent = rpShort(reduction.amount || 0);
  $('ph-movement-balance').title = rp(reduction.amount || 0);
  $('ph-movement-period').textContent = isSingleDay
    ? 'Kumulatif tanggal ' + progresFilterDateLabel(data.filters.progressDate)
    : selectedMonth ? 'Kumulatif seluruh tanggal · ' + data.monthLabel : 'Kumulatif semua bulan dan tanggal';
  $('ph-movement-note').textContent = fmt(reduction.loans || 0) + ' Loan No memiliki penurunan bersih positif · ' + (reduction.note || 'Nilai negatif menunjukkan kenaikan saldo.');
  progresRenderRanking('ph-top-new', data.topNew, 'new');
  progresRenderRanking('ph-top-done', data.topDone, 'done');
  const chartDaily = Array.isArray(data.dailyChart) ? data.dailyChart : data.daily;   // dailyChart: seluruh tanggal bulan berjalan
  const chartDate = data.filters && data.filters.progressDate;
  progresDayChart('ph-chart-out', chartDaily, 'newBalance', 'doneBalance', true, chartDate);
  progresDayChart('ph-chart-client', chartDaily, 'newCount', 'doneCount', false, chartDate);
  progresRenderTopUp(data);
  state.progressDetailSeq++;
  state.progressDetailCache.clear();
  state.progressDetailPending.clear();
  state.progressDetailResult = null;
  state.progressDetailPage = 1;
  // Hapus data lama pada tab tersembunyi agar hasil filter sebelumnya tidak sempat terlihat.
  progresSetDetailLoading('new', 'Menyiapkan daftar PAR Baru…');
  progresSetDetailLoading('done', 'Menyiapkan daftar PAR Selesai…');
  loadProgresDetail(1);

  const warnings = [];
  if (data.scopeMessage) warnings.push(data.scopeMessage);
  if (data.missingSheets && data.missingSheets.length) warnings.push('Sheet belum ditemukan: ' + data.missingSheets.join(', ') + '. Buat sheet melalui proses unggah versi terbaru.');
  if (data.missingDate) warnings.push(fmt(data.missingDate) + ' baris arsip memiliki Tanggal Progress kosong. Baris tersebut tidak masuk hitungan bulan berjalan.');
  if (data.invalidDate) warnings.push(fmt(data.invalidDate) + ' baris arsip memiliki Tanggal Progress tidak valid. Baris tersebut tidak masuk hitungan bulan berjalan.');
  const unknown = data.unknownTopUp || {};
  if (unknown.newCount || unknown.doneCount) warnings.push('Kategori TOP UP di luar NON TOP-UP, REGULER, dan KHUSUS: PAR Baru ' + fmt(unknown.newCount || 0) + ' Loan No; PAR Selesai ' + fmt(unknown.doneCount || 0) + ' Loan No.');
  const notice = $('ph-notice'); notice.hidden = !warnings.length; notice.textContent = warnings.join(' ');
  $('ph-sync').textContent = 'Terakhir diperiksa: ' + (data.generatedAt || '').replace('T', ' ') + (data.timezone ? ' · Zona waktu ' + data.timezone : '');
  $('ph-export').disabled = false; $('ph-img').disabled = false;
}

async function loadProgresHarian(options) {
  if (!state.user) return;
  if (state.progressLoading) {
    if (options && options.filterChange) state.progressReloadPending = true;
    return;
  }
  const automatic = !!(options && options.automatic);
  const requestedFilterRevision = state.progressFilterRevision;
  state.progressLoading = true;
  $('ph-refresh').disabled = true; $('ph-export').disabled = true; $('ph-img').disabled = true;
  if (!automatic) $('ph-sync').textContent = 'Mengambil arsip PAR Baru dan PAR Selesai…';
  setNote('ph-msg', '');
  try {
    if (automatic && !reserveAutoSync()) return;
    const filters = state.progressFilters || { branch: '', monthKey: '', progressDate: '', coordinatorNik: '' };
    const r = await api('webProgresHarian', {
      token: (session.get() || {}).token,
      branch: filters.branch || '', monthKey: filters.monthKey || '', progressDate: filters.progressDate || '', coordinatorNik: filters.coordinatorNik || ''
    }, { retry: true });
    if (r.code === 'SESSION_REQUIRED') return logout(r.error);
    if (!r.ok) {
      state.progress = null; state.progressDetailSeq++; state.progressDetailResult = null;
      state.progressDetailCache.clear(); state.progressDetailPending.clear();
      $('ph-meta').textContent = 'Progres Harian gagal dimuat.';
      $('ph-notice').hidden = false; $('ph-notice').textContent = r.error || 'Terjadi kesalahan saat mengambil data.';
      $('ph-sync').textContent = '';
      return;
    }
    state.progress = r; state.progressLoadedAt = Date.now();
    paintProgresHarian(r, requestedFilterRevision !== state.progressFilterRevision);
  } catch (err) {
    $('ph-notice').hidden = false;
    $('ph-notice').textContent = err && err.message ? err.message : 'Jaringan tidak dapat dihubungi. Periksa koneksi lalu muat ulang.';
  } finally {
    state.progressLoading = false; $('ph-refresh').disabled = false;
    $('ph-export').disabled = !state.progress; $('ph-img').disabled = !state.progress;
    if (state.progressReloadPending) {
      state.progressReloadPending = false;
      loadProgresHarian({ filterChange: true });
    }
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
    const sub = 'Periode ' + ((d.filters && d.filters.monthKey) ? d.monthLabel : 'Semua Bulan') + ' · s.d. ' + fmtStamp(d.asOf) + ' · ' + d.scope + (d.filterDescription ? ' · Filter: ' + d.filterDescription : '');
    const summaryRows = [
      ['PAR Baru · Client PAR', d.totals.newCount], ['PAR Baru · Out. PAR', d.totals.newBalance],
      ['PAR Selesai · Client PAR', d.totals.doneCount], ['PAR Selesai · Out. PAR', d.totals.doneBalance],
      ['Selisih Client PAR', d.netCount], ['Selisih Out. PAR', d.netBalance],
      ['Penurunan Outstanding PAR by Angsuran', (d.balanceReduction && d.balanceReduction.amount) || 0],
      ['Loan No dengan penurunan bersih positif', (d.balanceReduction && d.balanceReduction.loans) || 0],
      ['Filter Bulan', (d.filters && d.filters.monthKey) || 'Semua Bulan'],
      ['Filter Cabang', (d.filters && d.filters.branch) || 'Semua Cabang'],
      ['Filter Tanggal Progress', (d.filters && d.filters.progressDate) || 'Semua Tanggal'],
      ['Filter Branch Coordinator', (d.filters && d.filters.coordinatorName) || 'Semua Branch Coordinator']
    ];
    add('Ringkasan', progresExcelSheet('PROGRES HARIAN', sub, ['Indikator', 'Nilai'], summaryRows, [38, 24], [1]));
    add('Top 3 PAR Baru', progresExcelSheet('3 CABANG KENAIKAN TERTINGGI', sub, ['Cabang', 'Client PAR / Loan No', 'Out. PAR'],
      (d.topNew || []).map(x => [x.branch, x.newCount, x.newBalance]), [38, 24, 24], [1, 2]));
    add('Top 3 PAR Selesai', progresExcelSheet('3 CABANG PENYELESAIAN TERTINGGI', sub, ['Cabang', 'Client PAR / Loan No', 'Out. PAR'],
      (d.topDone || []).map(x => [x.branch, x.doneCount, x.doneBalance]), [38, 24, 24], [1, 2]));
    const dailyOutRows = (d.daily || []).map(x => [x.date.split('-').reverse().join('-'), x.newBalance, x.doneBalance, x.newBalance - x.doneBalance]);
    dailyOutRows.push(['TOTAL PERIODE', d.totals.newBalance, d.totals.doneBalance, d.netBalance]);
    const dailyClientRows = (d.daily || []).map(x => [x.date.split('-').reverse().join('-'), x.newCount, x.doneCount, x.newCount - x.doneCount]);
    dailyClientRows.push(['TOTAL PERIODE', d.totals.newCount, d.totals.doneCount, d.netCount]);
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
    const recapNewRows = (d.branches || []).slice().sort((a, b) => (Number(b.newBalance) - Number(a.newBalance)) || (Number(b.newCount) - Number(a.newCount)) || String(a.branch || '').localeCompare(String(b.branch || ''), 'id', { sensitivity: 'base', numeric: true }))
      .filter(x => (Number(x.newCount) || 0) || (Number(x.newBalance) || 0))
      .map((x, i) => [i + 1, x.branch, x.newCount, x.newBalance]);
    recapNewRows.push(['TOTAL', '', d.totals.newCount, d.totals.newBalance]);
    const recapDoneRows = (d.branches || []).slice().sort((a, b) => (Number(b.doneBalance) - Number(a.doneBalance)) || (Number(b.doneCount) - Number(a.doneCount)) || String(a.branch || '').localeCompare(String(b.branch || ''), 'id', { sensitivity: 'base', numeric: true }))
      .filter(x => (Number(x.doneCount) || 0) || (Number(x.doneBalance) || 0))
      .map((x, i) => [i + 1, x.branch, x.doneCount, x.doneBalance]);
    recapDoneRows.push(['TOTAL', '', d.totals.doneCount, d.totals.doneBalance]);
    add('PAR Baru', progresExcelSheet('REKAP PAR BARU PER CABANG', sub,
      ['No.', 'Cabang', 'Client PAR · Loan No', 'Out. PAR · Total Balance'], recapNewRows, [8, 38, 24, 26], [2, 3]));
    add('PAR Selesai', progresExcelSheet('REKAP PAR SELESAI PER CABANG', sub,
      ['No.', 'Cabang', 'Client PAR · Loan No', 'Out. PAR · Total Balance'], recapDoneRows, [8, 38, 24, 26], [2, 3]));
    add('Petunjuk', progresExcelSheet('PETUNJUK LAPORAN', sub, ['Istilah', 'Keterangan'], [
      ['Sumber', 'Sheet PAR Baru dan PAR Selesai.'],
      ['Periode', (d.filters && d.filters.monthKey) ? 'Tanggal Progress pada ' + d.monthLabel + ' sampai tanggal proses laporan.' : 'Seluruh Tanggal Progress yang tersedia sampai tanggal proses laporan.'],
      ['Client PAR', 'Jumlah Loan No tidak kosong pada arsip.'],
      ['Out. PAR', 'Penjumlahan Total Balance pada baris arsip.'],
      ['PAR Baru', 'Loan No yang tidak ditemukan pada Data_Upload ketika pemeriksaan upload.'],
      ['PAR Selesai', 'Loan No pada Data_Upload lama yang tidak ditemukan pada Excel terbaru.'],
      ['Penurunan Angsuran', 'Data_Upload lama dikurangi Excel baru. Nilai positif berarti saldo turun, nilai negatif berarti saldo naik. Kartu Penurunan Outstanding PAR by Angsuran menampilkan akumulasi bersih seluruh perubahan yang tercatat.'],
      ['Cakupan', 'Administrator melihat seluruh cabang; pengguna lain hanya cabang pada Wilayah_Kerja.'],
      ['TOP UP lainnya', 'Nilai TOP UP di luar NON TOP-UP, REGULER, dan KHUSUS tetap masuk total utama namun tidak masuk tabel tiga kategori.']
    ], [24, 100], []));
    XLSX.writeFile(wb, 'Progres_Harian_' + ((d.filters && d.filters.monthKey) || 'Semua_Period') + '.xlsx');
    setNote('ph-msg', 'Laporan Excel berhasil dibuat (10 sheet).', 'good');
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
// Label khusus dropdown agar tanggal ringkas dan konsisten: yyyy mmm dd.
// Contoh: 2025 Nov 01, 2026 Okt 31. Nilai option tetap ISO yyyy-MM-dd.
const AN_MONTHS_SHORT = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
function analisaDropdownDateLabel(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
  return m ? m[1] + ' ' + AN_MONTHS_SHORT[Number(m[2]) - 1] + ' ' + m[3] : iso || '—';
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
  // Tanpa filter tahun/bulan: tampilkan seluruh tanggal yang tersedia pada cakupan cabang terpilih.
  const filtered = (d.periods || []).filter(p =>
    f.branch === '*' || d.records.some(r => r.period === p && r.key === f.branch)
  );
  const period = keepPeriod && filtered.includes(f.period) ? f.period : (filtered[0] || '');
  f.period = period;
  const options = [h('option', { value: '', text: filtered.length ? 'Pilih tanggal analisa' : 'Tidak ada tanggal analisa' })]
    .concat(filtered.map(p => h('option', { value: p, text: analisaDropdownDateLabel(p) })));
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
  comparable.forEach(p => options.push(h('option', { value: p, text: analisaDropdownDateLabel(p) })));
  $('an-compare').replaceChildren(...options);
  $('an-compare').disabled = !f.period || !comparable.length;
  // Default pembanding = tanggal terdekat yang lebih lama pada cabang/cakupan yang sama.
  const earlier = comparable[0] || '';
  if (!analisaIsEarlierPeriod(f.compare, f.period) || !comparable.includes(f.compare)) f.compare = earlier;
  $('an-compare').value = f.compare;
}
function analisaInitFilters(data) {
  const f = state.analisaFilters;
  const branches = data.branches || [];
  const branchCurrent = branches.some(b => b.toLocaleUpperCase('id-ID') === f.branch) ? f.branch : '*';
  f.branch = branchCurrent;
  $('an-branch').replaceChildren(h('option', { value: '*', text: 'Semua Cabang' }), ...branches.map(b => h('option', { value: b.toLocaleUpperCase('id-ID'), text: b })));
  $('an-branch').value = f.branch;
  analisaUpdatePeriodOptions(true);
}
function analisaFormatDelta(metric, delta) {
  const sign = delta > 0 ? '+' : delta < 0 ? '−' : '';
  const absolute = Math.abs(delta);
  return sign + (metric.kind === 'money' ? rpShort(absolute) : fmt(absolute));
}
function analisaRenderCards(current, compare, currentPeriod, comparePeriod) {
  const host = $('an-kpis');
  put(host, ANALISA_METRICS.map(metric => {
    const value = current[metric.key] || 0;
    const hasCompare = !!compare && !!comparePeriod;
    const base = hasCompare ? (compare[metric.key] || 0) : null;
    const delta = hasCompare ? value - base : null;
    let changeText = '—', changeClass = 'neutral', pctText = 'Belum dibandingkan';
    if (hasCompare) {
      if (base === 0) {
        pctText = value === 0 ? '0,0% · tetap' : 'N/A · dasar 0';
        changeText = value === 0 ? '0' : analisaFormatDelta(metric, delta);
        changeClass = value > 0 ? 'up' : 'neutral';
      } else {
        const pct = delta / Math.abs(base) * 100;
        const sign = pct > 0 ? '+' : '';
        pctText = sign + nf1.format(pct) + '% · ' + (pct > 0 ? 'naik' : pct < 0 ? 'turun' : 'tetap');
        changeText = analisaFormatDelta(metric, delta);
        changeClass = pct > 0 ? 'up' : pct < 0 ? 'down' : 'neutral';
      }
    }
    const card = h('article', { class: 'an-kpi-card' },
      h('span', { class: 'an-kpi-label', text: metric.label }),
      h('div', { class: 'an-kpi-period an-kpi-current' },
        h('small', { text: 'Periode terbaru · ' + (currentPeriod ? analisaLabelDate(currentPeriod) : '—') }),
        h('strong', { class: 'an-kpi-value', text: analisaFormatMetric(metric, value), title: metric.kind === 'money' ? rp(value) : fmt(value) })),
      h('span', { class: 'an-kpi-hint', text: metric.hint }));

    if (hasCompare) {
      card.append(
        h('div', { class: 'an-kpi-baseline' },
          h('small', { text: 'Periode pembanding · ' + analisaLabelDate(comparePeriod) }),
          h('strong', { text: analisaFormatMetric(metric, base), title: metric.kind === 'money' ? rp(base) : fmt(base) })),
        h('div', { class: 'an-kpi-delta ' + changeClass },
          h('div', { class: 'an-kpi-delta-item' },
            h('small', { text: 'Perubahan nilai' }),
            h('b', { text: changeText, title: metric.kind === 'money' ? rp(delta) : fmt(delta) })),
          h('div', { class: 'an-kpi-delta-item' },
            h('small', { text: 'Persentase perubahan' }),
            h('b', { text: pctText }))));
    } else {
      card.append(h('div', { class: 'an-kpi-no-compare', text: 'Pilih tanggal pembanding yang lebih lama untuk melihat nilai awal, selisih, dan persentase perubahan.' }));
    }
    return card;
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
  analisaRenderCards(current, compare, period, f.compare);
  analisaRenderTable(curRows);
  $('an-table-caption').textContent = (f.branch === '*' ? 'Seluruh cabang dalam cakupan akses' : 'Cabang: ' + ($('an-branch').selectedOptions[0] ? $('an-branch').selectedOptions[0].textContent : f.branch))
    + ' · Periode ' + (period ? analisaLabelDate(period) : 'belum dipilih')
    + (f.compare ? ' · dibandingkan dengan ' + analisaLabelDate(f.compare) : ' · belum ada periode pembanding');
  const notices = [];
  if (d.scopeMessage) notices.push(d.scopeMessage);
  if (d.missingPeriod) notices.push(fmt(d.missingPeriod) + ' baris di sheet Analisa diabaikan karena Periode Analisa kosong/tidak valid.');
  if (period && !curRows.length) notices.push('Tidak ditemukan data untuk periode ' + analisaLabelDate(period) + '.');
  $('an-notice').hidden = !notices.length; $('an-notice').textContent = notices.join(' ');
  $('an-export').disabled = !period || !curRows.length; $('an-img').disabled = !period || !curRows.length;
}
async function loadAnalisaCabang(options) {
  if (state.analisaLoading || !state.user) return;
  const automatic = !!(options && options.automatic);
  state.analisaLoading = true;
  $('an-refresh').disabled = true; $('an-export').disabled = true; $('an-img').disabled = true;
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
    $('an-img').disabled = $('an-export').disabled;
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
      ['Indikator','Periode Pembanding','Periode Terpilih','Perubahan Nilai','Perubahan (%)'], summaryRows.map(row => [row[0], row[2], row[1], row[3], row[4]]),[28,22,22,22,18],[1,2,3,4]), 'Ringkasan');
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
   Setting: Enable Download (hanya ADMINISTRATOR; server memeriksa ulang)
   --------------------------------------------------------------------- */
function buildImageDownloadForm() {
  put($('image-download-level-list'), IMAGE_DOWNLOAD_USER_LEVELS.map(level =>
    h('label', { class: 'chk' },
      h('input', { type: 'checkbox', value: level }),
      h('span', null, h('b', { text: level }), h('small', { text: 'Izinkan menggunakan tombol Gambar F4.' }))
    )
  ));
  paintImageDownloadSettings();
}
async function saveImageDownloadAccess() {
  if (!state.user || state.user.tingkat !== 'ADMINISTRATOR') return setNote('image-download-msg', 'Hanya ADMINISTRATOR yang boleh mengubah Enable Download.', 'bad');
  if (!state.imageDownloadConfigLoaded) return setNote('image-download-msg', 'Konfigurasi belum berhasil dimuat. Muat ulang Dashboard lalu coba kembali.', 'bad');
  const levels = imageDownloadChecked();
  if (levels.join('|') === state.imageDownloadLevels.join('|')) return setNote('image-download-msg', 'Tidak ada perubahan.');
  const ok = await askConfirm('Simpan Enable Download?',
    (levels.length ? 'Tingkat yang dapat mengunduh Gambar F4: ' + levels.join(', ') + '. ' : 'Tombol Gambar F4 hanya tersedia untuk Administrator. ')
      + 'Hak akses unduh Excel tidak berubah.', 'Simpan');
  if (!ok) return;
  $('image-download-save').disabled = true;
  setNote('image-download-msg', 'Menyimpan pengaturan…');
  const r = await api('webSetImageDownload', { token: (session.get() || {}).token, levels: JSON.stringify(levels) });
  if (r.code === 'SESSION_REQUIRED') return logout(r.error);
  if (!r.ok) { paintImageDownloadSettings(); return setNote('image-download-msg', r.error || 'Gagal menyimpan Enable Download.', 'bad'); }
  state.imageDownloadLevels = IMAGE_DOWNLOAD_USER_LEVELS.filter(level => Array.isArray(r.enabledLevels) && r.enabledLevels.indexOf(level) > -1);
  state.imageDownloadConfigLoaded = true;
  state.imageDownloadDirty = false;
  paintImageDownloadSettings();
  paintImageDownloadControls();
  setNote('image-download-msg', levels.length ? 'Tersimpan. Akses Gambar F4 aktif untuk Administrator dan ' + levels.join(', ') + '.' : 'Tersimpan. Gambar F4 hanya tersedia untuk Administrator.', 'good');
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
  if (['global', 'wilayah', 'enable-download'].indexOf(which) < 0) which = 'global';
  if (which === 'enable-download' && (!state.user || state.user.tingkat !== 'ADMINISTRATOR')) which = 'global';
  state.settingPage = which;
  document.querySelectorAll('[data-setpage]').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.setpage === which)));
  $('set-global').hidden = which !== 'global';
  $('set-wilayah').hidden = which !== 'wilayah';
  $('set-enable-download').hidden = which !== 'enable-download';
  if (which === 'wilayah' && (!state.wilayah || (!wilayahIsDirty() && Date.now() - state.wilayahLoadedAt > 60000))) loadWilayah();
  if (which === 'enable-download') { paintImageDownloadSettings(); setNote('image-download-msg', ''); }
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
  invalidateRegionalNewsData();
  state.rekening = null; state.rekeningLoadedAt = 0;
  state.progress = null; state.progressLoadedAt = 0;
  state.analisaCabang = null; state.analisaLoadedAt = 0;
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
  $('force-password-form').addEventListener('submit', submitInitialPasswordChange);
  ['nik', 'pw'].forEach(i => $(i).addEventListener('input', () => $(i).classList.remove('bad')));
  ['force-new-password', 'force-confirm-password'].forEach(i => $(i).addEventListener('input', () => { $(i).classList.remove('bad'); showPasswordChangeErr(''); }));
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
  const versionSearch = $('version-q');
  if (versionSearch) versionSearch.addEventListener('input', () => {
    const q = versionSearch.value.trim().toLocaleLowerCase('id');
    let visible = 0;
    document.querySelectorAll('#version-list .version-item').forEach(item => {
      const match = !q || item.textContent.toLocaleLowerCase('id').includes(q);
      item.hidden = !match;
      if (match) visible++;
    });
    $('version-empty').hidden = visible !== 0;
  });
  $('app').querySelector('.tabs').addEventListener('keydown', e => {
    const vis = TABS.filter(k => !$('t-' + k).hidden), i = vis.findIndex(k => $('t-' + k).getAttribute('aria-selected') === 'true');
    const to = e.key === 'ArrowRight' ? vis[(i + 1) % vis.length] : e.key === 'ArrowLeft' ? vis[(i - 1 + vis.length) % vis.length]
      : e.key === 'Home' ? vis[0] : e.key === 'End' ? vis[vis.length - 1] : null;
    if (to) { e.preventDefault(); selectTab(to, true); }
  });

  // Dashboard: Rekening Bergerak
  $('rb-ref').onclick = () => loadRekeningBergerak();
  $('rb-xlsx').onclick = exportRekeningBergerak;
  $('rb-img').onclick = () => exportDashboardImage('move');
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
  document.querySelectorAll('[data-ph-recap]').forEach(button => {
    button.addEventListener('click', () => selectProgresRecap(button.dataset.phRecap));
  });
  $('ph-detail-prev').onclick = () => loadProgresDetail(Math.max(1, state.progressDetailPage - 1));
  $('ph-detail-next').onclick = () => loadProgresDetail(state.progressDetailPage + 1);
  selectProgresRecap('new');
  $('ph-refresh').onclick = () => loadProgresHarian();
  $('ph-export').onclick = exportProgresHarian;
  $('ph-img').onclick = () => exportDashboardImage('progress');
  $('ph-filter-month').onchange = () => changeProgresFilter('monthKey', $('ph-filter-month').value);
  $('ph-filter-branch').onchange = () => changeProgresFilter('branch', $('ph-filter-branch').value);
  $('ph-filter-date').onchange = () => changeProgresFilter('progressDate', $('ph-filter-date').value);
  $('ph-filter-user').onchange = () => changeProgresFilter('coordinatorNik', $('ph-filter-user').value);
  $('ph-filter-reset').onclick = () => { state.progressFilterRevision++; state.progressFilters = { branch: '', monthKey: '', progressDate: '', coordinatorNik: '' }; loadProgresHarian({ filterChange: true }); };

  // Dashboard Analisa Cabang
  $('an-refresh').onclick = () => loadAnalisaCabang();
  $('an-export').onclick = exportAnalisaCabang;
  $('an-img').onclick = () => exportDashboardImage('analisa');
  $('an-branch').onchange = () => { state.analisaFilters.branch = $('an-branch').value; analisaUpdatePeriodOptions(false); analisaRenderDashboard(); };
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
  $('ddl-img').onclick = () => exportDashboardImage('dash');
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
  document.querySelectorAll('#p-sum .seg button[data-g]').forEach(b => { b.onclick = () => { state.sumBy = b.dataset.g; state.sumSort = { key: 'n', dir: -1 }; drawSummary(); }; });
  $('sq').addEventListener('input', debounce(() => { state.sumQ = $('sq').value; drawSummary(); }, 150));
  $('sum-news-img').onclick = exportRegionalNewsImage;
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
  buildImageDownloadForm();
  $('wjform').addEventListener('submit', e => { e.preventDefault(); saveWajib(); });
  $('image-download-form').addEventListener('submit', e => { e.preventDefault(); saveImageDownloadAccess(); });
  $('image-download-level-list').addEventListener('change', () => {
    state.imageDownloadDirty = imageDownloadChecked().join('|') !== (state.imageDownloadLevels || []).join('|');
    paintImageDownloadSettings();
    setNote('image-download-msg', '');
  });
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
  // Invalidasi sesi v10.6 satu kali. Ini membuat semua pengguna login ulang setelah fitur keamanan dipasang;
  // tab versi lama menerima event storage dan ikut keluar dari sesi lama.
  try { localStorage.removeItem(LEGACY_SESSION_KEY); } catch (e) { /* abaikan */ }
  try { indexedDB.deleteDatabase(CACHE_DB_OLD); } catch (e) { /* abaikan */ }   // hapus cache browser sebelum reset database
  try { indexedDB.deleteDatabase(CACHE_DB_OLDER); } catch (e) { /* abaikan */ } // hapus cache browser versi lebih lama
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
