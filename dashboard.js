/* =====================================================================
   Dashboard Operasional -- dashboard.js
   Urutan:  1 Konfigurasi · 2 Alat bantu · 3 API & sesi · 4 Tampilan umum
            5 Masuk · 6 Dashboard · 7 Rincian · 8 Ringkasan
            9 Baca Excel · 10 Unggah · 11 Unduh Excel · 11C Write-Off · 12 Mulai
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
const UP_FIRST_ROW = 5, UP_CHUNK = 3000, UP_MAX_MB = 40;
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
  user: null, sessionToken: '', loadSeq: 0, loading: false, lastCheck: 0, dataVersion: '',
  imageDownloadLevels: [], imageDownloadConfigLoaded: false, imageDownloadDirty: false, settingPage: 'global', woPage: 'ewo', navOpen: '',
  data: null, byCat: {}, branches: new Map(), groups: {}, clientIdCounts: new Map(),
  cat: 'a', q: '', branch: '', staf: '', sort: { label: '', dir: 1 }, page: 0, size: 50, view: [],
  sumBy: 'branch', sumSort: { key: 'n', dir: -1 }, sumQ: '', wjDirty: false, pola: '', polaDirty: false, polaLoading: false,
  wilayah: null, wilayahSelected: '', wilayahDraft: [], wilayahDrafts: Object.create(null), wilayahLoading: false, wilayahLoadedAt: 0,
  rekening: null, rekeningLoading: false, rekeningLoadedAt: 0, dashboardPage: 'sum',
  menuAccess: null, menuAt: 0, menuLoading: false,   // peta menu yang tampil untuk akun ini (Manajemen)
  progress: null, progressLoading: false, progressLoadedAt: 0, progressReloadPending: false,
  progressFilters: { branch: '', monthKey: '', progressDate: '', coordinatorNik: '' }, progressFilterRevision: 0,
  progressRecap: 'new', progressDetailPage: 1, progressDetailSeq: 0, progressDetailResult: null,
  progressDetailCache: new Map(), progressDetailPending: new Map(),
  analisaCabang: null, analisaLoading: false, analisaLoadedAt: 0, analisaFilters: { branch: '*', period: '', compare: '' },
  kb: kbNew(), newsLoading: false, newsPromise: null, newsLoadSeq: 0, newsReady: false,
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
const TABS = ['dashboard', 'wo', 'set', 'up', 'versions'];      // 'wo' = Write-Off (dropdown: Eksekusi, Approve, Report); Manajemen kini sub menu Setting
const WO_PAGES = ['ewo', 'awo', 'rwo'];                        // urutan sub menu Write-Off
// Versi data bersama: setiap balasan halaman Dashboard membawa dataVersion yang sama untuk semua halaman.
// Bila berubah (unggahan baru, Segarkan data, dsb.), data halaman lain di memori dianggap usang bersamaan
// dan dimuat ulang saat dibuka (dari cache server, jadi cepat). Halaman sumber tidak dimuat ulang.
function noteDataVersion(version, source) {
  const v = String(version || '');
  if (!v) return false;
  if (!state.dataVersion) { state.dataVersion = v; return false; }
  if (state.dataVersion === v) return false;
  state.dataVersion = v;
  invalidateDashboardPages(source);
  if (state.user && !$('p-dashboard').hidden && state.dashboardPage !== source) selectDashboardPage(state.dashboardPage);
  return true;
}
function invalidateDashboardPages(source) {
  if (source !== 'move') { state.rekening = null; state.rekeningLoadedAt = 0; }
  if (source !== 'progress') {
    state.progress = null; state.progressLoadedAt = 0;
    state.progressDetailSeq++; state.progressDetailResult = null;
    state.progressDetailCache.clear(); state.progressDetailPending.clear();
  }
  if (source !== 'dash') state.lastCheck = 0;          // Tool PAR memeriksa ulang ke server (jawaban ringan bila isi sama)
  state.analisaCabang = null; state.analisaLoadedAt = 0;
  if (source !== 'sum') invalidateRegionalNewsData();   // sumber 'sum' = Kabar Berita sendiri: jangan membatalkan dirinya
}

function selectDashboardPage(which) {
  if (which !== 'move' && which !== 'dash' && which !== 'sum' && which !== 'progress' && which !== 'analisa') which = 'move';
  if (state.user && !menuOn('dashboard.' + which)) which = menuFirstDash() || which;   // sub menu yang disembunyikan Manajemen
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
    paintKabarBerita();
    if (state.user) prepareRegionalNewsData();
  }
  if (which === 'progress' && state.user && !state.progressLoading
      && (!state.progress || Date.now() - state.progressLoadedAt > AUTO_SYNC_MS)) loadProgresHarian();
  if (which === 'analisa' && state.user && !state.analisaLoading
      && (!state.analisaCabang || Date.now() - state.analisaLoadedAt > AUTO_SYNC_MS)) loadAnalisaCabang();
}
function selectTab(which, focus) {
  if (state.user && !menuOn(which)) which = menuFirstTab();   // menu yang tidak tampil untuk akun ini (Manajemen / tingkat)
  TABS.forEach(k => {
    const on = k === which;
    $('t-' + k).setAttribute('aria-selected', String(on));
    $('t-' + k).tabIndex = on ? 0 : -1;
    $('p-' + k).hidden = !on;
  });
  if (which === 'dashboard') selectDashboardPage(state.dashboardPage || 'move');
  if (which === 'set') { loadPola(); setSettingPage(state.settingPage || 'global'); }
  if (which === 'wo') selectWoPage(state.woPage || 'ewo');
  state.navOpen = !navMobile() && NAV_GROUPS.indexOf(which) > -1 ? which : '';   // desktop: sub menu halaman aktif terbuka; HP: kotak ditutup
  navPaint();
  if (focus) $('t-' + which).focus();
}

/* ---------- Sub menu di sidebar (accordion): Dashboard, Write-Off, Setting ----------
   Desktop: sub menu terbuka ke bawah tepat di bawah menu induknya (satu grup terbuka, yaitu grup halaman yang sedang dipakai).
   HP: sub menu muncul sebagai kotak di atas bilah navigasi bawah (dipindah keluar dari <nav> supaya tidak terpotong). */
const NAV_GROUPS = ['dashboard', 'wo', 'set'];
const navMobile = () => window.matchMedia('(max-width:760px)').matches;
const navSubItems = k => [...document.querySelectorAll('#sub-' + k + ' .tab-sub-item')].filter(b => !b.hidden);
function menuFirstWo() { return WO_PAGES.find(k => menuOn('wo.' + k)) || 'ewo'; }
function selectWoPage(which) {
  if (WO_PAGES.indexOf(which) < 0) which = 'ewo';
  if (state.user && !menuOn('wo.' + which)) which = menuFirstWo();   // sub menu yang disembunyikan Manajemen
  state.woPage = which;
  WO_PAGES.forEach(k => { $('p-' + k).hidden = k !== which; });
  document.querySelectorAll('[data-wopage]').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.wopage === which)));
  if (which === 'ewo') woEnter();
  if (which === 'awo') woEnterApprove();
  if (which === 'rwo') woEnterReport();
}
// Klik sub menu: atur halamannya lalu buka menu utamanya (bisa dari menu utama mana pun)
function openDashPage(which) { state.dashboardPage = which; selectTab('dashboard'); }
function openWoPage(which) { state.woPage = which; selectTab('wo'); }
function openSetPage(which) { state.settingPage = which; selectTab('set'); }

function navPlace(k) {          // hanya HP: kotak sub menu di atas tombol induknya
  const m = $('sub-' + k), r = $('t-' + k).getBoundingClientRect();
  const mw = m.offsetWidth, mh = m.offsetHeight, vw = document.documentElement.clientWidth;
  m.style.left = Math.round(Math.min(Math.max(8, r.left + r.width / 2 - mw / 2), Math.max(8, vw - mw - 8))) + 'px';
  m.style.top = Math.round(Math.max(8, r.top - mh - 8)) + 'px';
}
function navPaint() {
  NAV_GROUPS.forEach(k => {
    const t = $('t-' + k), m = $('sub-' + k);
    const open = !t.hidden && state.navOpen === k && navSubItems(k).length > 0;
    m.hidden = !open;
    t.setAttribute('aria-expanded', String(open));
    if (open && navMobile()) navPlace(k);
  });
}
// Letakkan kotak sub menu: desktop = tepat di bawah menu induk di sidebar; HP = di luar <nav> (backdrop-filter bilah bawah memotong elemen fixed)
function navDock() {
  const mobile = navMobile();
  NAV_GROUPS.forEach(k => {
    const m = $('sub-' + k);
    if (mobile) { if (m.parentNode !== $('app')) $('app').appendChild(m); }
    else if (m.previousElementSibling !== $('t-' + k)) $('t-' + k).after(m);
  });
  state.navOpen = mobile ? '' : (NAV_GROUPS.find(k => !$('p-' + k).hidden) || '');
  navPaint();
}
function navParentClick(k) {
  if (navMobile()) { state.navOpen = state.navOpen === k ? '' : k; navPaint(); return; }
  if (!$('p-' + k).hidden) { state.navOpen = state.navOpen === k ? '' : k; navPaint(); }   // menu aktif: buka / tutup sub menunya
  else selectTab(k);                                                                      // pindah ke menu itu; sub menunya ikut terbuka
}
function bindNavGroups() {
  const nav = $('app').querySelector('.tabs');
  NAV_GROUPS.forEach(k => {
    const m = $('sub-' + k);
    m.addEventListener('click', e => {
      const b = e.target.closest('.tab-sub-item'); if (!b || b.hidden) return;
      if (b.dataset.dashpage) openDashPage(b.dataset.dashpage);
      else if (b.dataset.wopage) openWoPage(b.dataset.wopage);
      else if (b.dataset.setpage) openSetPage(b.dataset.setpage);
    });
    m.addEventListener('keydown', e => {
      const items = navSubItems(k), i = items.indexOf(document.activeElement);
      if (e.key === 'Escape' && navMobile()) { e.preventDefault(); state.navOpen = ''; navPaint(); $('t-' + k).focus(); return; }
      const to = e.key === 'ArrowDown' ? items[(i + 1) % items.length] : e.key === 'ArrowUp' ? items[(i - 1 + items.length) % items.length]
        : e.key === 'Home' ? items[0] : e.key === 'End' ? items[items.length - 1] : null;
      if (to) { e.preventDefault(); to.focus(); }
    });
    $('t-' + k).addEventListener('keydown', e => {
      if (e.key !== 'ArrowDown' || !navSubItems(k).length) return;
      e.preventDefault(); e.stopPropagation();
      state.navOpen = k; navPaint(); const items = navSubItems(k); (items.find(b => b.getAttribute('aria-pressed') === 'true') || items[0]).focus();
    });
  });
  document.addEventListener('click', e => {          // HP: klik di luar menutup kotak sub menu
    if (!navMobile() || !state.navOpen) return;
    if (e.target.closest('.tab-sub') || e.target.closest('.tab')) return;
    state.navOpen = ''; navPaint();
  });
  nav.addEventListener('scroll', () => { if (navMobile() && state.navOpen) { state.navOpen = ''; navPaint(); } }, { passive: true });
  const mq = window.matchMedia('(max-width:760px)');
  (mq.addEventListener ? mq.addEventListener.bind(mq, 'change') : mq.addListener.bind(mq))(navDock);
  window.addEventListener('resize', () => { if (navMobile() && state.navOpen) { state.navOpen = ''; navPaint(); } });
  navDock();
}

function paintUser(user) {
  const name = user.nama || user.nik || '';
  $('uname').textContent = name;
  $('ulevel').textContent = user.tingkat || '';
  $('uavatar').textContent = name.split(/\s+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('').toUpperCase();
  paintImageDownloadControls();
  paintImageDownloadSettings();
  menuApply(false);   // tombol menu & sub menu menurut tingkat + pengaturan Manajemen (hanya tampilan; penegakan sebenarnya di server)
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
  if (state.dashboardPage === 'sum' && state.user && !state.newsLoading && !state.newsReady) {
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
    menuRefresh();   // perubahan Manajemen berlaku tanpa login ulang (tidak menambah hitungan sinkron otomatis)
    if (!$('p-dashboard').hidden && state.dashboardPage === 'move') {
      if (!state.rekeningLoading) loadRekeningBergerak({ automatic: true });
    } else if (!$('p-dashboard').hidden && state.dashboardPage === 'progress') {
      if (!state.progressLoading) loadProgresHarian({ automatic: true });
    } else if (!$('p-dashboard').hidden && state.dashboardPage === 'analisa') {
      if (!state.analisaLoading) loadAnalisaCabang({ automatic: true });
    } else if (!$('p-dashboard').hidden && state.dashboardPage === 'sum') {
      if (!state.newsLoading && reserveAutoSync()) prepareRegionalNewsData({ force: true });
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
  state.loadSeq++; state.loading = false; state.user = null; state.dataVersion = ''; state.sessionToken = ''; state.data = null; state.byCat = {}; state.groups = {}; state.clientIdCounts = new Map(); state.wjDirty = false; state.pola = ''; state.polaDirty = false;
  state.imageDownloadLevels = []; state.imageDownloadConfigLoaded = false; state.imageDownloadDirty = false; state.settingPage = 'global'; state.woPage = 'ewo';
  session.clear();
  const cacheClear = cache.clear();
  state.wilayah = null; state.wilayahDraft = []; state.wilayahDrafts = Object.create(null); state.wilayahSelected = ''; setNote('wjmsg', '');
  state.rekening = null; state.rekeningLoadedAt = 0; state.rekeningLoading = false;
  state.progress = null; state.progressLoadedAt = 0; state.progressLoading = false; state.progressReloadPending = false;
  state.progressRecap = 'new'; state.progressDetailPage = 1; state.progressDetailSeq++; state.progressDetailResult = null;
  state.progressDetailCache.clear(); state.progressDetailPending.clear();
  state.progressFilters = { branch: '', monthKey: '', progressDate: '', coordinatorNik: '' }; state.progressFilterRevision++;
  state.analisaCabang = null; state.analisaLoadedAt = 0; state.analisaLoading = false;
  resetKabarBerita();
  woReset();
  state.menuAccess = null; state.menuAt = 0; state.menuLoading = false; menuForget(); mgReset();
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
  state.menuAccess = menuCached(user.nik); state.menuAt = 0; state.menuLoading = false;   // sidebar langsung benar; disegarkan dari server di bawah
  state.dataVersion = '';
  state.sessionToken = (session.get() || {}).token || '';
  state.imageDownloadLevels = []; state.imageDownloadConfigLoaded = false; state.imageDownloadDirty = false; state.settingPage = 'global'; state.woPage = 'ewo';
  paintUser(user);
  state.cat = 'a'; state.q = ''; state.branch = ''; state.staf = ''; state.page = 0; state.sort = { label: '', dir: 1 };
  resetUpload(); setNote('setmsg', ''); setNote('wjmsg', ''); setNote('rb-setting-msg', ''); setNote('pola-msg', ''); state.wjDirty = false; state.pola = ''; state.polaDirty = false;
  state.rekening = null; state.rekeningLoadedAt = 0;
  state.progress = null; state.progressLoadedAt = 0; state.progressLoading = false; state.progressReloadPending = false;
  state.progressRecap = 'new'; state.progressDetailPage = 1; state.progressDetailSeq++; state.progressDetailResult = null;
  state.progressDetailCache.clear(); state.progressDetailPending.clear();
  state.progressFilters = { branch: '', monthKey: '', progressDate: '', coordinatorNik: '' }; state.progressFilterRevision++;
  $('ph-filter-month').value = ''; $('ph-filter-branch').value = ''; $('ph-filter-date').value = ''; $('ph-filter-user').value = '';
  state.analisaCabang = null; state.analisaLoadedAt = 0; state.analisaLoading = false;
  state.analisaFilters = { branch: '*', period: '', compare: '' };
  resetKabarBerita();
  woReset();
  state.dashboardPage = 'sum';
  state.rbSearch = { branch: '', coordinator: '' };
  state.rbSort = { branch: { key: 'name', dir: 1 }, coordinator: { key: 'name', dir: 1 } };
  $('rb-branch-q').value = ''; $('rb-coordinator-q').value = '';
  selectRekeningTable('branch');
  selectTab('dashboard'); showView('app');
  startAutoSync();
  loadRekeningBergerak();
  loadDash(); // pertahankan preload Tool PAR dan nilai Setting seperti versi sebelumnya
  menuRefresh(true);
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
  noteDataVersion(r.dataVersion, 'dash');
  state.lastCheck = Date.now();
  state.imageDownloadConfigLoaded = Array.isArray(r.imageDownloadLevels);
  state.imageDownloadLevels = state.imageDownloadConfigLoaded
    ? IMAGE_DOWNLOAD_USER_LEVELS.filter(level => r.imageDownloadLevels.indexOf(level) > -1) : [];
  paintImageDownloadControls();
  paintImageDownloadSettings();
  if (r.user) {
    const levelChanged = !!state.user && state.user.tingkat !== r.user.tingkat;   // TINGKAT diubah di Akses_2 saat sesi berjalan
    state.user = r.user; paintUser(r.user); const s = session.get(); if (s) session.set(Object.assign(s, { user: r.user }));
    if (levelChanged) { state.menuAccess = null; state.menuAt = 0; menuForget(); mgReset(); menuRefresh(true); }
    menuApply(true);   // halaman yang sedang dibuka tetapi tidak lagi boleh tampil dialihkan
  }
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
  if (!$('p-sum').hidden) { drawSummary(); paintKabarBerita(); }
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
  // Data master Write-Off: "WO Reguler yyyy.xlsx" / "WO Khusus yyyy.xlsx" (sama dengan woParseFile_ di WriteOff.gs)
  const woM = /^WO[\s_-]+(REGULER|KHUSUS)[\s_-]+((?:19|20)\d{2})(?!\d)/i.exec(base);
  if (woM) {
    if (/EWO-[RK]-\d{4}-\d{9}/i.test(base)) throw new Error('File ini hasil unduhan Eksekusi Write-Off (memuat ID_EWO). Isi No. THC dan No. TAK lewat tombol Otomatis di menu Write-Off > Eksekusi Write-Off, bukan lewat Unggah Excel.');
    const code = woM[1].toUpperCase() === 'REGULER' ? 'R' : 'K';
    return { type: 'WO', code, jenis: code === 'R' ? 'Reguler' : 'Khusus', sheet: code === 'R' ? 'WO Reguler' : 'WO Khusus', tahun: Number(woM[2]) };
  }
  const hasAnalisa = /(?:^|[^A-Z0-9])ANALISA(?=[^A-Z0-9])/i.test(base);
  const hasPar = /(?:^|[^A-Z0-9])PAR(?=[^A-Z0-9])/i.test(base);
  if (hasAnalisa && hasPar) throw new Error('Nama file tidak boleh memuat kata Analisa dan PAR sekaligus. Gunakan satu jenis proses per file.');
  const re = hasAnalisa ? /(?:^|[^A-Z0-9])ANALISA(?=[^A-Z0-9]).*?(\d{2}-\d{2}-\d{4})$/i
    : /(?:^|[^A-Z0-9])PAR(?=[^A-Z0-9]).*?(\d{2}-\d{2}-\d{4})$/i;
  let m = re.exec(base);
  const type = hasAnalisa ? 'ANALISA' : 'PAR';
  if (!m) throw new Error('Nama file harus mengikuti salah satu pola: "Analisa dd-mm-yyyy.xlsx", "PAR [nama opsional] dd-mm-yyyy.xlsx", "WO Reguler yyyy.xlsx", atau "WO Khusus yyyy.xlsx".');
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
    pending = info.type === 'ANALISA' ? await readAnalisaExcel(file, info) : info.type === 'WO' ? await readWoExcel(file, info) : await readExcel(file);
    if (info.type === 'PAR') {
      pending.type = 'PAR';
      pending.clientKey = [file.name, file.size, file.lastModified].join('|');
    }
    if (pending.type === 'WO') {
      renderWoPreview(pending);
      setNote('upmsg', 'File dikenali sebagai data master Write-Off (' + pending.sheet + ', tahun ' + pending.tahun + '). Periksa pratinjau, lalu simpan. Data_Upload dan Dashboard tidak berubah.');
    } else if (pending.type === 'ANALISA') {
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
        token: tok, uploadId: id, seq, offset: sent, rows: JSON.stringify(part)
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
  if (r.dataVersion) state.dataVersion = r.dataVersion;
  invalidateDashboardPages('dash');
  pending = null; $('file').value = ''; $('upgo').disabled = true;
  loadDash();
  // Pemanasan cache Rekening Bergerak dan Progres Harian di latar belakang agar halaman terbuka cepat bagi semua pengguna.
  api('webRefreshData', { token: (session.get() || {}).token, mode: 'warm' }, { timeoutMs: 300000 });
  if (state.dashboardPage === 'sum') prepareRegionalNewsData();
}

// Administrator: naikkan versi data, hitung ulang semua halaman Dashboard, lalu muat ulang bersamaan.
// Dipakai setelah data diubah langsung di Google Sheets (bukan lewat unggah web).
async function refreshAllData() {
  if (!state.user || state.user.tingkat !== 'ADMINISTRATOR') return;
  const btn = $('refresh-all'); btn.disabled = true;
  const t0 = performance.now();
  const tick = setInterval(() => setNote('refresh-msg', 'Menyegarkan data semua halaman… ' + clock(performance.now() - t0) + ' (mohon tunggu; bisa sampai beberapa menit untuk data besar).'), 500);
  setNote('refresh-msg', 'Menyegarkan data semua halaman…');
  const r = await api('webRefreshData', { token: (session.get() || {}).token, mode: 'refresh' }, { timeoutMs: 300000 });
  clearInterval(tick); btn.disabled = false;
  if (r.code === 'SESSION_REQUIRED') return logout(r.error);
  if (!r.ok) return setNote('refresh-msg', r.error || 'Gagal menyegarkan data. Coba lagi.', 'bad');
  if (r.dataVersion) state.dataVersion = r.dataVersion;
  invalidateDashboardPages('');
  setNote('refresh-msg', 'Selesai dalam ' + clock(performance.now() - t0) + '. Semua halaman Dashboard kini memakai data yang sama dan dimuat ulang saat dibuka.', 'good');
  toast('Data dashboard disegarkan.');
  loadDash();
  if (state.dashboardPage === 'sum') prepareRegionalNewsData();
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
  else if (!$('p-dashboard').hidden && state.dashboardPage === 'sum') prepareRegionalNewsData();
}

async function sendUpload() {
  if (!pending) return;
  if (pending.type === 'ANALISA') return saveAnalisaUpload();
  if (pending.type === 'WO') return saveWoUpload();
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
   Kabar Berita Ringkasan (gambar potret, lima poin)
   Sumber tetap memakai endpoint yang sudah menerapkan Wilayah_Kerja.
   Progres diminta tanpa filter tampilan agar mewakili seluruh cakupan user.
   --------------------------------------------------------------------- */
const REGIONAL_NEWS_W = 1600;
const REGIONAL_NEWS_H = 2620;
const REGIONAL_NEWS_M = 80;

async function regionalNewsWaitForModule(isLoading, label) {
  const deadline = Date.now() + API_TIMEOUT_MS + 15000;
  while (isLoading() && Date.now() < deadline) await sleep(160);
  if (isLoading()) throw new Error(label + ' belum selesai dimuat. Coba buka kembali Kabar Berita setelah sinkronisasi selesai.');
}

/* ---------------------------------------------------------------------
   Data Kabar Berita: lapisan data MANDIRI (state.kb). Tidak memakai data/filter halaman Dashboard lain,
   selalu mengambil data terbaru dari server (cache server sudah dipanaskan, jadi ringan), dan semua angka
   untuk layar maupun gambar unduhan dihitung dari satu model yang sama (regionalNewsModel).
   --------------------------------------------------------------------- */
const KB_FRESH_MS = 60 * 1000;       // data berita dianggap segar selama 1 menit; "Muat ulang" memaksa ambil ulang
function kbNew() {
  return { branch: '', rekening: null, rekeningAt: 0, analisa: null, analisaAt: 0, tool: null, toolAt: 0, progress: Object.create(null), errors: {}, branchSig: '' };
}
const kbKey = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toUpperCase();
function kbBranchNames() {
  const kb = state.kb, map = new Map();
  const add = name => { const k = kbKey(name); if (k && !map.has(k)) map.set(k, String(name).replace(/\s+/g, ' ').trim()); };
  if (kb.rekening) (kb.rekening.r.branchRows || []).forEach(x => add(x.branch));
  if (kb.analisa) (kb.analisa.r.records || []).forEach(x => add(x.branch));
  if (kb.tool) kb.tool.branches.forEach(add);
  return [...map.values()].sort((a, b) => collator.compare(a, b));
}
function kbGuard(nik) {
  if (!state.user || state.user.nik !== nik) throw new Error('Pemuatan Kabar Berita dibatalkan karena sesi atau data berubah.');
}
function kbSession(r) {
  if (r.code === 'SESSION_REQUIRED') { logout(r.error); throw new Error('Sesi berakhir. Silakan login kembali.'); }
}
async function kbFetchRekening(nik) {
  const r = await api('webRekeningBergerak', { token: (session.get() || {}).token }, { retry: true });
  kbGuard(nik); kbSession(r);
  if (!r.ok) throw new Error(r.error || 'Data Rekening Bergerak tidak dapat dimuat.');
  noteDataVersion(r.dataVersion, 'sum');
  state.kb.rekening = { r }; state.kb.rekeningAt = Date.now();
}
async function kbFetchAnalisa(nik) {
  const r = await api('webAnalisaCabang', { token: (session.get() || {}).token }, { retry: true });
  kbGuard(nik); kbSession(r);
  if (!r.ok) throw new Error(r.error || 'Data Analisa Cabang tidak dapat dimuat.');
  state.kb.analisa = { r }; state.kb.analisaAt = Date.now();
  if (r.polaHitung) { state.pola = r.polaHitung === 'B' ? 'B' : 'A'; if (!state.polaDirty) paintPola(); }
}
async function kbFetchProgress(nik, branch) {
  const r = await api('webProgresHarian', { token: (session.get() || {}).token, branch: branch || '', monthKey: '', progressDate: '', coordinatorNik: '' }, { retry: true });
  kbGuard(nik); kbSession(r);
  if (!r.ok) throw new Error(r.error || 'Data Progres Harian tidak dapat dimuat untuk Kabar Berita.');
  noteDataVersion(r.dataVersion, 'sum');
  state.kb.progress[kbKey(branch)] = { r, at: Date.now() };
}
function kbToolFrom(d) {
  const branches = new Set(); d.rows.forEach(r => { if (r[F.BRANCH]) branches.add(r[F.BRANCH]); });
  return { d, ver: d.ver, branches };
}
const kbToolValid = d => !!(d && d.scopeSchema === 'wilayah-v1' && Array.isArray(d.rows) && typeof d.wajib === 'string');
async function kbFetchTool(nik, force) {
  const kb = state.kb;
  await sleep(0);   // beri kesempatan permintaan Tool PAR yang baru dimulai saat login agar kita bergabung, bukan mengunduh dua kali
  if (state.loading) await regionalNewsWaitForModule(() => state.loading, 'Tool PAR');
  kbGuard(nik);
  if (!force && kbToolValid(state.data) && Date.now() - state.lastCheck <= KB_FRESH_MS) {   // "Muat ulang" (force) selalu memeriksa ke server
    if (!kb.tool || kb.tool.d !== state.data) kb.tool = kbToolFrom(state.data);
    kb.toolAt = Date.now();
    return;
  }
  let have = kb.tool ? kb.tool.d : null;
  if (!have) {
    const saved = await cache.get();
    kbGuard(nik);
    if (saved && saved.nik === nik && kbToolValid(saved.data)) have = saved.data;
  }
  const tok = (session.get() || {}).token;
  let r = await api('webDashboard', { token: tok, ver: have ? have.ver : '' }, { retry: true });
  kbGuard(nik); kbSession(r);
  if (!r.ok || r.scopeSchema !== 'wilayah-v1' || (!r.same && typeof r.wajib !== 'string')) throw new Error(r.ok ? MSG_OLD_SERVER : (r.error || 'Data Tool PAR tidak dapat dimuat.'));
  noteDataVersion(r.dataVersion, 'sum');
  let d = r;
  if (r.same) {
    if (have && have.ver === r.ver) d = have;
    else {
      r = await api('webDashboard', { token: tok, ver: '' }, { retry: true });
      kbGuard(nik); kbSession(r);
      if (!r.ok || !kbToolValid(r)) throw new Error(r.ok ? MSG_OLD_SERVER : (r.error || 'Data Tool PAR tidak dapat dimuat.'));
      d = r;
    }
  }
  if (d !== have) {   // simpan ke cache lokal bersama agar halaman Tool PAR tidak perlu mengunduh ulang
    const keep = () => cache.set({ nik, data: d });
    if (window.requestIdleCallback) window.requestIdleCallback(keep, { timeout: 2000 }); else setTimeout(keep, 200);
  }
  if (!kb.tool || kb.tool.d !== d) kb.tool = kbToolFrom(d);
  kb.toolAt = Date.now();
}

function regionalNewsSourcesReady() {
  const kb = state.kb;
  return !!(kb.rekening && kb.tool && kb.analisa && kb.progress[kbKey(kb.branch)]);
}
function regionalNewsAllReady() {
  return !!state.newsReady && regionalNewsSourcesReady();
}
// Tandai data berita usang (dimuat ulang saat Ringkasan dibuka). Data lama tetap tampil sampai data baru tiba.
function invalidateRegionalNewsData() {
  const kb = state.kb;
  state.newsLoadSeq++;
  kb.rekeningAt = 0; kb.analisaAt = 0; kb.toolAt = 0; kb.progress = Object.create(null); kb.errors = {};
  state.newsLoading = false; state.newsPromise = null; state.newsReady = false;
  const button = $('sum-news-img');
  if (button) button.disabled = true;
}
// Hapus seluruh data berita (login/logout), termasuk pilihan filter.
function resetKabarBerita() {
  state.kb = kbNew();
  invalidateRegionalNewsData();
  kbFilterPaint();
  const root = $('kb'); if (root) root.replaceChildren();
}
// Filter cabang di layar: mengisi pilihan dari cabang yang tersedia untuk akun ini.
function kbFilterPaint() {
  const sel = $('kb-branch'); if (!sel) return;
  const names = kbBranchNames(), kb = state.kb;
  if (kb.branch && !names.some(n => kbKey(n) === kbKey(kb.branch)) && (kb.rekening || kb.analisa || kb.tool)) kb.branch = '';
  const sig = names.join('\u001f');
  if (sig !== kb.branchSig || sel.options.length !== names.length + 1) {
    kb.branchSig = sig;
    sel.replaceChildren(h('option', { value: '', text: 'Semua Cabang' }), ...names.map(n => h('option', { value: n, text: n })));
  }
  sel.value = kb.branch || '';
  const note = $('kb-filter-note');
  if (note) note.textContent = kb.branch ? 'Semua angka dan gambar yang diunduh mengikuti cabang ' + kb.branch + '.' : 'Menampilkan seluruh cabang dalam cakupan akun ini.';
}

// Keterangan "siap diunduh": menyebut cakupan cabang dan waktu pemeriksaan data paling lama di antara keempat sumber.
function kbReadyNote() {
  const kb = state.kb, entry = kb.progress[kbKey(kb.branch)];
  const at = Math.min(kb.rekeningAt || Date.now(), kb.analisaAt || Date.now(), kb.toolAt || Date.now(), entry && entry.at || Date.now());
  const stamp = new Date(at).toLocaleString('id-ID', { dateStyle: 'medium', timeStyle: 'short' });
  setNote('sum-news-msg', canDownloadImage() ? 'Kabar Berita siap diunduh · ' + (kb.branch ? 'Cabang ' + kb.branch : 'Semua cabang') + ' · data terbaru, diperiksa ' + stamp + '.' : '', 'good');
}

async function prepareRegionalNewsData(options) {
  if (!state.user) return;
  if (state.newsPromise) return state.newsPromise;
  const force = !!(options && options.force), kb = state.kb, branch = kb.branch, bk = kbKey(branch), nik = state.user.nik, now = Date.now();
  const fresh = (at, has) => !force && !!has && !!at && now - at <= KB_FRESH_MS;
  const jobs = [];
  if (!fresh(kb.rekeningAt, kb.rekening)) jobs.push(['Rekening Bergerak', () => kbFetchRekening(nik)]);
  if (!fresh(kb.toolAt, kb.tool)) jobs.push(['Tool PAR', () => kbFetchTool(nik, force)]);
  if (!fresh(kb.analisaAt, kb.analisa)) jobs.push(['Analisa Cabang', () => kbFetchAnalisa(nik)]);
  if (!fresh(kb.progress[bk] && kb.progress[bk].at, kb.progress[bk])) jobs.push(['Progres Harian', () => kbFetchProgress(nik, branch)]);
  const button = $('sum-news-img');
  if (!jobs.length) {          // semua sumber masih segar: tidak ada permintaan baru
    state.newsReady = regionalNewsSourcesReady();
    if (button) button.disabled = !state.newsReady;
    if (state.newsReady) kbReadyNote();
    kbFilterPaint(); paintKabarBerita();
    return;
  }
  const seq = ++state.newsLoadSeq;
  state.newsLoading = true; state.newsReady = false;
  if (button) button.disabled = true;
  jobs.forEach(([label]) => { delete kb.errors[label]; });
  setNote('sum-news-msg', canDownloadImage() ? 'Mengambil data terbaru Kabar Berita…' : '');
  kbFilterPaint(); paintKabarBerita();

  state.newsPromise = (async () => {
    const results = await Promise.all(jobs.map(async ([label, run]) => {
      try {
        await run();
        if (state.user && state.user.nik === nik) { kbFilterPaint(); paintKabarBerita(); }
        return { label, ok: true };
      } catch (error) {
        const message = error && error.message ? error.message : String(error);
        if (state.user && state.user.nik === nik) kb.errors[label] = message;
        return { label, ok: false, error: message };
      }
    }));
    if (seq !== state.newsLoadSeq || !state.user || state.user.nik !== nik) return;
    const failed = results.filter(x => !x.ok);
    const complete = !failed.length && regionalNewsSourcesReady();
    state.newsReady = complete;
    if (button) button.disabled = !complete;
    if (complete) {
      kbReadyNote();
    } else {
      const details = failed.map(x => x.label + ': ' + x.error);
      if (!failed.length) details.push('Masih ada sumber data yang belum tersedia.');
      setNote('sum-news-msg', 'Kabar Berita belum lengkap. ' + details.join(' '), 'bad');
    }
  })().finally(() => {
    if (seq === state.newsLoadSeq) {
      state.newsLoading = false;
      state.newsPromise = null;
      if (button) button.disabled = !state.newsReady;
      kbFilterPaint(); paintKabarBerita();
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
  const value = String(text == null ? '' : text);
  ctx.font = fontWeight + ' ' + size + 'px Arial, sans-serif';
  // Ukur dengan ukuran huruf yang sedang dicoba agar teks panjang benar-benar mengecil sebelum dipotong.
  while (size > min && ctx.measureText(value).width > maxWidth) {
    size--;
    ctx.font = fontWeight + ' ' + size + 'px Arial, sans-serif';
  }
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

// Pembanding Analisa: tanggal TERAKHIR pada bulan sebelumnya (mis. 31 Okt untuk data terbaru 5 Nov).
// Bila bulan sebelumnya tidak punya data, dipakai tanggal terakhir pada bulan terdekat sebelum bulan terbaru.
function regionalNewsComparePeriod(periods, period) {
  if (!period) return '';
  const month = period.slice(0, 7);
  return periods.filter(p => p.slice(0, 7) < month).sort((a, b) => b.localeCompare(a))[0] || '';
}
function regionalNewsPrevMonth(period) {
  const y = Number(period.slice(0, 4)), m = Number(period.slice(5, 7));
  const d = new Date(Date.UTC(y, m - 2, 1));
  return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0');
}
function regionalNewsAggregateAnalisa(branchKey) {
  const data = state.kb.analisa ? state.kb.analisa.r : {}, all = Array.isArray(data.records) ? data.records : [];
  const records = branchKey ? all.filter(r => kbKey(r.branch) === branchKey) : all;
  const periods = Array.isArray(data.periods) ? data.periods.slice().sort((a, b) => b.localeCompare(a)) : [];
  const period = periods[0] || all.map(r => r.period).filter(Boolean).sort().reverse()[0] || '';
  const currentRows = period ? records.filter(r => r.period === period) : [];
  const comparePeriod = regionalNewsComparePeriod(periods, period);
  const compareRows = comparePeriod ? records.filter(r => r.period === comparePeriod) : [];
  return {
    period, comparePeriod, currentRows, current: analisaAggregate(currentRows),
    compare: compareRows.length ? analisaAggregate(compareRows) : null,
    comparePrevMonth: !!comparePeriod && comparePeriod.slice(0, 7) === regionalNewsPrevMonth(period),
    topParBranch: currentRows.slice().sort((a, b) => (Number(b.outstandingPAR) || 0) - (Number(a.outstandingPAR) || 0))[0] || null
  };
}

/* ---------------------------------------------------------------------
   Kabar Berita (potret): lima poin yang mudah dibaca.
   1 Perkembangan Cabang (Analisa Cabang) · 2 Kondisi Permasalahan Cabang (Rekening Bergerak + Analisa PAR)
   3 Rekomendasi Penyelesaian (Tool PAR) · 4 Kondisi Sampai Hari Kemarin (Progres Harian)
   5 Top Up Khusus yang kembali PAR (Progres Harian)
   --------------------------------------------------------------------- */
const REGIONAL_NEWS_MAX_H = 5200;      // kanvas kerja; hasil akhir dipotong sesuai isi
const RN_GREEN = '#087A5A', RN_RED = '#B43D3D', RN_GREY = '#52695D', RN_DARK = '#123D2F';

// Perubahan nilai terbaru terhadap pembanding. goodWhenUp: true bila kenaikan itu kabar baik (false untuk indikator PAR).
function regionalNewsDelta(cur, base, goodWhenUp) {
  if (base == null) return { has: false };
  const c = Number(cur) || 0, b = Number(base) || 0, delta = c - b;
  const dir = delta > 0 ? 1 : delta < 0 ? -1 : 0;
  return { has: true, delta, dir, pct: b ? delta / Math.abs(b) * 100 : null, good: dir === 0 ? null : (dir > 0) === !!goodWhenUp };
}
function regionalNewsSigned(money, value) {
  const sign = value > 0 ? '+' : value < 0 ? '−' : '';
  return sign + (money ? rpShort(Math.abs(value)) : fmt(Math.abs(value)));
}
function regionalNewsTriangle(ctx, cx, cy, s, dir, color) {
  ctx.fillStyle = color;
  if (!dir) { ctx.fillRect(cx - s, cy - 3, s * 2, 6); return; }
  ctx.beginPath();
  if (dir > 0) { ctx.moveTo(cx, cy - s); ctx.lineTo(cx + s, cy + s * 0.8); ctx.lineTo(cx - s, cy + s * 0.8); }
  else { ctx.moveTo(cx, cy + s); ctx.lineTo(cx + s, cy - s * 0.8); ctx.lineTo(cx - s, cy - s * 0.8); }
  ctx.closePath(); ctx.fill();
}
// Kartu: nilai terbaru besar, lalu persentase perubahan dan nilai perubahan dengan penanda naik/turun.
// Kartu rekomendasi: jumlah Loan No besar, lalu nominal / Total Balance.
// ----- Kartu gambar: gradien bergaya banner Kabar Berita (sama dengan kartu di layar) -----
const RN_TONES = { '': ['#064E3B', '#07644A', '#0A7A5A'], good: ['#065F46', '#0A7A5A', '#12A37A'], bad: ['#7F1D1D', '#A52F2F', '#C24A3C'], warn: ['#5C4300', '#7D5C06', '#96700A'] };
const RN_LABEL_FONT = '700 21px Arial, sans-serif';
const RN_PADX = 24, RN_ICON = 52;
// Tinggi kartu (satu sumber untuk penggambaran dan perhitungan tata letak). lines = jumlah baris label (1-2).
const rnMetricH = lines => 170 + (lines - 1) * 26;
const rnChangeH = (lines, hasCompare, bar) => (hasCompare ? 238 : 206) + (lines - 1) * 26 + (bar ? 28 : 0);
const rnCatH = lines => 288 + (lines - 1) * 26;

function regionalNewsRoundPath(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
}
// Pecah teks menjadi baris menurut lebar (font harus sudah diatur) tanpa memotong.
function regionalNewsLines(ctx, text, maxWidth) {
  const words = String(text == null ? '' : text).replace(/\s+/g, ' ').trim().split(' ').filter(Boolean), lines = [];
  let line = '';
  words.forEach(word => {
    const test = line ? line + ' ' + word : word;
    if (line && ctx.measureText(test).width > maxWidth) { lines.push(line); line = word; } else line = test;
  });
  if (line) lines.push(line);
  return lines;
}
// Jumlah baris label yang dibutuhkan (maksimal 2) agar semua kartu pada satu baris punya tinggi label yang sama.
function regionalNewsLabelLines(ctx, labels, cardWidth) {
  ctx.font = RN_LABEL_FONT;
  return Math.min(2, Math.max(1, ...labels.map(t => regionalNewsLines(ctx, String(t).toLocaleUpperCase('id-ID'), cardWidth - 136).length)));
}
function regionalNewsDrawIcon(ctx, name, x, y, size, color) {
  const shapes = KB_ICONS[name];
  if (!shapes) return;
  ctx.save();
  ctx.translate(x, y); ctx.scale(size / 24, size / 24);
  ctx.strokeStyle = color || '#FFFFFF'; ctx.lineWidth = 1.9; ctx.lineCap = 'round'; ctx.lineJoin = 'round';
  shapes.forEach(([tag, a]) => {
    if (tag === 'path') { ctx.stroke(new Path2D(a.d)); return; }
    ctx.beginPath();
    if (tag === 'circle') ctx.arc(+a.cx, +a.cy, +a.r, 0, Math.PI * 2);
    else if (tag === 'ellipse') ctx.ellipse(+a.cx, +a.cy, +a.rx, +a.ry, 0, 0, Math.PI * 2);
    else if (tag === 'rect') {
      const rx = +a.x, ry = +a.y, rw = +a.width, rh = +a.height, r = +(a.rx || 0);
      ctx.moveTo(rx + r, ry); ctx.arcTo(rx + rw, ry, rx + rw, ry + rh, r); ctx.arcTo(rx + rw, ry + rh, rx, ry + rh, r);
      ctx.arcTo(rx, ry + rh, rx, ry, r); ctx.arcTo(rx, ry, rx + rw, ry, r); ctx.closePath();
    }
    ctx.stroke();
  });
  ctx.restore();
}
// Dasar kartu: bayangan, gradien, lingkaran dekoratif emas/putih, dan lencana ikon.
function regionalNewsToneCard(ctx, x, y, w, h, tone, icon) {
  const c = RN_TONES[tone] || RN_TONES[''];
  ctx.save();
  ctx.shadowColor = tone === 'bad' ? 'rgba(127,29,29,0.28)' : 'rgba(6,78,59,0.26)'; ctx.shadowBlur = 22; ctx.shadowOffsetY = 8;
  imgRoundRect(ctx, x, y, w, h, 24, c[1]);
  ctx.restore();
  ctx.save();
  regionalNewsRoundPath(ctx, x, y, w, h, 24); ctx.clip();
  const g = ctx.createLinearGradient(x, y, x + w, y + h);
  g.addColorStop(0, c[0]); g.addColorStop(0.58, c[1]); g.addColorStop(1, c[2]);
  ctx.fillStyle = g; ctx.fillRect(x, y, w, h);
  ctx.fillStyle = 'rgba(246,201,21,0.20)'; ctx.beginPath(); ctx.arc(x + w - 22, y + 10, 88, 0, Math.PI * 2); ctx.fill();
  ctx.fillStyle = 'rgba(255,255,255,0.08)'; ctx.beginPath(); ctx.arc(x + 14, y + h + 26, 92, 0, Math.PI * 2); ctx.fill();
  ctx.restore();
  if (icon) {
    const bx = x + w - RN_PADX - RN_ICON, by = y + 20;
    imgRoundRect(ctx, bx, by, RN_ICON, RN_ICON, 16, 'rgba(255,255,255,0.16)', 'rgba(255,255,255,0.26)');
    regionalNewsDrawIcon(ctx, icon, bx + 12, by + 12, 28);
  }
}
function regionalNewsCardHead(ctx, x, y, w, label, lines) {
  ctx.fillStyle = '#F6C915'; ctx.fillRect(x + RN_PADX, y + 24, 6, 20);
  ctx.font = RN_LABEL_FONT;
  imgWrapText(ctx, String(label).toLocaleUpperCase('id-ID'), x + RN_PADX + 16, y + 40, w - 136, 26, lines, 'rgba(255,255,255,0.88)');
}
// Kartu angka sederhana: label, nilai besar, keterangan.
function regionalNewsMetricCard(ctx, x, y, w, tone, icon, label, value, detail, lines) {
  const h = rnMetricH(lines);
  regionalNewsToneCard(ctx, x, y, w, h, tone, icon);
  regionalNewsCardHead(ctx, x, y, w, label, lines);
  const vy = y + 40 + (lines - 1) * 26 + 66;
  regionalNewsSingleLine(ctx, value == null ? '—' : value, x + RN_PADX, vy, w - RN_PADX * 2, 56, 34, '700', '#FFFFFF');
  if (detail) regionalNewsSingleLine(ctx, detail, x + RN_PADX, vy + 38, w - RN_PADX * 2, 22, 17, '400', 'rgba(255,255,255,0.84)');
  return h;
}
// Kartu perubahan: nilai terbaru + kapsul (segitiga naik/turun, persentase, nilai perubahan).
function regionalNewsChangeCard(ctx, x, y, w, label, valueText, info, money, icon, lines, word, achv) {
  const has = achv ? achv.has : info.has;
  const tone = achv ? (achv.has && !achv.na && achv.reached ? 'good' : '') : (info.has && info.good !== null ? (info.good ? 'good' : 'bad') : '');
  const h = rnChangeH(lines, has, !!achv && achv.has);
  regionalNewsToneCard(ctx, x, y, w, h, tone, icon);
  regionalNewsCardHead(ctx, x, y, w, label, lines);
  const vy = y + 40 + (lines - 1) * 26 + 66, cy = vy + 22, cx = x + RN_PADX;
  regionalNewsSingleLine(ctx, valueText, cx, vy, w - RN_PADX * 2, 56, 34, '700', '#FFFFFF');
  if (!has) {
    const text = 'Belum ada periode pembanding';
    ctx.font = '600 21px Arial, sans-serif';
    const cw = Math.min(w - RN_PADX * 2, ctx.measureText(text).width + 36);
    imgRoundRect(ctx, cx, cy, cw, 52, 16, 'rgba(255,255,255,0.16)', 'rgba(255,255,255,0.24)');
    regionalNewsSingleLine(ctx, text, cx + 18, cy + 33, cw - 36, 21, 16, '600', 'rgba(255,255,255,0.92)');
    return h;
  }
  if (achv) {   // Pola Hitung B: persentase capaian + nilai bulan pembanding + bilah capaian
    const pctA = achv.na ? 'N/A (dasar 0)' : nf1.format(achv.pct) + '%', diffA = 'Achv. ' + achv.prevLabel + ': ' + achv.prevText;
    ctx.font = '700 32px Arial, sans-serif'; const pwA = ctx.measureText(pctA).width;
    ctx.font = '600 21px Arial, sans-serif'; const dwA = ctx.measureText(diffA).width;
    const cwA = Math.min(w - RN_PADX * 2, Math.max(pwA, dwA) + 36);
    imgRoundRect(ctx, cx, cy, cwA, 84, 16, 'rgba(255,255,255,0.16)', 'rgba(255,255,255,0.24)');
    regionalNewsSingleLine(ctx, pctA, cx + 18, cy + 40, cwA - 36, 32, 22, '700', '#FFFFFF');
    regionalNewsSingleLine(ctx, diffA, cx + 18, cy + 70, cwA - 36, 21, 16, '600', 'rgba(255,255,255,0.94)');
    if (!achv.na) {
      const bw = w - RN_PADX * 2, by = cy + 84 + 14;
      imgRoundRect(ctx, cx, by, bw, 12, 6, 'rgba(255,255,255,0.24)');
      imgRoundRect(ctx, cx, by, Math.max(12, bw * Math.min(1, achv.pct / 100)), 12, 6, '#FFFFFF');
    }
    return h;
  }
  const pctText = info.pct == null ? 'N/A (dasar 0)' : (info.pct > 0 ? '+' : info.pct < 0 ? '−' : '') + nf1.format(Math.abs(info.pct)) + '%';
  const diffText = (word || 'Dev.') + ' ' + regionalNewsSigned(money, info.delta);
  ctx.font = '700 32px Arial, sans-serif'; const pw = ctx.measureText(pctText).width;
  ctx.font = '600 21px Arial, sans-serif'; const dw = ctx.measureText(diffText).width;
  const cw = Math.min(w - RN_PADX * 2, Math.max(pw + 54, dw) + 36);
  imgRoundRect(ctx, cx, cy, cw, 84, 16, 'rgba(255,255,255,0.16)', 'rgba(255,255,255,0.24)');
  regionalNewsTriangle(ctx, cx + 30, cy + 29, 11, info.dir, '#FFFFFF');
  regionalNewsSingleLine(ctx, pctText, cx + 54, cy + 40, cw - 54 - 18, 32, 22, '700', '#FFFFFF');
  regionalNewsSingleLine(ctx, diffText, cx + 18, cy + 70, cw - 36, 21, 16, '600', 'rgba(255,255,255,0.94)');
  return h;
}
// Kartu rekomendasi: jumlah Loan No besar, lalu nominal / Total Balance.
function regionalNewsCategoryCard(ctx, x, y, w, title, count, label, value, detail, tone, icon, lines) {
  const h = rnCatH(lines);
  regionalNewsToneCard(ctx, x, y, w, h, tone, icon);
  regionalNewsCardHead(ctx, x, y, w, title, lines);
  const cy = y + 40 + (lines - 1) * 26 + 84, num = fmt(count), numMax = w - RN_PADX * 2 - 120;
  regionalNewsSingleLine(ctx, num, x + RN_PADX, cy, numMax, 68, 42, '700', '#FFFFFF');
  const numW = Math.min(ctx.measureText(num).width, numMax);   // ctx.font masih ukuran angka setelah penyesuaian
  regionalNewsSingleLine(ctx, 'Loan No', x + RN_PADX + numW + 12, cy, 110, 26, 20, '600', 'rgba(255,255,255,0.84)');
  ctx.fillStyle = 'rgba(255,255,255,0.28)'; ctx.fillRect(x + RN_PADX, cy + 24, w - RN_PADX * 2, 2);
  regionalNewsSingleLine(ctx, String(label).toLocaleUpperCase('id-ID'), x + RN_PADX, cy + 24 + 38, w - RN_PADX * 2, 20, 17, '700', 'rgba(255,255,255,0.86)');
  regionalNewsSingleLine(ctx, value, x + RN_PADX, cy + 24 + 38 + 46, w - RN_PADX * 2, 40, 26, '700', '#FFFFFF');
  if (detail) regionalNewsSingleLine(ctx, detail, x + RN_PADX, cy + 24 + 38 + 46 + 34, w - RN_PADX * 2, 21, 17, '400', 'rgba(255,255,255,0.84)');
  return h;
}

// Grafik garis dengan penanda (lingkaran = PAR Baru, belah ketupat = PAR Selesai), gayanya sama dengan Progres Harian.
function regionalNewsLineChart(ctx, x, y, w, h, title, dates, up, down, money) {
  imgRoundRect(ctx, x, y, w, h, 20, '#FFFFFF', '#DDE8E1');
  regionalNewsSingleLine(ctx, title, x + 24, y + 42, w - 48, 26, 20, '700', '#173F31');
  const marker = (shape, cx, cy, r, color) => {
    ctx.fillStyle = color; ctx.strokeStyle = '#FFFFFF'; ctx.lineWidth = 2; ctx.lineJoin = 'round'; ctx.beginPath();
    if (shape === 'circle') ctx.arc(cx, cy, r, 0, Math.PI * 2);
    else { const q = r * 1.35; ctx.moveTo(cx, cy - q); ctx.lineTo(cx + q, cy); ctx.lineTo(cx, cy + q); ctx.lineTo(cx - q, cy); ctx.closePath(); }
    ctx.fill(); ctx.stroke();
  };
  const ly = y + 78;
  ctx.font = '600 19px Arial, sans-serif';
  const w1 = ctx.measureText('PAR Baru').width, x2 = x + 24 + 30 + 12 + w1 + 34;
  [[x + 24, 'circle', PH_CHART_COLORS.new, 'PAR Baru'], [x2, 'diamond', PH_CHART_COLORS.done, 'PAR Selesai']].forEach(([lx, shape, color, text]) => {
    ctx.strokeStyle = color; ctx.lineWidth = 3; ctx.lineCap = 'round'; ctx.beginPath(); ctx.moveTo(lx, ly - 6); ctx.lineTo(lx + 30, ly - 6); ctx.stroke();
    marker(shape, lx + 15, ly - 6, 5.5, color);
    ctx.font = '600 19px Arial, sans-serif'; ctx.fillStyle = '#52665C'; ctx.textAlign = 'left'; ctx.fillText(text, lx + 42, ly);
  });
  const n = dates.length;
  if (!n) {
    ctx.font = '400 22px Arial, sans-serif'; ctx.fillStyle = '#728178'; ctx.textAlign = 'center';
    ctx.fillText('Belum ada data harian.', x + w / 2, y + h / 2 + 20); ctx.textAlign = 'left';
    return;
  }
  const left = x + 100, right = x + w - 30, top = y + 116, bottom = y + h - 54, plotW = right - left, plotH = bottom - top;
  const scale = progresNiceScale(Math.max(0, ...up, ...down), money);
  const pad = n > 1 ? 14 : 0, stepPx = n > 1 ? (plotW - pad * 2) / (n - 1) : plotW;
  const xAt = i => n > 1 ? left + pad + i * stepPx : left + plotW / 2, yAt = v => bottom - Math.max(0, v) / scale.max * plotH;
  ctx.font = '400 17px Arial, sans-serif';
  for (let t = 0; t <= scale.ticks; t++) {
    const yy = bottom - plotH * t / scale.ticks;
    ctx.fillStyle = t === 0 ? '#9BB5A8' : '#E5EEE9'; ctx.fillRect(left, yy - (t === 0 ? 1 : 0.5), right - left, t === 0 ? 2 : 1);
    ctx.fillStyle = '#66796F'; ctx.textAlign = 'right'; ctx.fillText(progresChartCompact(scale.max * t / scale.ticks, money), left - 12, yy + 6);
  }
  ctx.textAlign = 'center';
  const every = Math.max(1, Math.ceil(44 / Math.max(stepPx, 1)));
  let lastX = -Infinity;
  dates.forEach((date, i) => {
    const xx = xAt(i);
    if ((n - 1 - i) % every !== 0 || xx - lastX < 34) return;   // hitung dari titik terakhir agar label tanggal terbaru selalu ada dan jaraknya rata
    lastX = xx; ctx.fillStyle = '#66796F'; ctx.fillText(String(Number(String(date).slice(8, 10))), xx, bottom + 30);
  });
  ctx.textAlign = 'left';
  const r = Math.max(4.5, Math.min(8, stepPx * 0.26));
  [[up, PH_CHART_COLORS.new, 'circle'], [down, PH_CHART_COLORS.done, 'diamond']].forEach(([vals, color, shape]) => {
    if (n > 1) {
      ctx.strokeStyle = color; ctx.lineWidth = 3.5; ctx.lineJoin = 'round'; ctx.lineCap = 'round'; ctx.beginPath();
      vals.forEach((v, i) => { if (i) ctx.lineTo(xAt(i), yAt(v)); else ctx.moveTo(xAt(i), yAt(v)); });
      ctx.stroke();
    }
    vals.forEach((v, i) => marker(shape, xAt(i), yAt(v), r, color));
  });
}
function regionalNewsSubhead(ctx, x, y, width, text, color) {
  ctx.fillStyle = color || RN_GREEN; ctx.fillRect(x, y - 22, 8, 28);
  regionalNewsSingleLine(ctx, text, x + 20, y, width - 20, 28, 22, '700', '#2F4A3F');
}
// Grafik batang berpasangan: Peningkatan (PAR Baru) vs Penyelesaian (PAR Selesai) per tanggal progress.
// Semua angka Kabar Berita dihitung di satu tempat agar gambar unduhan dan tampilan layar Ringkasan selalu sama.
// Aman dipanggil saat sebagian sumber data belum selesai dimuat (bagian itu dibiarkan kosong/nol).
function regionalNewsModel() {
  // ---- Data (semua dari state.kb; difilter menurut cabang terpilih)
  const kb = state.kb, branch = kb.branch || '', bk = kbKey(branch), keepBranch = name => !bk || kbKey(name) === bk;
  const rb = kb.rekening ? kb.rekening.r : {}, rbRows = (Array.isArray(rb.branchRows) ? rb.branchRows : []).filter(x => keepBranch(x.branch));
  const rbTotal = rbRows.reduce((s, r) => s + (Number(r.total) || 0), 0), rbActive = rbRows.reduce((s, r) => s + (Number(r.active) || 0), 0);
  const rbPct = rbTotal ? rbActive / rbTotal * 100 : 0, rbReady = !(rb.settingRequired || !rb.settingDate);
  // Tanggal data terbaru = tanggal pada nama file unggahan (mis. PAR 05-11-2026.xlsx); cadangan: tanggal unggah, lalu tanggal progres terakhir.
  const rbFile = /(\d{2})-(\d{2})-(\d{4})/.exec(String(rb.dataFile || ''));
  const rbAsOfIso = rbFile ? rbFile[3] + '-' + rbFile[2] + '-' + rbFile[1] : String(rb.dataUpdated || '').slice(0, 10);
  const an = regionalNewsAggregateAnalisa(bk), anReady = !!(an.period && an.currentRows.length);
  const anDelta = (key, goodWhenUp) => regionalNewsDelta(an.current[key], an.compare ? an.compare[key] : null, goodWhenUp);
  const metricOf = key => ANALISA_METRICS.find(m => m.key === key);
  const anValue = key => { const m = metricOf(key); return m.kind === 'money' ? rpShort(an.current[key] || 0) : fmt(an.current[key] || 0); };
  const anLabel = an.period
    ? analisaLabelDate(an.period) + (an.comparePeriod ? ' dibanding ' + (an.comparePrevMonth ? 'akhir bulan lalu, ' : '') + analisaLabelDate(an.comparePeriod) : '')
    : 'Belum ada periode analisa';
  // Pola Hitung (Setting): A = selisih (default), B = persentase capaian terhadap tanggal pembanding.
  const pola = kb.analisa && kb.analisa.r.polaHitung === 'B' ? 'B' : 'A';
  const anAchv = key => {
    if (!an.compare || !an.comparePeriod) return { has: false };
    const m = metricOf(key), base = Number(an.compare[key]) || 0, cur = Number(an.current[key]) || 0, money = m.kind === 'money';
    const prevLabel = MONTHS_ID[Number(an.comparePeriod.slice(5, 7)) - 1] + ' ' + an.comparePeriod.slice(2, 4), prevText = money ? rpShort(base) : fmt(base);
    if (!(base > 0)) return { has: true, na: true, prevLabel, prevText, money };
    const pct = Math.max(0, cur / base * 100);   // tidak pernah minus; bisa lebih dari 100%
    return { has: true, na: false, pct, reached: pct >= 100, prevLabel, prevText, money };
  };
  const sum = (rows, fn) => rows.reduce((s, r) => s + (Number(fn(r)) || 0), 0);
  const d = kb.tool ? kb.tool.d : {}, wajib = String(d.wajib || '');
  const toolRows = kb.tool ? d.rows.filter(r => keepBranch(r[F.BRANCH])) : [];
  const catRows = k => toolRows.filter(r => r[F.FLAGS].indexOf(k) > -1);
  const rowsA = catRows('a'), rows1 = catRows('1'), rowsS = catRows('s');
  const kali = r => r[F.CICILAN] > 0 ? Math.floor((r[F.SUKARELA] + (wajib.indexOf('a') > -1 ? r[F.SIMP] : 0)) / r[F.CICILAN]) : 0;
  const angKali = sum(rowsA, kali);
  const angNominal = sum(rowsA, r => kali(r) * (Number(r[F.CICILAN]) || 0));
  const phEntry = kb.progress[bk], ph = phEntry ? phEntry.r : {}, pt = ph.totals || {}, red = ph.balanceReduction || { amount: 0, loans: 0 };
  const newBalance = Number(pt.newBalance) || 0, doneBalance = Number(pt.doneBalance) || 0;
  const newCount = Number(pt.newCount) || 0, doneCount = Number(pt.doneCount) || 0;
  const daily = (Array.isArray(ph.dailyChart) ? ph.dailyChart : []).filter(x => x && x.date).slice().sort((a, b) => a.date.localeCompare(b.date));
  const lastDate = daily.length ? daily[daily.length - 1].date : (ph.asOf || '');
  const recent = daily.slice(-14), dd = iso => iso.slice(8, 10) + '/' + iso.slice(5, 7);
  const rbAsOfLabel = (rbAsOfIso || lastDate) ? analisaLabelDate(rbAsOfIso || lastDate) : 'tanggal terbaru';
  const rbSinceLabel = rb.settingDate ? analisaLabelDate(String(rb.settingDate).slice(0, 10)) : '—';
  const khusus = (Array.isArray(ph.topUp) ? ph.topUp : []).find(x => x.category === 'KHUSUS') || { newCount: 0, newBalance: 0 };
  const khCount = Number(khusus.newCount) || 0, khBalance = Number(khusus.newBalance) || 0;

  // ---- Sorotan (ringkasan satu kalimat per poin)
  const lead = [];
  if (anReady && an.compare) {
    const dl = anDelta('outstandingPAR', false);
    lead.push('Outstanding PAR ' + (dl.dir < 0 ? 'turun ' : dl.dir > 0 ? 'naik ' : 'tetap ') + (dl.dir ? rpShort(Math.abs(dl.delta)) + ' ' : '') + 'dibanding ' + analisaLabelDate(an.comparePeriod) + '.');
  }
  if (rbReady && kb.rekening) lead.push('Rekening bergerak ' + nf1.format(rbPct) + '% (' + fmt(rbActive) + ' dari ' + fmt(rbTotal) + ' rekening PAR).');
  const net = newBalance - doneBalance;
  if (phEntry) lead.push(net > 0 ? 'PAR Baru ' + rpShort(newBalance) + ' lebih tinggi ' + rpShort(net) + ' dari PAR Selesai ' + rpShort(doneBalance) + '.'
    : net < 0 ? 'PAR Selesai ' + rpShort(doneBalance) + ' lebih tinggi ' + rpShort(Math.abs(net)) + ' dari PAR Baru ' + rpShort(newBalance) + '.'
      : 'PAR Baru dan PAR Selesai seimbang pada ' + rpShort(newBalance) + '.');
  return {
    rb, rbRows, rbTotal, rbActive, rbPct, rbReady, an, anReady, anDelta, metricOf, anValue, anLabel, sum, catRows, rowsA, rows1, rowsS, angKali, angNominal, d, ph, pt, red, newBalance, doneBalance, newCount, doneCount, daily, lastDate, recent, dd, khusus, khCount, khBalance, lead, branch, rbAsOfLabel, rbSinceLabel, pola, anAchv,
    loaded: { rekening: !!kb.rekening, analisa: !!kb.analisa, tool: !!kb.tool, progress: !!phEntry }
  };
}

function imgBuildRegionalNews() {
  if (!canDownloadImage()) throw new Error('Fitur Download Gambar belum diaktifkan untuk tingkat akun ini.');
  if (!regionalNewsAllReady()) throw new Error('Empat sumber data belum lengkap. Tunggu proses pemuatan selesai, lalu coba kembali.');
  const W = REGIONAL_NEWS_W, M = REGIONAL_NEWS_M, CW = W - M * 2, GAP = 20, SECTION_GAP = 26, INNER = CW - 60;
  const col3 = (INNER - GAP * 2) / 3, col2 = (INNER - GAP) / 2;
  const canvas = document.createElement('canvas');
  canvas.width = W; canvas.height = REGIONAL_NEWS_MAX_H;
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('Browser tidak dapat menyiapkan gambar Kabar Berita.');
  ctx.fillStyle = '#EEF4F0'; ctx.fillRect(0, 0, W, REGIONAL_NEWS_MAX_H);

  const {
    rb, rbRows, rbTotal, rbActive, rbPct, rbReady, an, anReady, anDelta, metricOf, anValue, anLabel, sum, catRows, rowsA, rows1, rowsS, angKali, angNominal, d, ph, pt, red, newBalance, doneBalance, newCount, doneCount, daily, lastDate, recent, dd, khusus, khCount, khBalance, lead, branch, rbAsOfLabel, rbSinceLabel, pola, anAchv
  } = regionalNewsModel();
  const hasCmp = !!an.compare;
  const wrapH = (text, width, font, lh, maxLines) => { ctx.font = font; return Math.min(maxLines, Math.max(1, regionalNewsLines(ctx, text, width).length)) * lh; };

  // ---- Masthead (gradien + lingkaran dekoratif, seperti kartu judul di layar)
  const mg = ctx.createLinearGradient(0, 0, W, 300);
  mg.addColorStop(0, '#064E3B'); mg.addColorStop(0.6, '#07644A'); mg.addColorStop(1, '#0A7A5A');
  ctx.save(); ctx.beginPath(); ctx.rect(0, 0, W, 300); ctx.clip();
  ctx.fillStyle = mg; ctx.fillRect(0, 0, W, 300);
  ctx.fillStyle = 'rgba(246,201,21,0.16)'; ctx.beginPath(); ctx.arc(W - 70, 10, 210, 0, Math.PI * 2); ctx.fill();
  ctx.fillStyle = 'rgba(255,255,255,0.07)'; ctx.beginPath(); ctx.arc(W - 470, 320, 160, 0, Math.PI * 2); ctx.fill();
  ctx.restore();
  ctx.fillStyle = '#F6C915'; ctx.fillRect(M, 50, 14, 48);
  regionalNewsSingleLine(ctx, 'REGIONAL J  ·  KABAR OPERASIONAL', M + 34, 84, CW - 34, 22, 22, '700', '#BFE6D4');
  regionalNewsSingleLine(ctx, 'Kabar Berita', M, 184, CW, 88, 70, '700', '#FFFFFF');
  // Cakupan cabang yang sedang ditampilkan di layar (filter) selalu ikut tercetak.
  const scopeLabel = branch ? 'Cabang: ' + branch : 'Semua Cabang';
  ctx.font = '700 24px Arial, sans-serif';
  const pillW = Math.min(CW * 0.5, ctx.measureText(scopeLabel).width + 48);
  imgRoundRect(ctx, W - M - pillW, 150, pillW, 46, 23, '#F6C915', '#F6C915');
  regionalNewsSingleLine(ctx, scopeLabel, W - M - pillW + 24, 182, pillW - 48, 24, 16, '700', '#3B2E00');
  regionalNewsSingleLine(ctx, 'Lima poin yang perlu kita ketahui bersama.', M, 237, CW, 34, 30, '400', '#E2F1E8');
  regionalNewsSingleLine(ctx, new Date().toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }), M, 278, CW, 25, 22, '600', '#D6E9DF');

  // ---- Sorotan (kartu emas yang menonjol; kalimat dibungkus, tidak dipotong)
  const leadW = CW - 64 - 96, leadFont = '700 32px Arial, sans-serif';
  ctx.font = leadFont;
  const leadRows = lead.map(t => Math.min(2, Math.max(1, regionalNewsLines(ctx, t, leadW).length)));
  let y = 324;
  const leadH = 128 + leadRows.reduce((s, n) => s + n * 44 + 18, 0);
  ctx.save();
  ctx.shadowColor = 'rgba(176,134,10,0.30)'; ctx.shadowBlur = 24; ctx.shadowOffsetY = 8;
  imgRoundRect(ctx, M, y, CW, leadH, 26, '#FFEBA8');
  ctx.restore();
  ctx.save();
  regionalNewsRoundPath(ctx, M, y, CW, leadH, 26); ctx.clip();
  const lg = ctx.createLinearGradient(M, y, M + CW, y + leadH);
  lg.addColorStop(0, '#FFF6D1'); lg.addColorStop(1, '#FFE08A');
  ctx.fillStyle = lg; ctx.fillRect(M, y, CW, leadH);
  ctx.fillStyle = '#E2B83E'; ctx.fillRect(M, y, 16, leadH);
  ctx.fillStyle = 'rgba(255,255,255,0.35)'; ctx.beginPath(); ctx.arc(M + CW - 40, y - 10, 150, 0, Math.PI * 2); ctx.fill();
  ctx.restore();
  imgRoundRect(ctx, M, y, CW, leadH, 26, null, '#E2B83E');
  imgRoundRect(ctx, M + 44, y + 30, 64, 64, 20, '#F6C915');
  regionalNewsDrawIcon(ctx, 'megaphone', M + 44 + 14, y + 30 + 14, 36, '#3B2E00');
  regionalNewsSingleLine(ctx, 'SOROTAN', M + 128, y + 74, CW - 160, 36, 32, '800', '#5C4300');
  let by = y + 140;
  lead.forEach((text, i) => {
    ctx.fillStyle = '#F6C915'; ctx.beginPath(); ctx.arc(M + 62, by - 12, 19, 0, Math.PI * 2); ctx.fill();
    ctx.fillStyle = '#3B2E00'; ctx.font = '800 22px Arial, sans-serif'; ctx.textAlign = 'center'; ctx.fillText(String(i + 1), M + 62, by - 4); ctx.textAlign = 'left';
    regionalNewsText(ctx, text, M + 96, by, leadW, leadFont, '#3B2E00', 2, 44);
    by += leadRows[i] * 44 + 18;
  });
  y += leadH + SECTION_GAP + 4;

  // ---- 1. Perkembangan Cabang
  const keys1 = [...KB_DEV_KEYS, ...KB_ACHV_KEYS];
  const l1 = regionalNewsLabelLines(ctx, keys1.map(k => metricOf(k).label), col3);
  const note1 = kbLegendNote(pola);
  const n1 = wrapH(note1, INNER, '400 19px Arial, sans-serif', 28, 2);
  const chH1 = rnChangeH(l1, hasCmp, false), chH2 = rnChangeH(l1, hasCmp, pola === 'B' && hasCmp);   // baris bawah menyediakan ruang bilah capaian
  const c1 = anReady ? chH1 + GAP + chH2 + 18 + n1 + 4 : 130;
  let p = regionalNewsPanel(ctx, y, 170 + c1, 1, 'Perkembangan Cabang', 'Analisa Cabang · ' + anLabel, '#0A7A5A');
  if (anReady) {
    keys1.forEach((key, i) => {
      const m = metricOf(key), col = i % 3, row = Math.floor(i / 3);
      regionalNewsChangeCard(ctx, p.x + col * (col3 + GAP), p.y + (row ? chH1 + GAP : 0), col3, m.label, anValue(key), anDelta(key, true), m.kind === 'money', KB_METRIC_ICON[key], l1,
        KB_DEV_KEYS.includes(key) ? 'Dev.' : 'Achv.', pola === 'B' && KB_ACHV_KEYS.includes(key) ? anAchv(key) : null);
    });
    regionalNewsText(ctx, note1, p.x, p.y + chH1 + GAP + chH2 + 18 + 20, INNER, '400 19px Arial, sans-serif', '#728178', 2, 28);
  } else {
    regionalNewsText(ctx, 'Data Analisa Cabang belum tersedia. Unggah file Analisa melalui menu Import Data agar perkembangan cabang dapat ditampilkan.',
      p.x, p.y + 40, p.width, '400 30px Arial, sans-serif', '#53665D', 3, 38);
  }
  y += 170 + c1 + SECTION_GAP;

  // ---- 2. Kondisi Permasalahan Cabang
  const l2a = regionalNewsLabelLines(ctx, ['Total PAR', 'Rekening Bergerak', 'Persentase Bergerak'], col3);
  const l2b = regionalNewsLabelLines(ctx, ['Member PAR', 'Outstanding PAR'], col2);
  const rbH = rbReady ? rnMetricH(l2a) : rnMetricH(1), anH = anReady ? rnChangeH(l2b, hasCmp) : 80;
  const c2 = 40 + rbH + 30 + 40 + anH + 6;
  p = regionalNewsPanel(ctx, y, 170 + c2, 2, 'Kondisi Permasalahan Cabang', 'Rekening Bergerak dan Analisa PAR', '#B43D3D');
  regionalNewsSubhead(ctx, p.x, p.y + 20, p.width, rbReady ? 'Perkembangan Rekening Bergerak sampai ' + rbAsOfLabel + ', Data dihitung dari tanggal ' + rbSinceLabel + '.' : 'Rekening Bergerak', '#B43D3D');
  if (rbReady) {
    regionalNewsMetricCard(ctx, p.x, p.y + 40, col3, '', 'alert', 'Total PAR', fmt(rbTotal), 'Total Rekening PAR sampai ' + rbAsOfLabel, l2a);
    regionalNewsMetricCard(ctx, p.x + col3 + GAP, p.y + 40, col3, 'good', 'pulse', 'Rekening Bergerak', fmt(rbActive), 'Total Rekening Transaksi sampai ' + rbAsOfLabel, l2a);
    regionalNewsMetricCard(ctx, p.x + (col3 + GAP) * 2, p.y + 40, col3, 'good', 'gauge', 'Persentase Bergerak', nf1.format(rbPct) + '%', 'Dibanding total PAR', l2a);
  } else {
    regionalNewsMetricCard(ctx, p.x, p.y + 40, p.width, 'warn', 'alert', 'Status', 'Belum dihitung', 'Tanggal N belum ditetapkan oleh Administrator atau User 0', 1);
  }
  const rowB = p.y + 40 + rbH + 30;
  regionalNewsSubhead(ctx, p.x, rowB + 14, p.width, 'Analisa Cabang · PAR' + (anReady ? '  ·  ' + anLabel : ''), '#B43D3D');
  if (anReady) {
    regionalNewsChangeCard(ctx, p.x, rowB + 40, col2, 'Member PAR', anValue('memberPAR'), anDelta('memberPAR', false), false, 'users', l2b, 'Dev.');
    regionalNewsChangeCard(ctx, p.x + col2 + GAP, rowB + 40, col2, 'Outstanding PAR', anValue('outstandingPAR'), anDelta('outstandingPAR', false), true, 'alert', l2b, 'Dev.');
  } else {
    regionalNewsText(ctx, 'Data Analisa Cabang belum tersedia.', p.x, rowB + 70, p.width, '400 30px Arial, sans-serif', '#53665D', 2, 38);
  }
  y += 170 + c2 + SECTION_GAP;

  // ---- 3. Rekomendasi Penyelesaian
  const l3 = regionalNewsLabelLines(ctx, ['Angsuran By Sukarela', 'Pelunasan di Bawah 100.000', 'Pelunasan dari Simpanan'], col3);
  const note3 = 'Nominal angsuran = jumlah angsuran × Cicilan (Pokok + Margin). Satu pinjaman dapat masuk lebih dari satu kategori.';
  const n3 = wrapH(note3, INNER, '400 19px Arial, sans-serif', 28, 2), c3 = rnCatH(l3) + 18 + n3 + 4;
  p = regionalNewsPanel(ctx, y, 170 + c3, 3, 'Rekomendasi Penyelesaian', 'Tool PAR · ' + (branch ? 'Cabang ' + branch : fmt(Number(d.scanned) || 0) + ' pinjaman diperiksa'), '#D0A12B');
  regionalNewsCategoryCard(ctx, p.x, p.y, col3, 'Angsuran By Sukarela', rowsA.length, 'Nominal angsuran', rpShort(angNominal), fmt(angKali) + ' kali angsuran', 'good', 'calendar', l3);
  regionalNewsCategoryCard(ctx, p.x + col3 + GAP, p.y, col3, 'Pelunasan di Bawah 100.000', rows1.length, 'Total Balance', rpShort(sum(rows1, r => r[F.BALANCE])), 'Sisa pelunasan sampai Rp 100.000', 'warn', 'check', l3);
  regionalNewsCategoryCard(ctx, p.x + (col3 + GAP) * 2, p.y, col3, 'Pelunasan dari Simpanan', rowsS.length, 'Total Balance', rpShort(sum(rowsS, r => r[F.BALANCE])), 'Simpanan sudah menutup Total Balance', 'good', 'lock', l3);
  regionalNewsText(ctx, note3, p.x, p.y + rnCatH(l3) + 18 + 20, INNER, '400 19px Arial, sans-serif', '#728178', 2, 28);
  y += 170 + c3 + SECTION_GAP;

  // ---- 4. Kondisi Sampai Hari Kemarin
  const l4 = regionalNewsLabelLines(ctx, ['PAR Baru · Peningkatan', 'PAR Selesai · Penyelesaian', 'Penurunan Outstanding PAR By Angsuran'], col3);
  const chartH = 420, c4 = rnMetricH(l4) + 30 + chartH;
  p = regionalNewsPanel(ctx, y, 170 + c4, 4, 'Kondisi Sampai Hari Kemarin', 'Progres Harian · data sampai ' + (lastDate ? analisaLabelDate(lastDate) : '—') + ' · seluruh periode', '#2B6CB0');
  regionalNewsMetricCard(ctx, p.x, p.y, col3, 'bad', 'plus', 'PAR Baru · Peningkatan', rpShort(newBalance), fmt(newCount) + ' Member PAR', l4);
  regionalNewsMetricCard(ctx, p.x + col3 + GAP, p.y, col3, 'good', 'check', 'PAR Selesai · Penyelesaian', rpShort(doneBalance), fmt(doneCount) + ' Member PAR', l4);
  regionalNewsMetricCard(ctx, p.x + (col3 + GAP) * 2, p.y, col3, Number(red.amount) > 0 ? 'good' : Number(red.amount) < 0 ? 'warn' : '', 'down', 'Penurunan Outstanding PAR By Angsuran', rpShort(red.amount || 0), fmt(red.loans || 0) + ' Loan No turun', l4);
  const cy4 = p.y + rnMetricH(l4) + 30, ctxt = recent.length ? ' · ' + recent.length + ' tanggal terakhir' : '';
  regionalNewsLineChart(ctx, p.x, cy4, col2, chartH, 'Peningkatan vs Penyelesaian Out. PAR' + ctxt, recent.map(x => x.date),
    recent.map(x => Number(x.newBalance) || 0), recent.map(x => Number(x.doneBalance) || 0), true);
  regionalNewsLineChart(ctx, p.x + col2 + GAP, cy4, col2, chartH, 'Peningkatan vs Penyelesaian Member PAR' + ctxt, recent.map(x => x.date),
    recent.map(x => Number(x.newCount) || 0), recent.map(x => Number(x.doneCount) || 0), false);
  y += 170 + c4 + SECTION_GAP;

  // ---- 5. Top Up Khusus yang kembali PAR
  const l5 = regionalNewsLabelLines(ctx, ['Total PAR Baru Top Up Khusus', 'Outstanding PAR'], col2);
  const note5 = 'Top Up Khusus = pinjaman kategori TOP UP KHUSUS yang tercatat sebagai PAR Baru.';
  const n5 = wrapH(note5, INNER, '400 19px Arial, sans-serif', 28, 2), c5 = rnMetricH(l5) + 18 + n5 + 4;
  p = regionalNewsPanel(ctx, y, 170 + c5, 5, 'Top Up Khusus yang Kembali PAR', 'Progres Harian · PAR Baru kategori Top Up Khusus · seluruh periode', '#B43D3D');
  regionalNewsMetricCard(ctx, p.x, p.y, col2, 'bad', 'star', 'Total PAR Baru Top Up Khusus', fmt(khCount) + ' Loan No',
    (newCount ? nf1.format(khCount / newCount * 100) : '0') + '% dari seluruh PAR Baru', l5);
  regionalNewsMetricCard(ctx, p.x + col2 + GAP, p.y, col2, 'bad', 'coins', 'Outstanding PAR', rpShort(khBalance),
    (newBalance ? nf1.format(khBalance / newBalance * 100) : '0') + '% dari Out. PAR Baru', l5);
  regionalNewsText(ctx, note5, p.x, p.y + rnMetricH(l5) + 18 + 20, INNER, '400 19px Arial, sans-serif', '#728178', 2, 28);
  y += 170 + c5 + SECTION_GAP;

  // ---- Footer (kalimat dibungkus, tidak dipotong)
  ctx.fillStyle = '#D8E4DC'; ctx.fillRect(M, y, CW, 3);
  const mainScope = state.user && state.user.tingkat === 'ADMINISTRATOR' ? 'Semua cabang' : 'Sesuai Wilayah Kerja akun';
  const footer = [
    ['Cakupan Analisa Cabang, Rekening Bergerak, dan Tool PAR: ' + (branch ? 'Cabang ' + branch : mainScope), '600 20px Arial, sans-serif', '#405B4E'],
    ['Cakupan Progres Harian: ' + (ph.scope || mainScope), '600 20px Arial, sans-serif', '#405B4E'],
    ['Dibuat ' + new Date().toLocaleString('id-ID', { dateStyle: 'medium', timeStyle: 'short' }) + '  ·  Sumber: Analisa Cabang, Rekening Bergerak, Data_Upload, PAR Baru/PAR Selesai, dan Rekening_Bergerak_Angsuran.', '400 18px Arial, sans-serif', '#728178'],
    ['Ringkasan informasi. Untuk rincian dan daftar pinjaman, buka Dashboard masing-masing.', '400 18px Arial, sans-serif', '#87958D']
  ];
  let fy = y + 42;
  footer.forEach(([text, font, color]) => { fy = regionalNewsText(ctx, text, M, fy, CW, font, color, 2, 26) + 10; });
  const total = fy + 34;

  // Potong sesuai isi agar tidak ada ruang kosong di bawah.
  const out = document.createElement('canvas');
  out.width = W; out.height = total;
  const octx = out.getContext('2d');
  if (!octx) throw new Error('Browser tidak dapat menyiapkan gambar Kabar Berita.');
  octx.drawImage(canvas, 0, 0, W, total, 0, 0, W, total);
  return out;
}

/* ---------------------------------------------------------------------
   Kabar Berita di layar (halaman Ringkasan). Memakai regionalNewsModel() yang sama dengan gambar unduhan,
   jadi angkanya selalu identik. Tiap bagian terisi begitu sumber datanya tiba; yang belum siap tampil sebagai kerangka.
   --------------------------------------------------------------------- */
const SVG_NS = 'http://www.w3.org/2000/svg';
function svgEl(tag, attrs, ...kids) {
  const e = document.createElementNS(SVG_NS, tag);
  for (const k in attrs || {}) if (attrs[k] != null) e.setAttribute(k, attrs[k]);
  for (const c of kids.flat()) if (c != null && c !== false) e.append(c);
  return e;
}
// Ikon dekoratif kartu (garis sederhana 24x24, tidak membawa informasi sehingga disembunyikan dari pembaca layar).
// Ikon dekoratif kartu (garis sederhana 24x24, hanya hiasan sehingga disembunyikan dari pembaca layar). Dipasang lewat DOM.
const KB_ICONS = {
  "megaphone":[["path",{"d":"M3 10v4h3l6 4V6L6 10H3z"}],["path",{"d":"M15.5 9a4 4 0 010 6"}],["path",{"d":"M18.5 6a8 8 0 010 12"}]],
  "users":[["circle",{"cx":"9","cy":"8","r":"3"}],["path",{"d":"M3 20c0-3.3 2.7-6 6-6s6 2.7 6 6"}],["circle",{"cx":"17","cy":"9","r":"2.5"}],["path",{"d":"M16.2 14.2c2.7.4 4.8 2.7 4.8 5.8"}]],"coins":[["ellipse",{"cx":"12","cy":"6","rx":"7","ry":"3"}],["path",{"d":"M5 6v6c0 1.7 3.1 3 7 3s7-1.3 7-3V6M5 12v6c0 1.7 3.1 3 7 3s7-1.3 7-3v-6"}]],"send":[["path",{"d":"M7 17L17 7M9 7h8v8"}]],"wallet":[["rect",{"x":"3","y":"6","width":"18","height":"13","rx":"3"}],["path",{"d":"M3 10h18M16 14.5h2"}]],"repay":[["path",{"d":"M17 7L7 17M15 17H7V9"}]],"percent":[["path",{"d":"M19 5L5 19"}],["circle",{"cx":"7","cy":"7","r":"2.5"}],["circle",{"cx":"17","cy":"17","r":"2.5"}]],"alert":[["path",{"d":"M12 4l9 16H3z"}],["path",{"d":"M12 10v4M12 17.5v.5"}]],"pulse":[["path",{"d":"M3 12h4l3-8 4 16 3-8h4"}]],"gauge":[["circle",{"cx":"12","cy":"12","r":"9"}],["path",{"d":"M12 12V3M12 12l6.4 6.4"}]],"calendar":[["rect",{"x":"3","y":"5","width":"18","height":"16","rx":"3"}],["path",{"d":"M3 10h18M8 3v4M16 3v4M9 15l2 2 4-4"}]],"check":[["circle",{"cx":"12","cy":"12","r":"9"}],["path",{"d":"M8 12.5l3 3 5-6"}]],"lock":[["rect",{"x":"5","y":"11","width":"14","height":"10","rx":"2"}],["path",{"d":"M8 11V8a4 4 0 018 0v3"}]],"plus":[["circle",{"cx":"12","cy":"12","r":"9"}],["path",{"d":"M12 8v8M8 12h8"}]],"down":[["path",{"d":"M3 7l6 6 4-4 8 8M21 11v6h-6"}]],"star":[["path",{"d":"M12 3l2.7 5.6 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.9 1-6.1L3.2 9.5l6.1-.9z"}]]};
// Ikon tiap indikator Analisa; dipakai bersama oleh tampilan layar dan gambar unduhan agar selalu sama.
// Keterangan singkatan di bawah kartu Perkembangan Cabang (sama untuk layar dan gambar).
const kbLegendNote = pola => pola === 'B'
  ? 'Dev. = Deviation (Perubahan) · Achv. = Achievements (Capaian): nilai tanggal terbaru ÷ nilai akhir bulan sebelumnya × 100%.'
  : 'Dev. = Deviation (Perubahan) · Achv. = Achievements (Capaian), dihitung terhadap tanggal pembanding.';
// Baris atas Perkembangan Cabang memakai "Dev." (Deviation/Perubahan); baris bawah "Achv." (Achievements/Capaian).
const KB_DEV_KEYS = ['member', 'outstanding', 'saving'], KB_ACHV_KEYS = ['disbursement', 'repayment', 'margin'];
const KB_METRIC_ICON = { member: 'users', outstanding: 'coins', disbursement: 'send', saving: 'wallet', repayment: 'repay', margin: 'percent' };
function kbIconSvg(name) {
  const shapes = KB_ICONS[name];
  if (!shapes) return null;
  return svgEl('svg', { viewBox: '0 0 24 24', focusable: 'false' }, shapes.map(([tag, attrs]) => svgEl(tag, attrs)));
}
function kbIcon(name) {
  const svg = kbIconSvg(name);
  return svg ? h('span', { class: 'kb-ico', 'aria-hidden': 'true' }, svg) : null;
}
function kbDeltaView(info, money, word) {
  if (!info.has) return h('div', { class: 'kb-delta neutral' }, h('span', { class: 'kb-diff', text: 'Belum ada periode pembanding' }));
  const tone = info.good === null ? 'neutral' : info.good ? 'good' : 'bad';
  const shape = info.dir > 0 ? 'up' : info.dir < 0 ? 'down' : 'flat';
  const pct = info.pct == null ? 'N/A (dasar 0)' : (info.pct > 0 ? '+' : info.pct < 0 ? '−' : '') + nf1.format(Math.abs(info.pct)) + '%';
  return h('div', { class: 'kb-delta ' + tone + ' ' + shape },
    h('span', { class: 'kb-tri', 'aria-hidden': 'true' }),
    h('span', { class: 'kb-pct', text: pct }),
    h('span', { class: 'kb-diff', text: (word || 'Dev.') + ' ' + regionalNewsSigned(money, info.delta) }));
}
// Isi kartu Pola Hitung B: persentase capaian, nilai bulan pembanding, dan bilah capaian (tidak ada warna merah).
function kbAchvView(a) {
  if (!a.has) return h('div', { class: 'kb-delta neutral' }, h('span', { class: 'kb-diff', text: 'Belum ada periode pembanding' }));
  const pct = a.na ? 'N/A (dasar 0)' : nf1.format(a.pct) + '%';
  return h('div', { class: 'kb-achv-wrap' },
    h('div', { class: 'kb-delta achv' }, h('span', { class: 'kb-pct', text: pct }), h('span', { class: 'kb-diff', text: 'Achv. ' + a.prevLabel + ': ' + a.prevText })),
    a.na ? null : h('div', { class: 'kb-achv-bar', role: 'img', 'aria-label': 'Capaian ' + pct }, h('i', { style: 'width:' + Math.min(100, a.pct).toFixed(1) + '%' })));
}
function kbChangeCard(label, valueText, info, money, fullValue, icon, word, achv) {
  const tone = achv ? (achv.has && !achv.na && achv.reached ? 'good' : '') : (info.has && info.good !== null ? (info.good ? 'good' : 'bad') : '');
  return h('article', { class: 'kb-card' + (tone ? ' tone-' + tone : '') },
    kbIcon(icon),
    h('span', { class: 'kb-label', text: label }),
    h('strong', { class: 'kb-value', text: valueText, title: fullValue || null }),
    achv ? kbAchvView(achv) : kbDeltaView(info, money, word));
}
function kbMetricCard(label, valueText, note, tone, icon) {
  return h('article', { class: 'kb-card' + (tone ? ' tone-' + tone : '') },
    kbIcon(icon),
    h('span', { class: 'kb-label', text: label }),
    h('strong', { class: 'kb-value', text: valueText }),
    note ? h('span', { class: 'kb-note', text: note }) : null);
}
function kbCategoryCard(title, count, label, value, detail, tone, icon) {
  return h('article', { class: 'kb-card kb-cat' + (tone ? ' tone-' + tone : '') },
    kbIcon(icon),
    h('span', { class: 'kb-label', text: title }),
    h('div', { class: 'kb-count' }, h('b', { text: fmt(count) }), h('span', { text: 'Loan No' })),
    h('div', { class: 'kb-cat-sum' }, h('span', { class: 'kb-label', text: label }), h('strong', { class: 'kb-money', text: value })),
    detail ? h('span', { class: 'kb-note', text: detail }) : null);
}
// Grafik batang berpasangan (SVG): Peningkatan (PAR Baru) vs Penyelesaian (PAR Selesai) per tanggal progress.
// Bingkai grafik garis dengan penanda; gambarnya sendiri dibuat progresDayChart (sama persis dengan halaman Progres Harian).
function kbLineFigure(title, hostId) {
  return h('figure', { class: 'kb-chart' },
    h('figcaption', { text: title }),
    h('div', { class: 'ph-legend' }, h('span', null, h('i', { class: 'new' }), 'PAR Baru'), h('span', null, h('i', { class: 'done' }), 'PAR Selesai')),
    h('div', { id: hostId, class: 'ph-chart-host' }));
}
function kbSection(num, title, subtitle, accent, body, links) {
  return h('section', { class: 'kb-sec', style: '--kb-accent:' + accent },
    h('div', { class: 'kb-sec-head' },
      h('span', { class: 'kb-num', text: String(num).padStart(2, '0') }),
      h('div', { class: 'kb-sec-title' }, h('h3', { text: title }), h('p', { text: subtitle })),
      links && links.length ? h('div', { class: 'kb-links' }, links.map(([text, page]) =>
        h('button', { type: 'button', class: 'kb-more', onclick: () => selectDashboardPage(page), text: text + ' →' }))) : null),
    body);
}
const kbSkeleton = (n, two) => h('div', { class: 'kb-grid' + (two ? ' two' : ''), 'aria-hidden': 'true' },
  Array.from({ length: n }, () => h('div', { class: 'kb-card skel kb-skel' })));
const kbMessage = text => h('p', { class: 'kb-empty', text });
// Kerangka pemuatan, atau pesan galat bila sumbernya gagal dimuat.
const kbWait = (label, n, two) => state.kb.errors[label]
  ? kbMessage('Gagal memuat ' + label + ': ' + state.kb.errors[label] + ' Tekan "Muat ulang" untuk mencoba lagi.')
  : kbSkeleton(n, two);

function paintKabarBerita() {
  const root = $('kb');
  if (!root) return;
  if (!state.user) { root.replaceChildren(); return; }
  const m = regionalNewsModel(), L = m.loaded;
  const allLoaded = L.rekening && L.analisa && L.tool && L.progress;
  const stamp = (m.rb && m.rb.dataUpdated) || (m.d && m.d.updated) || '';
  const mainScope = state.user.tingkat === 'ADMINISTRATOR' ? 'Semua cabang' : 'Sesuai Wilayah Kerja akun';

  const masthead = h('header', { class: 'kb-mast' },
    h('div', { class: 'kb-mast-text' },
      h('span', { class: 'kb-eyebrow', text: 'REGIONAL J  ·  KABAR OPERASIONAL' }),
      h('h2', { class: 'kb-title', text: 'Kabar Berita' }),
      h('p', { class: 'kb-tagline', text: 'Lima poin yang perlu kita ketahui bersama.' })),
    h('div', { class: 'kb-mast-meta' },
      h('span', { class: 'kb-date', text: new Date().toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }) }),
      h('span', { class: 'kb-chip gold', text: m.branch ? 'Cabang: ' + m.branch : 'Semua Cabang' }),
      stamp ? h('span', { class: 'kb-chip', text: 'Data per ' + fmtStamp(stamp) }) : null,
      state.newsLoading && !allLoaded ? h('span', { class: 'kb-chip' }, h('span', { class: 'spin' }), 'Memuat data…') : null));

  const lead = m.lead.length
    ? h('section', { class: 'kb-lead', 'aria-label': 'Sorotan' },
      h('div', { class: 'kb-lead-head' }, h('span', { class: 'kb-lead-badge', 'aria-hidden': 'true' }, kbIconSvg('megaphone')), h('h3', { text: 'SOROTAN' })),
      h('ol', null, m.lead.map(t => h('li', { text: t }))))
    : null;

  // 1. Perkembangan Cabang
  const s1 = !L.analisa ? kbWait('Analisa Cabang', 6)
    : !m.anReady ? kbMessage('Data Analisa Cabang belum tersedia. Unggah file Analisa melalui menu Import Data agar perkembangan cabang dapat ditampilkan.')
      : h('div', null,
        h('div', { class: 'kb-grid' }, [...KB_DEV_KEYS, ...KB_ACHV_KEYS].map(key => {
          const mt = m.metricOf(key), money = mt.kind === 'money';
          return kbChangeCard(mt.label, m.anValue(key), m.anDelta(key, true), money, money ? rp(m.an.current[key] || 0) : null, KB_METRIC_ICON[key], KB_DEV_KEYS.includes(key) ? 'Dev.' : 'Achv.',
            m.pola === 'B' && KB_ACHV_KEYS.includes(key) ? m.anAchv(key) : null);
        })),
        h('p', { class: 'kb-foot', text: kbLegendNote(m.pola) }));

  // 2. Kondisi Permasalahan Cabang
  const rbBlock = !L.rekening ? kbWait('Rekening Bergerak', 3)
    : !m.rbReady ? kbMessage('Rekening Bergerak belum dihitung: Tanggal N belum ditetapkan oleh Administrator atau User 0.')
      : h('div', null,
        h('div', { class: 'kb-grid' },
          kbMetricCard('Total PAR', fmt(m.rbTotal), 'Total Rekening PAR sampai ' + m.rbAsOfLabel, '', 'alert'),
          kbMetricCard('Rekening Bergerak', fmt(m.rbActive), 'Total Rekening Transaksi sampai ' + m.rbAsOfLabel, 'good', 'pulse'),
          kbMetricCard('Persentase Bergerak', nf1.format(m.rbPct) + '%', 'Dibanding total PAR', 'good', 'gauge')),
        h('div', { class: 'kb-meter', role: 'img', 'aria-label': 'Persentase rekening bergerak ' + nf1.format(m.rbPct) + ' persen' },
          h('i', { style: 'width:' + Math.min(100, Math.max(0, m.rbPct)).toFixed(1) + '%' })));
  const anBlock = !L.analisa ? kbWait('Analisa Cabang', 2, true)
    : !m.anReady ? kbMessage('Data Analisa Cabang belum tersedia.')
      : h('div', { class: 'kb-grid two' },
        kbChangeCard('Member PAR', m.anValue('memberPAR'), m.anDelta('memberPAR', false), false, null, 'users', 'Dev.'),
        kbChangeCard('Outstanding PAR', m.anValue('outstandingPAR'), m.anDelta('outstandingPAR', false), true, rp(m.an.current.outstandingPAR || 0), 'alert', 'Dev.'));
  const s2 = h('div', { class: 'kb-stack' },
    h('h4', { class: 'kb-subhead', text: m.rbReady ? 'Perkembangan Rekening Bergerak sampai ' + m.rbAsOfLabel + ', Data dihitung dari tanggal ' + m.rbSinceLabel + '.' : 'Rekening Bergerak' }), rbBlock,
    h('h4', { class: 'kb-subhead', text: 'Analisa Cabang · PAR' + (m.anReady ? '  ·  ' + m.anLabel : '') }), anBlock);

  // 3. Rekomendasi Penyelesaian
  const s3 = !L.tool ? kbWait('Tool PAR', 3)
    : h('div', null,
      h('div', { class: 'kb-grid' },
        kbCategoryCard('Angsuran By Sukarela', m.rowsA.length, 'Nominal angsuran', rpShort(m.angNominal), fmt(m.angKali) + ' kali angsuran', 'good', 'calendar'),
        kbCategoryCard('Pelunasan di Bawah 100.000', m.rows1.length, 'Total Balance', rpShort(m.sum(m.rows1, r => r[F.BALANCE])), 'Sisa pelunasan sampai Rp 100.000', 'warn', 'check'),
        kbCategoryCard('Pelunasan dari Simpanan', m.rowsS.length, 'Total Balance', rpShort(m.sum(m.rowsS, r => r[F.BALANCE])), 'Simpanan sudah menutup Total Balance', 'good', 'lock')),
      h('p', { class: 'kb-foot', text: 'Nominal angsuran = jumlah angsuran × Cicilan (Pokok + Margin). Satu pinjaman dapat masuk lebih dari satu kategori.' }));

  // 4. Kondisi Sampai Hari Kemarin
  const s4 = !L.progress ? (state.kb.errors['Progres Harian'] ? kbWait('Progres Harian', 3) : h('div', { class: 'kb-stack' }, kbSkeleton(3), h('div', { class: 'kb-charts' }, h('div', { class: 'kb-card skel kb-skel tall' }), h('div', { class: 'kb-card skel kb-skel tall' }))))
    : h('div', { class: 'kb-stack' },
      h('div', { class: 'kb-grid' },
        kbMetricCard('PAR Baru · Peningkatan', rpShort(m.newBalance), fmt(m.newCount) + ' Member PAR', 'bad', 'plus'),
        kbMetricCard('PAR Selesai · Penyelesaian', rpShort(m.doneBalance), fmt(m.doneCount) + ' Member PAR', 'good', 'check'),
        kbMetricCard('Penurunan Outstanding PAR By Angsuran', rpShort(m.red.amount || 0), fmt(m.red.loans || 0) + ' Loan No turun',
          Number(m.red.amount) > 0 ? 'good' : Number(m.red.amount) < 0 ? 'warn' : '', 'down')),
      h('div', { class: 'kb-charts' },
        kbLineFigure('Peningkatan vs Penyelesaian Out. PAR' + (m.recent.length ? ' · ' + m.recent.length + ' tanggal terakhir' : ''), 'kb-chart-out'),
        kbLineFigure('Peningkatan vs Penyelesaian Member PAR' + (m.recent.length ? ' · ' + m.recent.length + ' tanggal terakhir' : ''), 'kb-chart-client')));

  // 5. Top Up Khusus yang Kembali PAR
  const s5 = !L.progress ? kbWait('Progres Harian', 2, true)
    : h('div', null,
      h('div', { class: 'kb-grid two' },
        kbMetricCard('Total PAR Baru Top Up Khusus', fmt(m.khCount) + ' Loan No', (m.newCount ? nf1.format(m.khCount / m.newCount * 100) : '0') + '% dari seluruh PAR Baru', 'bad', 'star'),
        kbMetricCard('Outstanding PAR', rpShort(m.khBalance), (m.newBalance ? nf1.format(m.khBalance / m.newBalance * 100) : '0') + '% dari Out. PAR Baru', 'bad', 'coins')),
      h('p', { class: 'kb-foot', text: 'Top Up Khusus = pinjaman kategori TOP UP KHUSUS yang tercatat sebagai PAR Baru.' }));

  put(root, masthead, lead,
    kbSection(1, 'Perkembangan Cabang', 'Analisa Cabang · ' + (L.analisa ? m.anLabel : 'memuat data…'), '#0a7a5a', s1, [['Analisa Cabang', 'analisa']]),
    kbSection(2, 'Kondisi Permasalahan Cabang', 'Rekening Bergerak dan Analisa PAR', '#b43d3d', s2, [['Rekening Bergerak', 'move'], ['Analisa Cabang', 'analisa']]),
    kbSection(3, 'Rekomendasi Penyelesaian', 'Tool PAR · ' + (L.tool ? (m.branch ? 'Cabang ' + m.branch : fmt(Number(m.d.scanned) || 0) + ' pinjaman diperiksa') : 'memuat data…'), '#d0a12b', s3, [['Tool PAR', 'dash']]),
    kbSection(4, 'Kondisi Sampai Hari Kemarin', 'Progres Harian · ' + (L.progress ? 'data sampai ' + (m.lastDate ? analisaLabelDate(m.lastDate) : '—') + ' · seluruh periode' : 'memuat data…'), '#2b6cb0', s4, [['Progres Harian', 'progress']]),
    kbSection(5, 'Top Up Khusus yang Kembali PAR', 'Progres Harian · PAR Baru kategori Top Up Khusus · seluruh periode', '#b43d3d', s5, [['Progres Harian', 'progress']]),
    h('p', { class: 'kb-foot kb-scope', text: 'Cakupan: ' + mainScope + (m.ph.scope ? ' · Progres Harian: ' + m.ph.scope : '') + '. Ringkasan informasi; rincian dan daftar pinjaman ada di Dashboard masing-masing.' }));
  if (L.progress) {   // grafik garis dengan penanda: fungsi yang sama dengan halaman Progres Harian
    progresDayChart('kb-chart-out', m.recent, 'newBalance', 'doneBalance', true, '');
    progresDayChart('kb-chart-client', m.recent, 'newCount', 'doneCount', false, '');
  }
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
    if (!regionalNewsAllReady()) throw new Error('Data belum lengkap. Periksa keterangan pemuatan di halaman Kabar Berita.');
    await nextFrame();
    const canvas = imgBuildRegionalNews();
    const scopeName = state.kb.branch ? state.kb.branch.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '') : 'Regional_J';
    await imgDownload(canvas, 'Kabar_Berita_' + scopeName + '_' + today() + '.png', 0);
    setNote('sum-news-msg', 'Gambar "Kabar Berita" berhasil diunduh dan siap dibagikan.', 'good');
  } catch (error) {
    setNote('sum-news-msg', error && error.message ? error.message : 'Gagal membuat gambar Kabar Berita.', 'bad');
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
    noteDataVersion(r.dataVersion, 'move');
    state.rekening = r;
    state.rekeningLoadedAt = Date.now();
    paintRekeningBergerak(r);
    $('rb-sync').textContent = 'Terakhir diperbarui: ' + new Date(state.rekeningLoadedAt).toLocaleString('id-ID')
      + (r.dataUpdated ? ' · Data per ' + fmtStamp(r.dataUpdated) : '');
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
    if (state.dashboardPage === 'sum') prepareRegionalNewsData();
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
  $('ph-sync').textContent = 'Terakhir diperiksa: ' + (data.generatedAt || '').replace('T', ' ') + (data.timezone ? ' · Zona waktu ' + data.timezone : '')
    + (data.dataUpdated ? ' · Data per ' + fmtStamp(data.dataUpdated) : '');
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
    noteDataVersion(r.dataVersion, 'progress');
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
   Setting: Pola Hitung (ADMINISTRATOR dan USER 0; server memeriksa ulang)
   A = Mencari Jarak (default), B = Mencari Capaian. Berlaku untuk Disbursement, Repayment, dan Margin di Kabar Berita.
   --------------------------------------------------------------------- */
const POLA_LABEL = { A: 'A. Mencari Jarak', B: 'B. Mencari Capaian' };
const polaSelected = () => { const el = document.querySelector('#pola-list input:checked'); return el ? el.value : ''; };
function syncPolaForm() {   // tombol Simpan aktif hanya bila pilihan berbeda dari yang tersimpan
  $('pola-save').disabled = !state.pola || !polaSelected() || polaSelected() === state.pola;
}
function paintPola() {
  const cur = state.pola;
  $('pola-cur').textContent = cur ? POLA_LABEL[cur] + (cur === 'A' ? ' (Default)' : '') : 'Memuat…';
  if (cur && !state.polaDirty) document.querySelectorAll('#pola-list input').forEach(i => { i.checked = i.value === cur; });   // jangan menimpa pilihan yang sedang diubah
  syncPolaForm();
}
async function loadPola() {
  if (state.polaLoading || !state.user) return;
  state.polaLoading = true;
  try {
    const r = await api('webGetPolaHitung', { token: (session.get() || {}).token }, { retry: true });
    if (r.code === 'SESSION_REQUIRED') return logout(r.error);
    if (r.ok) { state.pola = r.pola === 'B' ? 'B' : 'A'; paintPola(); }
  } finally { state.polaLoading = false; }
}
async function savePola() {
  const value = polaSelected();
  if (!value || value === state.pola) return setNote('pola-msg', 'Tidak ada perubahan.');
  const ok = await askConfirm('Simpan Pola Hitung?', (value === 'B'
    ? 'Disbursement, Repayment dan Margin pada Kabar Berita akan menampilkan persentase capaian terhadap akhir bulan sebelumnya.'
    : 'Disbursement, Repayment dan Margin pada Kabar Berita akan menampilkan selisih terhadap akhir bulan sebelumnya.') + ' Pengaturan ini berlaku untuk semua pengguna.', 'Simpan');
  if (!ok) return;
  $('pola-save').disabled = true; setNote('pola-msg', 'Menyimpan…');
  const r = await api('webSetPolaHitung', { token: (session.get() || {}).token, value });
  if (r.code === 'SESSION_REQUIRED') return logout(r.error);
  if (!r.ok) { syncPolaForm(); return setNote('pola-msg', r.error || 'Gagal menyimpan.', 'bad'); }
  state.pola = r.pola === 'B' ? 'B' : 'A'; state.polaDirty = false;
  if (state.kb.analisa) state.kb.analisa.r.polaHitung = state.pola;   // langsung dipakai Kabar Berita (layar dan gambar unduhan)
  paintPola(); paintKabarBerita();
  setNote('pola-msg', 'Tersimpan: Pola Hitung ' + POLA_LABEL[state.pola] + '.', 'good');
}

/* ---------------------------------------------------------------------
   Manajemen menu: tampil / sembunyi menu dan sub menu per tingkat (menu "Manajemen" hanya untuk ADMINISTRATOR)
   Peta menu tiap akun datang dari server (webMenuGet, Manajemen.gs). Salinannya disimpan di perangkat supaya sidebar langsung benar
   saat aplikasi dibuka, lalu disegarkan dari server. Bila Manajemen.gs belum dipasang, aturan lama dipakai (tidak ada yang berubah).
   Ini pengaturan TAMPILAN; pembatasan data tetap dijalankan server per tingkat dan Wilayah Kerja.
   --------------------------------------------------------------------- */
const MENU_CACHE = 'dash_menu_v1';
const MENU_REFRESH_MS = 60 * 1000;
const DASH_PAGES = ['sum', 'progress', 'analisa', 'move', 'dash'];      // urutan tombol sub menu Dashboard
const SET_PAGES = ['global', 'wilayah', 'enable-download', 'write-off', 'manajemen'];             // urutan tombol sub menu Setting
const MSG_OLD_MENU = 'Server belum mendukung menu Manajemen. Tambahkan file Manajemen.gs dan perbarui Akses.gs di Apps Script, lalu Deploy > Manage deployments > Edit > Version: New version > Deploy.';

// Aturan sebelum Manajemen ada: Setting untuk ADMINISTRATOR dan USER 0; Import Data, Enable Download, Write-Off (Setting), dan Manajemen (sub menu Setting) khusus ADMINISTRATOR.
function menuLegacy(key, u) {
  if (key === 'set.manajemen' || key === 'up' || key === 'set.enable-download' || key === 'set.write-off') return u.tingkat === 'ADMINISTRATOR';
  if (key === 'set' || key.indexOf('set.') === 0) return canSetting(u);
  return true;
}
// Apakah menu/sub menu "key" tampil untuk akun yang sedang masuk. ADMINISTRATOR selalu melihat semuanya; Manajemen hanya ADMINISTRATOR.
function menuOn(key) {
  const u = state.user;
  if (!u) return false;
  if (u.tingkat === 'ADMINISTRATOR') return true;
  if (key === 'set.manajemen') return false;
  const m = state.menuAccess, has = k => !!m && Object.prototype.hasOwnProperty.call(m, k);
  if (has(key)) return !!m[key];
  // Peta menu dari server lama (sebelum Write-Off digabung): Eksekusi/Approve/Report masih berdiri sendiri -> pakai pengaturannya
  const old = { 'wo.ewo': 'ewo', 'wo.awo': 'awo', 'wo.rwo': 'rwo' }[key];
  if (old && has(old)) return !!m[old];
  if (key === 'wo' && (has('ewo') || has('awo') || has('rwo'))) return WO_PAGES.some(k => menuOn('wo.' + k));
  return menuLegacy(key, u);
}
function menuFirstTab() { return TABS.find(k => menuOn(k)) || 'versions'; }
function menuFirstDash() { return DASH_PAGES.find(p => menuOn('dashboard.' + p)) || ''; }
function menuFirstSet() { return SET_PAGES.find(p => menuOn('set.' + p)) || 'global'; }

function menuCached(nik) {
  try {
    const o = JSON.parse(localStorage.getItem(MENU_CACHE) || 'null');
    return o && o.nik === nik && o.menus && typeof o.menus === 'object' ? o.menus : null;
  } catch (e) { return null; }
}
function menuRemember(nik, menus) { try { localStorage.setItem(MENU_CACHE, JSON.stringify({ nik: nik, menus: menus })); } catch (e) { /* abaikan */ } }
function menuForget() { try { localStorage.removeItem(MENU_CACHE); } catch (e) { /* abaikan */ } }

// Pasang tampil/sembunyi pada tombol menu & sub menu. redirect = true: bila halaman yang sedang dibuka kini tersembunyi, pindah ke yang tampil.
function menuApply(redirect) {
  if (!state.user) return;
  TABS.forEach(k => { const t = $('t-' + k); if (t) t.hidden = !menuOn(k); });
  document.querySelectorAll('[data-dashpage]').forEach(b => { b.hidden = !menuOn('dashboard.' + b.dataset.dashpage); });
  document.querySelectorAll('[data-setpage]').forEach(b => { b.hidden = !menuOn('set.' + b.dataset.setpage); });
  document.querySelectorAll('[data-wopage]').forEach(b => { b.hidden = !menuOn('wo.' + b.dataset.wopage); });
  const nav = document.querySelector('#app>.tabs');
  if (nav) {
    const count = TABS.filter(k => !$('t-' + k).hidden).length;
    nav.style.setProperty('--nav-count', String(count));
    nav.classList.toggle('is-dense', count >= 8);   // 8 tombol (ADMINISTRATOR): huruf dan jarak bilah bawah di layar HP dirapatkan
  }
  navPaint();
  if (!redirect) return;
  const cur = TABS.find(k => !$('p-' + k).hidden);
  if (cur && !menuOn(cur)) { selectTab(menuFirstTab()); return; }
  if (cur === 'dashboard' && !menuOn('dashboard.' + state.dashboardPage)) selectDashboardPage(menuFirstDash() || 'move');
  if (cur === 'set' && !menuOn('set.' + state.settingPage)) setSettingPage(menuFirstSet());
  if (cur === 'wo' && !menuOn('wo.' + state.woPage)) selectWoPage(menuFirstWo());
}

// Ambil peta menu akun dari server. ADMINISTRATOR tidak perlu (selalu semua menu). Gagal / server lama: aturan terakhir tetap dipakai.
async function menuRefresh(force) {
  const u = state.user;
  if (!u || u.tingkat === 'ADMINISTRATOR' || state.menuLoading) return;
  if (!force && Date.now() - state.menuAt < MENU_REFRESH_MS) return;
  state.menuLoading = true;
  const r = await api('webMenuGet', { token: (session.get() || {}).token }, { retry: true });
  state.menuLoading = false;
  if (!state.user || state.user.nik !== u.nik) return;
  if (r.code === 'SESSION_REQUIRED') return logout(r.error);
  if (!r.ok || !r.menus || typeof r.menus !== 'object') return;
  state.menuAt = Date.now();
  const same = JSON.stringify(r.menus) === JSON.stringify(state.menuAccess);
  state.menuAccess = r.menus;
  menuRemember(u.nik, r.menus);
  if (!same) menuApply(true);
}

/* ---------- Halaman Manajemen (ADMINISTRATOR) ---------- */
const mg = { data: null, draft: null, loading: false, saving: false, at: 0 };
const mgClone = o => JSON.parse(JSON.stringify(o));

function mgReset() {
  mg.data = null; mg.draft = null; mg.loading = false; mg.saving = false; mg.at = 0;
  const wrap = $('mg-wrap'); if (wrap) wrap.replaceChildren();
  setNote('mg-msg', ''); const last = $('mg-last'); if (last) last.textContent = '';
  mgPaintButtons();
}
// Yang benar-benar tampil untuk satu tingkat (aturan sama dengan server): induk tersembunyi -> sub menu ikut; induk tanpa sub menu tampil -> induk tersembunyi.
function mgEffective(matrix, cat, level) {
  const vis = {};
  cat.forEach(c => { vis[c.key] = !!(matrix[c.key] && matrix[c.key][level]); });
  cat.filter(c => !c.parent).forEach(p => {
    const kids = cat.filter(c => c.parent === p.key);
    if (!kids.length) return;
    if (!vis[p.key]) kids.forEach(k => { vis[k.key] = false; });
    else if (!kids.some(k => vis[k.key])) vis[p.key] = false;
  });
  return vis;
}
function mgChanges() {
  if (!mg.data || !mg.draft) return 0;
  let n = 0;
  mg.data.catalog.forEach(c => mg.data.levels.forEach(l => { if (!!mg.draft[c.key][l] !== !!mg.data.matrix[c.key][l]) n++; }));
  return n;
}
const mgDirty = () => mgChanges() > 0;
function mgPaintButtons() {
  const busy = mg.loading || mg.saving, has = !!mg.data;
  $('mg-save').disabled = busy || !has || !mgDirty();
  $('mg-undo').disabled = busy || !has || !mgDirty();
  $('mg-default').disabled = busy || !has;
  $('mg-refresh').disabled = busy;
}
function mgEnter() {
  if (!state.user || state.user.tingkat !== 'ADMINISTRATOR') return;
  if (!mg.data || (!mgDirty() && Date.now() - mg.at > 60000)) mgLoad();
}
async function mgLoad() {
  if (mg.loading) return;
  mg.loading = true; mgPaintButtons();
  setNote('mg-msg', 'Memuat pengaturan menu…');
  const r = await api('webMenuGet', { token: (session.get() || {}).token, full: '1' }, { retry: true });
  mg.loading = false;
  if (r.code === 'SESSION_REQUIRED') return logout(r.error);
  if (!r.ok || !Array.isArray(r.catalog) || !r.matrix || !Array.isArray(r.levels)) {
    const old = !r.ok && /Aksi tidak dikenal|belum ditambahkan/.test(r.error || '');
    mgPaintButtons();
    return setNote('mg-msg', r.ok ? MSG_OLD_MENU : old ? MSG_OLD_MENU : (r.error || 'Gagal memuat pengaturan menu.'), 'bad');
  }
  mg.data = r; mg.draft = mgClone(r.matrix); mg.at = Date.now();
  setNote('mg-msg', '');
  mgRender();
}
function mgRefresh() {
  if (mg.loading || mg.saving) return;
  if (!mgDirty()) return mgLoad();
  askConfirm('Muat ulang dan buang perubahan?', 'Perubahan yang belum disimpan akan hilang.', 'Muat ulang').then(ok => { if (ok) mgLoad(); });
}
function mgLockTitle(c) { return 'Dikunci sistem: hanya ' + ['Administrator'].concat(c.max).join(', ') + ' yang boleh memakai menu ini.'; }

function mgRender(focus) {
  const d = mg.data; if (!d) return;
  const cat = d.catalog, levels = d.levels, byKey = {};
  cat.forEach(c => { byKey[c.key] = c; });
  const eff = {}; levels.forEach(l => { eff[l] = mgEffective(mg.draft, cat, l); });
  const allOn = l => cat.filter(c => c.max.indexOf(l) > -1).every(c => mg.draft[c.key][l]);
  const head = h('tr', null,
    h('th', { scope: 'col', class: 'mg-menu', text: 'Menu / Sub menu' }),
    h('th', { scope: 'col', class: 'mg-lv' }, h('span', { class: 'mg-lvname', text: 'ADMINISTRATOR' }), h('span', { class: 'mg-fixed', text: 'selalu tampil' })),
    levels.map(l => h('th', { scope: 'col', class: 'mg-lv' }, h('span', { class: 'mg-lvname', text: l }),
      h('button', { type: 'button', class: 'mg-all', 'data-all': l, text: allOn(l) ? 'Kosongkan' : 'Pilih semua' }))));
  const cell = (c, l) => {
    if (c.max.indexOf(l) < 0) return h('td', { class: 'mg-lv' }, h('span', { class: 'mg-lock', title: mgLockTitle(c), text: 'Dikunci' }));
    const on = !!mg.draft[c.key][l], ghost = on && !eff[l][c.key];
    const label = (c.parent ? byKey[c.parent].label + ' › ' : '') + c.label;
    return h('td', { class: 'mg-lv' }, h('label', { class: 'mg-sw' + (ghost ? ' is-ghost' : '') + (on ? ' is-on' : ''), title: ghost ? 'Dicentang, tetapi tidak tampil karena menu induk atau seluruh sub menunya disembunyikan.' : null },
      h('input', { type: 'checkbox', 'data-k': c.key, 'data-lv': l, checked: on ? true : null, 'aria-label': (on ? 'Sembunyikan ' : 'Tampilkan ') + label + ' untuk ' + l }),
      h('span', { class: 'mg-knob', 'aria-hidden': 'true' }),
      h('span', { class: 'mg-state', text: !on ? 'Sembunyi' : ghost ? 'Ikut tersembunyi' : 'Tampil' })));
  };
  const rows = cat.map(c => h('tr', { class: 'mg-row ' + (c.parent ? 'is-sub' : 'is-main') },
    h('th', { scope: 'row', class: 'mg-menu' }, h('span', { class: 'mg-name', text: c.label })),
    h('td', { class: 'mg-lv' }, h('span', { class: 'mg-always', text: 'Tampil' })),
    levels.map(l => cell(c, l))));
  $('mg-wrap').replaceChildren(h('table', { class: 'mg-tbl' }, h('thead', null, head), h('tbody', null, rows)));
  // keterangan perubahan terakhir
  let last = null;
  Object.keys(d.meta || {}).forEach(k => { const m = d.meta[k]; if (m && m.at && (!last || m.at > last.at)) last = m; });
  $('mg-last').textContent = last ? 'Terakhir diubah oleh ' + (last.by || '-') + ' · ' + fmtStamp(last.at) + '. Disimpan di sheet "' + (d.tab || 'Menu_Akses') + '".'
    : 'Belum pernah diubah: seluruh menu tampil seperti aturan bawaan. Pengaturan akan disimpan di sheet "' + (d.tab || 'Menu_Akses') + '".';
  const n = mgChanges();
  if (!mg.saving) setNote('mg-msg', n ? n + ' pengaturan berubah dan belum disimpan.' : '');
  mgPaintButtons();
  if (focus) {
    const el = focus.k ? $('mg-wrap').querySelector('input[data-k="' + focus.k + '"][data-lv="' + focus.l + '"]') : $('mg-wrap').querySelector('button[data-all="' + focus.l + '"]');
    if (el) el.focus();
  }
}
function mgChange(e) {
  const t = e.target;
  if (!mg.draft || !t || t.type !== 'checkbox' || !t.dataset.k) return;
  mg.draft[t.dataset.k][t.dataset.lv] = t.checked;
  mgRender({ k: t.dataset.k, l: t.dataset.lv });
}
function mgClick(e) {
  const b = e.target.closest('button[data-all]');
  if (!b || !mg.draft) return;
  const l = b.dataset.all, cat = mg.data.catalog, usable = cat.filter(c => c.max.indexOf(l) > -1);
  const turnOn = !usable.every(c => mg.draft[c.key][l]);
  usable.forEach(c => { mg.draft[c.key][l] = turnOn; });
  mgRender({ l: l });
}
function mgUndo() {
  if (!mg.data) return;
  mg.draft = mgClone(mg.data.matrix);
  mgRender();
}
function mgDefault() {   // bawaan = semua menu yang boleh dipakai tingkat itu tampil (sama seperti sebelum Manajemen ada)
  if (!mg.data) return;
  mg.data.catalog.forEach(c => mg.data.levels.forEach(l => { mg.draft[c.key][l] = c.max.indexOf(l) > -1; }));
  mgRender();
  if (mgDirty()) setNote('mg-msg', 'Pengaturan dikembalikan ke bawaan (' + mgChanges() + ' perubahan). Tekan Simpan untuk menerapkan.');
}
async function mgSave() {
  if (!mg.data || mg.saving || !mgDirty()) return;
  const cat = mg.data.catalog;
  for (const l of mg.data.levels) {
    const e = mgEffective(mg.draft, cat, l);
    if (!cat.some(c => !c.parent && e[c.key])) return setNote('mg-msg', 'Tingkat ' + l + ' harus memiliki minimal satu menu utama yang tampil.', 'bad');
  }
  const ok = await askConfirm('Simpan pengaturan menu?',
    mgChanges() + ' pengaturan berubah. Berlaku untuk semua akun pada tingkat terkait: akun yang sedang aktif menerimanya dalam beberapa menit, akun lain saat berikutnya membuka aplikasi.', 'Simpan');
  if (!ok) return;
  mg.saving = true; mgPaintButtons();
  setNote('mg-msg', 'Menyimpan…');
  const r = await api('webMenuSave', { token: (session.get() || {}).token, matrix: JSON.stringify(mg.draft) });   // aksi tulis: tidak diulang otomatis
  mg.saving = false;
  if (r.code === 'SESSION_REQUIRED') return logout(r.error);
  if (!r.ok) { mgPaintButtons(); return setNote('mg-msg', r.error || 'Pengaturan menu gagal disimpan.', 'bad'); }
  mg.data = r; mg.draft = mgClone(r.matrix); mg.at = Date.now();
  mgRender();
  setNote('mg-msg', 'Tersimpan. Perubahan berlaku untuk akun terkait.', 'good');
  toast('Pengaturan menu berhasil disimpan.');
}

/* ---------------------------------------------------------------------
   Setting: Wilayah Kerja (ADMINISTRATOR dan USER 0; akses server diperiksa ulang)
   --------------------------------------------------------------------- */
function setSettingPage(which) {
  if (SET_PAGES.indexOf(which) < 0) which = 'global';
  if (state.user && !menuOn('set.' + which)) which = menuFirstSet();   // sub menu yang disembunyikan / tidak boleh untuk tingkat ini
  state.settingPage = which;
  document.querySelectorAll('[data-setpage]').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.setpage === which)));
  $('set-global').hidden = which !== 'global';
  $('set-wilayah').hidden = which !== 'wilayah';
  $('set-enable-download').hidden = which !== 'enable-download';
  $('set-write-off').hidden = which !== 'write-off';
  $('set-manajemen').hidden = which !== 'manajemen';
  if (which === 'wilayah' && (!state.wilayah || (!wilayahIsDirty() && Date.now() - state.wilayahLoadedAt > 60000))) loadWilayah();
  if (which === 'enable-download') { paintImageDownloadSettings(); setNote('image-download-msg', ''); }
  if (which === 'write-off') woSettingPaint();
  if (which === 'manajemen') mgEnter();
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
  resetKabarBerita();   // cakupan cabang berubah: buang data berita lama (termasuk pilihan filter)
  state.rekening = null; state.rekeningLoadedAt = 0;
  state.progress = null; state.progressLoadedAt = 0;
  state.analisaCabang = null; state.analisaLoadedAt = 0;
  setNote('wilayah-msg', 'Tersimpan. ' + fmt(r.assignments) + ' penugasan cabang untuk ' + fmt(r.users) + ' user.', 'good');
  await loadWilayah();
  loadDashFresh();
}

/* ---------------------------------------------------------------------
   11C. Write-Off: unggah data master, Eksekusi Write-Off, Approve Write-Off
   Server: WriteOff.gs (aksi webWo*). Hak akses, status ID_EWO, dan semua perhitungan DITEGAKKAN DI SERVER;
   bagian ini hanya membaca file, menampilkan data, dan mengirim isian pengguna.
   --------------------------------------------------------------------- */
const WO_ST_CONFIRM = 'Confirm', WO_ST_APPROVE = 'Approve Write-Off', WO_BATAL = 'Batal', WO_VER_OK = 'Sesuai', WO_VER_NO = 'Tidak';
const WO_SHEET = { R: 'WO Reguler', K: 'WO Khusus' };
const WO_CHUNK = 2000;        // baris per permintaan unggah master (batas server 2.500)
const WO_PAGE = 50;           // baris per halaman di area kerja Eksekusi
const WO_REF_MAX = 100;       // panjang maksimum No. THC / No. TAK / No JV (sama dengan server)
const WO_HEADER_ROW = 3;      // header file master ada di baris Excel 3; data mulai baris 4
const WO_EWO_RE = /EWO-[RK]-\d{4}-\d{9}/i;
const WO_ALASAN_MAX = 200;    // panjang maksimum Alasan Batal (sama dengan server)
const WO_BATAL_OPTIONS = ['Anggota sudah non-aktif', 'Masuk ke list WO lainnya', 'Anggota memiliki lebih dari 1 pinjaman'];   // pilihan ke-4 = Lainnya (teks bebas)

// [judul di sheet, tipe (t teks, n angka, d periode), nama yang dikenali pada header Excel (huruf kecil tanpa tanda baca)].
// 19 kolom pertama dibaca dari Excel; sisanya diisi di sistem. Urutan WAJIB sama dengan WO_COLS di WriteOff.gs.
const WO_COLS = [
  ['No.', 'n', ['no']], ['Branch Name', 't', ['branchname']], ['Loan No', 't', ['loanno']], ['Ctr ID', 't', ['ctrid']], ['Client ID', 't', ['clientid']],
  ['Client Name', 't', ['clientname']], ['Product Name', 't', ['productname']], ['Periode Disbursed', 'd', ['disburseddate', 'periodedisbursed']],
  ['Total Balance', 'n', ['totalbalance']], ['Simp. Wajib', 'n', ['simpwajib']], ['Simp. Sukarela', 'n', ['simpsukarela']], ['Simp. Pensiun', 'n', ['simppensiun']],
  ['Simp. Hari Raya', 'n', ['simphariraya']], ['Simp. Khusus', 'n', ['simpkhusus']], ['Simp. Qurban', 'n', ['simpqurban']], ['Simp. Sipadan', 'n', ['simpsipadan']],
  ['Total Simpanan', 'n', ['totalsimpanan']], ['PIUTANG YANG DI WO', 'n', ['piutangyangdiwo']], ['Alasan diajukan WO', 't', ['alasandiajukanwo']],
  ['Tahun', 'n'], ['No. THC', 't'], ['No. TAK', 't'], ['No JV', 't'], ['Verifikasi Nilai WO', 't'], ['ID_EWO', 't'], ['Status Baris', 't'], ['Alasan Batal', 't'],
];
const WC = { NO: 0, BRANCH: 1, LOAN: 2, CTR: 3, CLIENT: 4, NAME: 5, PRODUCT: 6, PERIODE: 7, TB: 8, WAJIB: 9, SUKARELA: 10, PENSIUN: 11, HARIRAYA: 12, KHUSUS: 13,
  QURBAN: 14, SIPADAN: 15, TOTSIMP: 16, PIUTANG: 17, ALASAN: 18, TAHUN: 19, THC: 20, TAK: 21, JV: 22, VERIF: 23, EWO: 24, STATUS: 25, ALASAN_BATAL: 26 };
const WO_NIMPORT = 19, WO_NBASE = 24;   // kolom 1-19 dari Excel; 1-24 = struktur dokumen; 25-27 (ID_EWO, Status Baris, Alasan Batal) hanya sistem
// Nominal yang bisa diubah lewat Manual: [nama parameter server, kolom]
const WO_NOMINAL = [['tb', WC.TB], ['wajib', WC.WAJIB], ['sukarela', WC.SUKARELA], ['pensiun', WC.PENSIUN], ['hariraya', WC.HARIRAYA],
  ['khusus', WC.KHUSUS], ['qurban', WC.QURBAN], ['sipadan', WC.SIPADAN]];

const woCompact = v => String(v == null ? '' : v).toLowerCase().replace(/[^a-z0-9]/g, '');
const woTok = () => (session.get() || {}).token;
const woNum = x => Math.round(x * 100) / 100;
const woPlain = v => v === '' || v == null ? '—' : String(v);

const wo = {
  epoch: 0,                                                   // naik setiap Keluar/Masuk: balasan server yang terlambat diabaikan
  opt: null, ewos: [], optAt: 0, optLoading: false, pick: { b: '', j: '', y: '' }, busy: false,
  cur: null, rows: [], curAt: 0, q: '', page: 0,
  awo: null, awoAt: 0, awoLoading: false, awoJ: '', rep: null,
  manual: null, auto: null, jvId: '', verifId: '', reopenId: '', det: null, batal: null, lastBackup: null, setBusy: false,
};

/* ---------- Unggah data master (ADMINISTRATOR, menu Import Data) ---------- */

// Murni: array baris Excel (indeks 0 = baris 1) -> data master. Dipisah dari pembacaan file agar bisa diuji tanpa pustaka Excel.
function parseWoMaster(aoa, info) {
  const hr = WO_HEADER_ROW - 1, head = aoa[hr] || [];
  const cols = WO_COLS.slice(0, WO_NIMPORT).map(c => head.findIndex(v => c[2].indexOf(woCompact(v)) > -1));
  const missing = cols.map((c, k) => c < 0 ? WO_COLS[k][0] : '').filter(Boolean);
  if (missing.length) {
    throw new Error('Header tabel harus berada di baris Excel ' + WO_HEADER_ROW + ' dan memuat semua kolom. Belum ditemukan: ' + missing.join(', ') + '.');
  }
  // Data = baris 4 sampai baris yang nilai "No."-nya paling besar.
  let maxNo = -Infinity, last = -1;
  for (let i = hr + 1; i < aoa.length; i++) {
    const v = toNum((aoa[i] || [])[cols[WC.NO]]);
    if (typeof v === 'number' && v > maxNo) { maxNo = v; last = i; }
  }
  if (last < 0) throw new Error('Kolom "No." tidak berisi angka mulai baris Excel ' + (WO_HEADER_ROW + 1) + '. Batas akhir data tidak dapat ditentukan.');
  const rows = [], seen = new Map(), branches = new Set();
  let skipped = 0, sumTb = 0, sumPiutang = 0;
  for (let i = hr + 1; i <= last; i++) {
    const r = aoa[i] || [];
    const raw = cols.map(c => r[c]);
    if (raw.every(blank)) { skipped++; continue; }
    const loan = toTxt(r[cols[WC.LOAN]]);
    if (!loan) { skipped++; continue; }                       // tanpa Loan No = bukan baris data
    const row = WO_COLS.slice(0, WO_NIMPORT).map((c, k) => {
      const v = raw[k];
      if (c[1] === 'n') {
        if (blank(v)) return '';
        const x = toNum(typeof v === 'string' ? v.replace(/^Rp\.?/i, '').trim() : v);
        if (x === '') throw new Error('Nilai angka tidak valid pada baris Excel ' + (i + 1) + ', kolom "' + c[0] + '": ' + String(v).slice(0, 30) + '.');
        return x;
      }
      if (c[1] === 'd') return toDate(v) || toTxt(v);
      return toTxt(v);
    });
    if (!row[WC.BRANCH]) throw new Error('Branch Name kosong pada baris Excel ' + (i + 1) + ' (Loan No ' + loan + ').');
    if (seen.has(loan)) throw new Error('Loan No ' + loan + ' muncul lebih dari sekali (baris Excel ' + seen.get(loan) + ' dan ' + (i + 1) + ').');
    seen.set(loan, i + 1);
    branches.add(row[WC.BRANCH].toUpperCase());
    sumTb += Number(row[WC.TB]) || 0; sumPiutang += Number(row[WC.PIUTANG]) || 0;
    rows.push(row);
  }
  if (!rows.length) throw new Error('Tidak ada baris data (Loan No terisi) pada baris Excel ' + (WO_HEADER_ROW + 1) + ' sampai ' + (last + 1) + '.');
  return { type: 'WO', code: info.code, jenis: info.jenis, sheet: info.sheet, tahun: info.tahun, rows, skipped, lastRow: last + 1, maxNo,
    branchCount: branches.size, sumTb: woNum(sumTb), sumPiutang: woNum(sumPiutang), sent: 0 };
}

async function woWorkbookRows(file, sheetName, fallbackFirst) {
  if (file.size > UP_MAX_MB * 1048576) throw new Error('Ukuran file melebihi ' + UP_MAX_MB + ' MB.');
  const buf = await file.arrayBuffer(), b = new Uint8Array(buf.slice(0, 4));
  const zip = b[0] === 0x50 && b[1] === 0x4B, ole = b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0;
  if (!/\.(xlsx|xls)$/i.test(file.name) || !(zip || ole)) throw new Error('File harus berformat Excel (.xlsx atau .xls).');
  await loadXLSX();
  const wb = XLSX.read(buf, { type: 'array', dense: true });
  let name = wb.SheetNames.find(n => woCompact(n) === woCompact(sheetName));
  if (!name && fallbackFirst) name = wb.SheetNames[0];
  if (!name) throw new Error('Sheet "' + sheetName + '" tidak ditemukan di file. Sheet yang ada: ' + wb.SheetNames.join(', ') + '.');
  const ws = wb.Sheets[name];
  if (!ws || !ws['!ref']) throw new Error('Sheet "' + name + '" kosong.');
  // Selalu mulai dari A1: tanpa ini, bila baris-baris awal kosong SheetJS memulai dari baris pertama yang terisi dan nomor baris bergeser.
  const range = XLSX.utils.decode_range(ws['!ref']);
  range.s.r = 0; range.s.c = 0;
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true, range });
}

async function readWoExcel(file, info) {
  const d = parseWoMaster(await woWorkbookRows(file, info.sheet, false), info);
  d.name = file.name;
  return d;
}

function renderWoPreview(d) {
  const chip = (text, warn) => h('span', { class: 'st' + (warn ? ' warn' : ''), text });
  put($('upstats'), chip('DATA MASTER WRITE-OFF'), chip(d.name), chip('Sheet: ' + d.sheet + ' · Tahun ' + d.tahun), chip(fmt(d.rows.length) + ' baris data'),
    chip(fmt(d.branchCount) + ' cabang'), chip('Total Balance ' + rp(d.sumTb)), chip('PIUTANG YANG DI WO ' + rp(d.sumPiutang)),
    chip('Baris Excel ' + (WO_HEADER_ROW + 1) + '–' + d.lastRow + ' (No. terbesar ' + fmt(d.maxNo) + ')'),
    d.skipped ? chip(fmt(d.skipped) + ' baris kosong / tanpa Loan No dilewati', true) : null);
  const show = [WC.LOAN, WC.BRANCH, WC.NAME, WC.TB, WC.PIUTANG];
  const head = h('tr', null, show.map(c => h('th', { scope: 'col', class: WO_COLS[c][1] === 'n' ? 'num' : '', text: WO_COLS[c][0] })));
  const body = d.rows.slice(0, 5).map(r => h('tr', null, show.map(c => {
    const num = WO_COLS[c][1] === 'n';
    return h('td', { class: num ? 'num' : '', text: num && r[c] !== '' ? fmt(r[c]) : r[c] });
  })));
  $('uptbl').replaceChildren(h('thead', null, head), h('tbody', null, body));
  $('updiff').hidden = true;
  $('upgo').textContent = 'Simpan ke ' + d.sheet;
}

async function saveWoUpload() {
  const d = pending;
  if (!d || d.type !== 'WO') return;
  if (!d.sent) {
    const ok = await askConfirm('Simpan data master ' + d.sheet + '?',
      'File: ' + d.name + '. ' + fmt(d.rows.length) + ' baris untuk tahun ' + d.tahun + ' akan disimpan ke sheet "' + d.sheet + '". Loan No yang sudah ada diperbarui (tidak digandakan). '
      + 'Data yang sudah masuk ID_EWO, atau milik tahun lain, tidak diubah dan dilaporkan sebagai dilewati. Loan No baru pada cabang yang sudah memiliki ID_EWO tetap ditambahkan '
      + 'dan menunggu diproses menjadi ID_EWO berikutnya.', 'Simpan data master');
    if (!ok) return;
  }
  if (pending !== d) return;
  const tok = woTok();
  d.tot = d.tot || { inserted: 0, updated: 0, unchanged: 0, skippedLocked: 0, skippedYear: 0 }; d.samples = d.samples || [];
  $('upgo').disabled = true; upBar(d.sent / d.rows.length);
  while (d.sent < d.rows.length) {
    const part = d.rows.slice(d.sent, d.sent + WO_CHUNK);
    setNote('upmsg', 'Menyimpan ke sheet ' + d.sheet + '… ' + fmt(d.sent) + ' dari ' + fmt(d.rows.length) + ' baris.');
    const r = await api('webWoImport', { token: tok, file: d.name, offset: d.sent, rows: JSON.stringify(part) }, { timeoutMs: 180000 });
    if (pending !== d) return;                                // pengguna memilih file lain atau keluar saat proses berjalan
    if (r.code === 'SESSION_REQUIRED') return logout(r.error);
    if (!r.ok) {
      upBar(null); $('upgo').disabled = false; $('upgo').textContent = 'Lanjutkan Penyimpanan';
      return setNote('upmsg', (r.error || 'Gagal menyimpan data master.') + ' Terkirim ' + fmt(d.sent) + ' dari ' + fmt(d.rows.length)
        + ' baris. Tekan "Lanjutkan Penyimpanan" untuk melanjutkan; data yang sudah masuk tidak akan digandakan.', 'bad');
    }
    d.sent += part.length;
    Object.keys(d.tot).forEach(k => { d.tot[k] += Number(r[k]) || 0; });
    (r.samples || []).forEach(s => { if (d.samples.length < 8) d.samples.push(s); });
    upBar(d.sent / d.rows.length);
  }
  const t = d.tot, skipped = t.skippedLocked + t.skippedYear;
  setNote('upmsg', 'Selesai. Sheet "' + d.sheet + '" tahun ' + d.tahun + ': ' + fmt(t.inserted) + ' baru · ' + fmt(t.updated) + ' diperbarui · ' + fmt(t.unchanged) + ' tanpa perubahan · '
    + fmt(skipped) + ' dilewati' + (d.samples.length ? ' (contoh: ' + d.samples.join('; ') + ')' : '') + '.', skipped ? '' : 'good');
  toast('Data master ' + d.sheet + ' tersimpan.');
  wo.optAt = 0; wo.awoAt = 0;                                  // pilihan filter Eksekusi ikut berubah
  pending = null; $('file').value = ''; $('upgo').disabled = true; $('upgo').textContent = 'Pilih file berikutnya'; upBar(null);
}

/* ---------- Dasar halaman Eksekusi / Approve ---------- */

function woReset() {
  wo.epoch++;
  Object.assign(wo, { opt: null, ewos: [], optAt: 0, optLoading: false, pick: { b: '', j: '', y: '' }, busy: false, cur: null, rows: [], curAt: 0, q: '', page: 0,
    awo: null, awoAt: 0, awoLoading: false, awoJ: '', rep: woRepFresh(), rsb: rsbFresh(), manual: null, auto: null, jvId: '', verifId: '', reopenId: '', det: null, batal: null, setBusy: false });
  ['wo-manual', 'wo-auto', 'wo-jv', 'wo-verif', 'wo-detail', 'wo-reopen', 'wo-ttd', 'wo-batal', 'wo-clear'].forEach(id => { const dlg = $(id); if (dlg && dlg.open) dlg.close(); });
  $('ewo-work').hidden = true; $('ewo-hist-panel').hidden = true; $('ewo-q').value = '';
  ['ewo-branch', 'ewo-jenis', 'ewo-tahun', 'ewo-hist', 'ewo-tbl', 'awo-tbl', 'ewo-stats', 'ewo-pager'].forEach(id => $(id).replaceChildren());
  $('ewo-proses').disabled = true; $('ewo-refresh').disabled = false; $('awo-refresh').disabled = false;
  $('awo-jenis').value = '';
  woRepReset(); rsbReset();
  setNote('ewo-note', ''); setNote('awo-note', '');
}

// Panggilan ke server Write-Off. Hasil null = sesi berakhir (sudah diarahkan ke login) atau tampilan sudah diganti selama menunggu.
async function woCall(action, data, opts) {
  const epoch = wo.epoch;
  const r = await api(action, Object.assign({ token: woTok() }, data), opts);
  if (epoch !== wo.epoch) return null;
  if (r && r.code === 'SESSION_REQUIRED') { logout(r.error); return null; }
  return r;
}

const woRefError = (v, label) => {
  if (v.length > WO_REF_MAX) return label + ' maksimal ' + WO_REF_MAX + ' karakter.';
  if (/[\u0000-\u001F\u007F]/.test(v)) return label + ' memuat karakter yang tidak diperbolehkan.';
  if (/^[=+\-@]/.test(v)) return label + ' tidak boleh diawali tanda = + - atau @.';
  return '';
};
const woStatusPill = st => h('span', { class: 'wo-status' + (st === WO_ST_APPROVE ? ' is-approve' : ''), text: st });
const woEditable = () => !!wo.cur && wo.cur.st === WO_ST_CONFIRM;
const woRowOf = loan => wo.rows.find(r => r[WC.LOAN] === loan) || null;
function woReplaceRow(row) {
  const i = wo.rows.findIndex(r => r[WC.LOAN] === row[WC.LOAN]);
  if (i > -1) wo.rows[i] = row;
}
function woStatsOf(rows) {
  const s = { n: 0, batal: 0, total: 0, noThc: 0, noTak: 0 };
  rows.forEach(r => {
    if (r[WC.STATUS] === WO_BATAL) { s.batal++; return; }
    s.n++; s.total += Number(r[WC.PIUTANG]) || 0;
    if (!r[WC.THC]) s.noThc++;
    if (!r[WC.TAK]) s.noTak++;
  });
  s.total = woNum(s.total);
  return s;
}
const woBtn = (label, cls, act, extra) => h('button', Object.assign({ class: 'btn sm ' + cls, type: 'button', 'data-act': act, text: label }, extra || {}));

/* ---------- Menu Eksekusi Write-Off: tahap A (filter + Proses Data) ---------- */

function woEnter() {
  if (!state.user) return;
  if (!wo.opt || Date.now() - wo.optAt > AUTO_SYNC_MS) woLoadOptions();
  else { paintEwoFilters(); paintEwoHistory(); }
  if (wo.cur && Date.now() - wo.curAt > 20000) woReopen(true);
}

async function woLoadOptions() {
  if (wo.optLoading) return;
  wo.optLoading = true; $('ewo-refresh').disabled = true;
  if (!wo.opt) setNote('ewo-note', 'Memuat pilihan cabang, jenis, dan tahun…');
  const r = await woCall('webWoOptions', {}, { retry: true });
  wo.optLoading = false; $('ewo-refresh').disabled = false;
  if (!r) return;
  if (!r.ok) return setNote('ewo-note', r.error || 'Gagal memuat pilihan Write-Off.', 'bad');
  wo.opt = r.options || []; wo.ewos = r.ewos || []; wo.optAt = Date.now();
  setNote('ewo-note', '');
  paintEwoFilters(); paintEwoHistory();
}

// Muat ulang manual: pilihan filter, daftar ID_EWO, dan ID_EWO yang sedang dibuka.
async function woRefreshAll() {
  wo.optAt = 0;
  await woLoadOptions();
  if (wo.cur) await woReopen(false);
}

function woChoices() {
  const opt = wo.opt || [], p = wo.pick;
  return {
    branches: [...new Set(opt.map(o => o.b))].sort(collator.compare),
    jenis: ['R', 'K'].filter(j => opt.some(o => o.b === p.b && o.j === j)),
    years: [...new Set(opt.filter(o => o.b === p.b && o.j === p.j).map(o => o.y))].sort((a, b) => b - a),
  };
}
const woPicked = () => (wo.opt || []).find(o => o.b === wo.pick.b && o.j === wo.pick.j && String(o.y) === String(wo.pick.y)) || null;
function woFillSelect(sel, placeholder, items, value, disabled) {
  put(sel, h('option', { value: '', text: placeholder }), items.map(([v, t]) => h('option', { value: String(v), text: t })));
  sel.value = value === '' || value == null ? '' : String(value);
  sel.disabled = !!disabled;
}

function paintEwoFilters() {
  const p = wo.pick;
  let c = woChoices();
  // Pilihan yang sudah tidak tersedia dikosongkan; bila hanya ada satu pilihan, langsung dipilih.
  if (c.branches.indexOf(p.b) < 0) { p.b = ''; p.j = ''; p.y = ''; }
  if (!p.b && c.branches.length === 1) p.b = c.branches[0];
  c = woChoices();
  if (c.jenis.indexOf(p.j) < 0) { p.j = ''; p.y = ''; }
  if (!p.j && c.jenis.length === 1) p.j = c.jenis[0];
  c = woChoices();
  if (c.years.map(String).indexOf(String(p.y)) < 0) p.y = '';
  if (!p.y && c.years.length === 1) p.y = c.years[0];
  woFillSelect($('ewo-branch'), 'Pilih cabang…', c.branches.map(b => [b, b]), p.b, !c.branches.length);
  woFillSelect($('ewo-jenis'), 'Pilih jenis WO…', c.jenis.map(j => [j, 'WO ' + (j === 'R' ? 'Reguler' : 'Khusus')]), p.j, !p.b);
  woFillSelect($('ewo-tahun'), 'Pilih tahun…', c.years.map(y => [y, String(y)]), p.y, !p.j);
  const o = woPicked();
  $('ewo-proses').disabled = !o || !o.ok || wo.busy;          // ok (dari server): tidak ada ID_EWO sebelumnya yang belum "Sesuai" dan ada data siap diproses
  if (wo.busy) return;
  if (wo.opt && !wo.opt.length && !wo.ewos.length) {
    const admin = state.user && state.user.tingkat === 'ADMINISTRATOR';
    setNote('ewo-note', admin ? 'Belum ada data master Write-Off. Unggah file "WO Reguler yyyy.xlsx" atau "WO Khusus yyyy.xlsx" di menu Import Data.'
      : state.user && state.user.tingkat === 'USER 3' ? 'Belum ada data Write-Off pada cabang akun Anda' + (state.user.cabang ? ' (' + state.user.cabang + ')' : '') + '.'
      : 'Belum ada data Write-Off pada cabang di Wilayah Kerja akun Anda.');
  } else if (o && o.blk) {
    // ID_EWO sebelumnya belum "Sesuai": ID_EWO baru belum boleh dibuat.
    setNote('ewo-note', 'Cabang ini sudah memiliki ' + o.blk.id + ' (status ' + o.blk.st + ', Verifikasi Nilai WO ' + (o.blk.ver ? '"' + o.blk.ver + '"' : 'belum diisi')
      + '). ID_EWO baru hanya dapat dibuat setelah Verifikasi Nilai WO ID_EWO sebelumnya berstatus "' + WO_VER_OK + '". Buka ID_EWO tersebut dari daftar di bawah.', 'bad');
  } else if (o && !o.n) {
    // Tidak ada data dengan "Verifikasi Nilai WO" dan "Status Baris" kosong: Proses Data gagal, ID_EWO tidak terbentuk.
    setNote('ewo-note', 'Tidak ada data WO ' + (o.j === 'R' ? 'Reguler' : 'Khusus') + ' tahun ' + o.y + ' cabang ' + o.b + ' yang kolom "Verifikasi Nilai WO" dan "Status Baris"-nya masih kosong'
      + (o.k ? ' (semua ID_EWO sebelumnya sudah "' + WO_VER_OK + '")' : '') + '. Input WO tidak dapat diproses dan ID_EWO tidak dibentuk.', 'bad');
  } else if (o) {
    setNote('ewo-note', fmt(o.n) + ' data WO ' + (o.j === 'R' ? 'Reguler' : 'Khusus') + ' tahun ' + o.y + ' cabang ' + o.b + ' siap diproses'
      + (o.k ? ' menjadi ID_EWO baru (ID_EWO sebelumnya: ' + fmt(o.k) + ', semuanya "' + WO_VER_OK + '").' : '.'));
  } else setNote('ewo-note', wo.opt ? 'Pilih cabang, jenis WO, dan tahun WO, lalu tekan Proses Data.' : '');
}

function woFilterChanged(id) {
  const p = wo.pick;
  if (id === 'ewo-branch') { p.b = $(id).value; p.j = ''; p.y = ''; }
  else if (id === 'ewo-jenis') { p.j = $(id).value; p.y = ''; }
  else p.y = $(id).value;
  paintEwoFilters();
}

function paintEwoHistory() {
  const list = wo.ewos || [];
  $('ewo-hist-panel').hidden = !list.length;
  if (!list.length) return $('ewo-hist').replaceChildren();
  const th = (t, num) => h('th', { scope: 'col', class: num ? 'num' : '', text: t });
  const rows = list.map(e => h('tr', { class: wo.cur && wo.cur.id === e.id ? 'is-cur' : '' },
    h('td', { class: 'wo-name', 'data-label': 'ID_EWO', text: e.id }), h('td', { 'data-label': 'Cabang', text: e.b }), h('td', { 'data-label': 'Jenis', text: e.jenis }),
    h('td', { 'data-label': 'Tahun', text: String(e.y) }), h('td', { 'data-label': 'Status' }, woStatusPill(e.st)),
    h('td', { class: 'num', 'data-label': 'Data', text: fmt(e.n) }), h('td', { class: 'num', 'data-label': 'PIUTANG YANG DI WO', text: fmt(e.total) }),
    h('td', { 'data-label': 'Dibuat', text: fmtStamp(e.at) }),
    h('td', { class: 'num wo-act' }, h('button', { class: 'btn sm ghost', type: 'button', 'data-act': 'open', 'data-id': e.id, text: e.st === WO_ST_CONFIRM ? 'Buka' : 'Lihat' }))));
  $('ewo-hist').className = 'wo-hist wo-cards';
  put($('ewo-hist'), h('thead', null, h('tr', null, th('ID_EWO'), th('Cabang'), th('Jenis'), th('Tahun'), th('Status'), th('Data', true), th('PIUTANG YANG DI WO', true), th('Dibuat'), th('', true))),
    h('tbody', null, rows));
}

async function woProses() {
  const o = woPicked();
  if (!o || !o.ok || wo.busy) return;
  const ok = await askConfirm('Proses data Write-Off?',
    'Cabang ' + o.b + ' · WO ' + (o.j === 'R' ? 'Reguler' : 'Khusus') + ' · Tahun ' + o.y + ' · ' + fmt(o.n) + ' data yang kolom "Verifikasi Nilai WO" dan "Status Baris"-nya masih kosong. '
    + 'Satu ID_EWO baru berstatus "Confirm" akan dibuat' + (o.k ? ' (ID_EWO sebelumnya untuk kombinasi ini tidak diubah)' : '') + '.',
    'Proses Data');
  if (!ok) return;
  wo.busy = true; $('ewo-proses').disabled = true; setNote('ewo-note', 'Memproses data dan membuat ID_EWO…');
  const r = await woCall('webWoProses', { branch: o.b, jenis: o.j, tahun: o.y }, { timeoutMs: 180000 });
  wo.busy = false;
  if (!r) return;
  if (!r.ok) { setNote('ewo-note', r.error || 'Proses Data gagal.', 'bad'); wo.optAt = 0; woLoadOptions(); return; }
  toast(r.head.id + ' dibuat.');
  woShowWork(r);
  wo.awoAt = 0;
  wo.optAt = 0; await woLoadOptions();
}

async function woOpenEwo(id) {
  setNote('ewo-note', 'Membuka ' + id + '…');
  const r = await woCall('webWoOpen', { ewo: id }, { retry: true, timeoutMs: 120000 });
  if (!r) return;
  if (!r.ok) return setNote('ewo-note', r.error || 'Gagal membuka ID_EWO.', 'bad');
  setNote('ewo-note', '');
  woShowWork(r, true);
}

// Muat ulang ID_EWO yang sedang dibuka (status bisa berubah dari menu Approve). quiet = tanpa pesan.
async function woReopen(quiet) {
  if (!wo.cur) return;
  const id = wo.cur.id;
  const r = await woCall('webWoOpen', { ewo: id }, { retry: true, timeoutMs: 120000 });
  if (!r || !wo.cur || wo.cur.id !== id) return;
  if (!r.ok) { if (!quiet) setNote('ewo-note', r.error || 'Gagal memuat ulang ID_EWO.', 'bad'); return; }
  wo.cur = r.head; wo.rows = r.rows; wo.curAt = Date.now();
  paintEwoWork(); paintEwoHistory();
}

function woShowWork(payload, scroll) {
  wo.cur = payload.head; wo.rows = payload.rows || []; wo.curAt = Date.now(); wo.q = ''; wo.page = 0; $('ewo-q').value = '';
  paintEwoWork(); paintEwoHistory();
  if (scroll !== false) $('ewo-work').scrollIntoView({ block: 'start', behavior: 'smooth' });
}
function woCloseWork() {
  wo.cur = null; wo.rows = []; wo.q = ''; wo.page = 0;
  paintEwoWork(); paintEwoHistory();
}

/* ---------- Eksekusi Write-Off: tahap B (area kerja) dan C (Eksekusi Data) ---------- */

function paintEwoWork() {
  const c = wo.cur;
  $('ewo-work').hidden = !c;
  if (!c) return;
  const editable = woEditable(), s = woStatsOf(wo.rows);
  $('ewo-id').textContent = c.id;
  const st = $('ewo-status'); st.textContent = c.st; st.className = 'wo-status' + (editable ? '' : ' is-approve');
  $('ewo-work-meta').textContent = 'WO ' + c.jenis + ' · Tahun ' + c.y + ' · Cabang ' + c.b + ' · dibuat ' + fmtStamp(c.at) + ' oleh ' + c.by;
  const banner = $('ewo-banner');
  if (!editable) {
    banner.className = 'wo-banner'; banner.hidden = false;
    banner.textContent = 'ID_EWO ini berstatus "' + c.st + '" sehingga hanya bisa dilihat. Data dapat diedit lagi hanya bila statusnya kembali "' + WO_ST_CONFIRM + '".';
  } else if (c.ver === WO_VER_NO) {
    banner.className = 'wo-banner is-bad'; banner.hidden = false;
    banner.textContent = 'Verifikasi Nilai WO sebelumnya berhasil dijawab "' + WO_VER_NO + '" (' + fmtStamp(c.verAt) + '). Periksa dan revisi nilainya, lalu tekan Eksekusi Data kembali.';
  } else banner.hidden = true;
  const chip = (text, warn) => h('span', { class: 'st' + (warn ? ' warn' : ''), text });
  put($('ewo-stats'), chip(fmt(s.n) + ' data aktif'), s.batal ? chip(fmt(s.batal) + ' dibatalkan', true) : null, chip('PIUTANG YANG DI WO ' + rp(s.total)),
    chip('Belum No. THC: ' + fmt(s.noThc), s.noThc > 0), chip('Belum No. TAK: ' + fmt(s.noTak), s.noTak > 0), c.jv ? chip('No JV: ' + c.jv) : null);
  paintEwoTable();
  $('ewo-exec').disabled = !editable || s.n === 0;
  $('ewo-final-note').textContent = editable
    ? 'Setelah Eksekusi Data, status menjadi "' + WO_ST_APPROVE + '" dan data hanya bisa dilihat.' + (s.noThc || s.noTak ? ' Perhatian: ' + fmt(s.noThc) + ' data belum memiliki No. THC dan ' + fmt(s.noTak) + ' data belum memiliki No. TAK.' : '')
    : 'Sudah dieksekusi' + (c.exAt ? ' pada ' + fmtStamp(c.exAt) : '') + (c.exBy ? ' oleh ' + c.exBy : '') + '. Menunggu Input JV dan Verifikasi Nilai WO di menu Approve Write-Off.';
}

function woView() {
  const q = wo.q.trim().toLocaleLowerCase('id');
  if (!q) return wo.rows;
  return wo.rows.filter(r => [r[WC.CTR], r[WC.CLIENT], r[WC.NAME], r[WC.LOAN]].some(v => String(v).toLocaleLowerCase('id').includes(q)));
}

function paintEwoTable() {
  const view = woView(), editable = woEditable();
  const pages = Math.max(1, Math.ceil(view.length / WO_PAGE));
  if (wo.page >= pages) wo.page = pages - 1;
  const slice = view.slice(wo.page * WO_PAGE, (wo.page + 1) * WO_PAGE);
  const th = (t, cls) => h('th', { scope: 'col', class: cls || '', text: t });
  const ref = (v, label) => v ? h('td', { class: 'wo-ref', 'data-label': label, text: v }) : h('td', { class: 'is-dim', 'data-label': label, text: '—' });
  const rows = slice.map(r => {
    const batal = r[WC.STATUS] === WO_BATAL, loan = r[WC.LOAN];
    const act = !editable ? h('span', { class: 'is-dim', text: 'Hanya dilihat' })
      : h('div', { class: 'wo-rowbtns' },
        batal ? null : woBtn('Manual', 'ghost', 'manual', { 'data-loan': loan }),
        batal ? null : woBtn('Otomatis', 'ghost', 'auto', { 'data-loan': loan }),
        woBtn(batal ? 'Pulihkan' : 'Batal', batal ? 'warn' : 'danger', 'batal', { 'data-loan': loan }));
    return h('tr', { class: batal ? 'is-batal' : '' },
      h('td', { class: 'num', 'data-label': 'No.', text: woPlain(r[WC.NO]) }), h('td', { 'data-label': 'Ctr ID', text: woPlain(r[WC.CTR]) }),
      h('td', { 'data-label': 'Client ID', text: woPlain(r[WC.CLIENT]) }),
      h('td', { class: 'wo-name', 'data-label': 'Client Name' }, r[WC.NAME], batal ? h('span', { class: 'wo-flag', text: 'BATAL' }) : null),
      h('td', { 'data-label': 'Loan No', text: loan }), ref(r[WC.THC], 'No. THC'), ref(r[WC.TAK], 'No. TAK'), h('td', { class: 'wo-act' }, act));
  });
  $('ewo-tbl').className = 'wo-tbl wo-cards';
  put($('ewo-tbl'), h('thead', null, h('tr', null, th('No.', 'num'), th('Ctr ID'), th('Client ID'), th('Client Name'), th('Loan No'), th('No. THC'), th('No. TAK'), th('Aksi'))),
    h('tbody', null, rows.length ? rows : h('tr', null, h('td', { colspan: '8', class: 'wo-empty', text: wo.rows.length ? 'Tidak ada data yang cocok dengan pencarian.' : 'ID_EWO ini tidak memiliki data.' }))));
  const total = view.length, from = total ? wo.page * WO_PAGE + 1 : 0, to = Math.min(total, (wo.page + 1) * WO_PAGE);
  const go = d => () => { wo.page = Math.max(0, Math.min(pages - 1, wo.page + d)); paintEwoTable(); $('ewo-tbl').scrollIntoView({ block: 'nearest' }); };
  put($('ewo-pager'), h('span', { text: total ? 'Menampilkan ' + fmt(from) + '–' + fmt(to) + ' dari ' + fmt(total) + ' data' : '' }),
    pages > 1 ? h('div', { class: 'pnav' },
      h('button', { class: 'btn ghost sm', type: 'button', disabled: wo.page === 0, onclick: go(-1), text: 'Sebelumnya' }),
      h('span', { text: 'Halaman ' + (wo.page + 1) + ' / ' + pages }),
      h('button', { class: 'btn ghost sm', type: 'button', disabled: wo.page >= pages - 1, onclick: go(1), text: 'Berikutnya' })) : null);
}

function woRowClick(e) {
  const b = e.target.closest('button[data-act]');
  if (!b || b.disabled || !woEditable()) return;
  if (b.dataset.act === 'manual') woOpenManual(b.dataset.loan);
  else if (b.dataset.act === 'auto') woOpenAuto();
  else if (b.dataset.act === 'batal') woToggleBatal(b.dataset.loan, b);
}
function woHistClick(e) {
  const b = e.target.closest('button[data-act="open"]');
  if (b && !b.disabled) woOpenEwo(b.dataset.id);
}

// Batal -> dialog Alasan Batal (woBatalSubmit); Pulihkan -> langsung (alasan dikosongkan server).
async function woToggleBatal(loan, btn) {
  const row = woRowOf(loan);
  if (!row || !woEditable()) return;
  if (row[WC.STATUS] !== WO_BATAL) return woOpenBatal(loan);
  btn.disabled = true;
  const r = await woCall('webWoBatalRow', { ewo: wo.cur.id, loan, batal: '0' }, { timeoutMs: 90000 });
  if (!r) return;
  if (!r.ok) { btn.disabled = false; return toast(r.error || 'Gagal mengubah status data.', true); }
  woReplaceRow(r.row); wo.cur = r.head; wo.curAt = Date.now();
  paintEwoWork();
  toast('Loan No ' + loan + ' dipulihkan.');
}

/* ----- Dialog Alasan Batal ----- */

function woOpenBatal(loan) {
  const row = woRowOf(loan);
  if (!row || !woEditable()) return;
  wo.batal = { loan };
  const opts = WO_BATAL_OPTIONS.concat(['Lainnya']);
  put($('wb-form'),
    h('h2', { id: 'wb-title', text: 'Alasan Batal' }),
    h('p', { class: 'dlg-sub', text: wo.cur.id + ' · Loan No ' + loan + ' · ' + row[WC.NAME] }),
    h('p', { class: 'wm-hint', text: 'Data ini tidak akan ikut dieksekusi dan tidak masuk file unduhan berikutnya. Pilih alasan pembatalan. Data bisa dipulihkan selama status ID_EWO masih "' + WO_ST_CONFIRM + '".' }),
    h('fieldset', { class: 'wo-reasons' },
      h('legend', { class: 'sr-only', text: 'Alasan batal' }),
      ...opts.map((t, i) => h('label', { class: 'chk' }, h('input', { type: 'radio', name: 'wb-alasan', value: String(i) }), h('span', { text: (i + 1) + '. ' + t }))),
      h('div', { class: 'wj-field', id: 'wb-other-wrap', hidden: true }, h('label', { for: 'wb-other', text: 'Alasan lainnya' }),
        h('input', { id: 'wb-other', type: 'text', maxlength: String(WO_ALASAN_MAX), autocomplete: 'off', placeholder: 'Tulis alasan pembatalan' }))),
    h('p', { class: 'wm-err', id: 'wb-err', role: 'alert' }),
    h('div', { class: 'dlg-act' },
      h('button', { class: 'btn ghost', type: 'button', onclick: () => $('wo-batal').close(), text: 'Kembali' }),
      h('button', { class: 'btn fit', type: 'submit', id: 'wb-save', text: 'Batalkan data' })));
  $('wb-form').querySelectorAll('input[name="wb-alasan"]').forEach(r => r.addEventListener('change', () => {
    const other = r.value === String(WO_BATAL_OPTIONS.length) && r.checked;
    $('wb-other-wrap').hidden = !other; $('wb-err').textContent = '';
    if (other) $('wb-other').focus();
  }));
  $('wo-batal').showModal();
}

async function woBatalSubmit(ev) {
  ev.preventDefault();
  const b = wo.batal, err = $('wb-err'), save = $('wb-save');
  if (!b || !save) return;
  const picked = $('wb-form').querySelector('input[name="wb-alasan"]:checked');
  let alasan = '', msg = '';
  if (!picked) msg = 'Pilih salah satu alasan batal.';
  else if (Number(picked.value) < WO_BATAL_OPTIONS.length) alasan = WO_BATAL_OPTIONS[Number(picked.value)];
  else {
    alasan = $('wb-other').value.replace(/\s+/g, ' ').trim();
    msg = !alasan ? 'Alasan lainnya wajib diisi.' : alasan.length > WO_ALASAN_MAX ? 'Alasan maksimal ' + WO_ALASAN_MAX + ' karakter.' : /^[=+\-@]/.test(alasan) ? 'Alasan tidak boleh diawali tanda = + - atau @.' : '';
  }
  err.textContent = msg;
  if (msg) return;
  save.disabled = true;
  const r = await woCall('webWoBatalRow', { ewo: wo.cur.id, loan: b.loan, batal: '1', alasan }, { timeoutMs: 90000 });
  if (!r) return;
  save.disabled = false;
  if (!r.ok) { err.textContent = r.error || 'Gagal membatalkan data.'; return; }
  woReplaceRow(r.row); wo.cur = r.head; wo.curAt = Date.now(); wo.batal = null;
  $('wo-batal').close(); paintEwoWork();
  toast('Loan No ' + b.loan + ' dibatalkan.');
}

/* ----- Manual ----- */

function woOpenManual(loan) {
  const row = woRowOf(loan);
  if (!row || !woEditable() || row[WC.STATUS] === WO_BATAL) return;
  wo.manual = { loan, orig: row.slice() };
  const input = (id, val, extra) => h('input', Object.assign({ id, type: 'text', value: val, autocomplete: 'off', spellcheck: 'false' }, extra || {}));
  const refField = (id, label, val) => h('div', null, h('label', { for: id, text: label }), input(id, val, { maxlength: String(WO_REF_MAX) }));
  const fact = c => h('div', { class: c === WC.BRANCH || c === WC.NAME || c === WC.PRODUCT ? 'wide' : '' }, h('dt', { text: WO_COLS[c][0] }), h('dd', { text: woPlain(row[c]) }));
  const nominal = WO_NOMINAL.map(([key, c]) => h('div', null, h('label', { for: 'wm-' + key, text: WO_COLS[c][0] }),
    input('wm-' + key, row[c] === '' ? '0' : String(row[c]), { inputmode: 'decimal' })));
  put($('wm-form'),
    h('h2', { id: 'wm-title', text: 'Input Manual' }),
    h('p', { class: 'dlg-sub', text: wo.cur.id + ' · Loan No ' + loan }),
    h('div', { class: 'wm-ref' }, refField('wm-thc', 'No. THC', row[WC.THC]), refField('wm-tak', 'No. TAK', row[WC.TAK])),
    h('h3', { text: 'Data pinjaman' }),
    h('dl', { class: 'wm-facts' }, [WC.NO, WC.BRANCH, WC.LOAN, WC.CTR, WC.CLIENT, WC.NAME, WC.PRODUCT, WC.PERIODE, WC.TAHUN, WC.JV, WC.VERIF].map(fact)),
    h('h3', { text: 'Nominal (bisa diubah)' }),
    h('div', { class: 'wm-grid' }, nominal),
    h('div', { class: 'wm-calc' },
      h('div', null, h('span', { text: 'Total Simpanan' }), h('b', { id: 'wm-totsimp' })),
      h('div', null, h('span', { text: 'PIUTANG YANG DI WO' }), h('b', { id: 'wm-piutang' }))),
    h('p', { class: 'wm-hint', text: 'Total Simpanan = jumlah 7 simpanan. PIUTANG YANG DI WO = Total Balance − Total Simpanan. Keduanya dihitung ulang server hanya bila ada nominal yang Anda ubah.' }),
    h('p', { class: 'wm-err', id: 'wm-err', role: 'alert' }),
    h('div', { class: 'dlg-act' },
      h('button', { class: 'btn ghost', type: 'button', onclick: () => $('wo-manual').close(), text: 'Batal' }),
      h('button', { class: 'btn fit', type: 'submit', id: 'wm-save', text: 'Simpan' })));
  woManualCalc();
  $('wo-manual').showModal();
  $('wm-thc').focus();
}

// Baca 8 nominal dari form. Kosong = 0. Mengembalikan { out: {tb: n, ...}, bad: [[input, pesan]] }.
function woManualParse() {
  const out = {}, bad = [];
  WO_NOMINAL.forEach(([key, c]) => {
    const el = $('wm-' + key), t = el.value.trim();
    if (t === '') { out[key] = 0; return; }
    const x = toNum(t), label = WO_COLS[c][0];
    if (x === '') bad.push([el, label + ' harus berupa angka.']);
    else if (x < 0) bad.push([el, label + ' tidak boleh negatif.']);
    else if (x > 1e15) bad.push([el, label + ' terlalu besar.']);
    else out[key] = x;
  });
  return { out, bad };
}

function woManualCalc(ev) {
  const m = wo.manual;
  if (!m || !$('wm-totsimp')) return;
  if (ev && ev.target && ev.target.classList) {                // pengguna mengetik: tanda salah & pesan lama dihapus
    ev.target.classList.remove('bad'); ev.target.removeAttribute('aria-invalid'); $('wm-err').textContent = '';
  }
  const { out, bad } = woManualParse(), o = m.orig;
  const tot = $('wm-totsimp'), pi = $('wm-piutang');
  if (bad.length) { tot.textContent = '—'; pi.textContent = '—'; pi.className = ''; return; }
  const changed = WO_NOMINAL.some(([k, c]) => Math.abs(out[k] - (o[c] === '' ? 0 : o[c])) > 1e-9);
  let simp, piutang;
  if (changed) {
    simp = woNum(WO_NOMINAL.slice(1).reduce((a, [k]) => a + out[k], 0));
    piutang = woNum(out.tb - simp);
  } else { simp = o[WC.TOTSIMP]; piutang = o[WC.PIUTANG]; }   // tidak ada yang diubah: nilai dari Excel dipertahankan
  tot.textContent = simp === '' ? '—' : rp(simp);
  pi.textContent = piutang === '' ? '—' : rp(piutang);
  pi.className = piutang !== '' && piutang < 0 ? 'is-neg' : '';
}

async function woManualSubmit(e) {
  e.preventDefault();
  const m = wo.manual;
  if (!m || !woEditable()) return;
  const err = $('wm-err'), save = $('wm-save');
  err.textContent = '';
  $('wm-form').querySelectorAll('input').forEach(i => { i.classList.remove('bad'); i.removeAttribute('aria-invalid'); });
  const thc = $('wm-thc').value.trim(), tak = $('wm-tak').value.trim();
  const { out, bad } = woManualParse();
  [['No. THC', thc, $('wm-thc')], ['No. TAK', tak, $('wm-tak')]].reverse().forEach(([label, v, el]) => { const msg = woRefError(v, label); if (msg) bad.unshift([el, msg]); });
  if (bad.length) {
    bad.forEach(([el]) => { el.classList.add('bad'); el.setAttribute('aria-invalid', 'true'); });
    err.textContent = bad[0][1]; bad[0][0].focus();
    return;
  }
  save.disabled = true;
  const r = await woCall('webWoSaveRow', Object.assign({ ewo: wo.cur.id, loan: m.loan, thc, tak }, out), { timeoutMs: 90000 });
  if (!r) return;
  save.disabled = false;
  if (!r.ok) { err.textContent = r.error || 'Gagal menyimpan data.'; return; }
  woReplaceRow(r.row); wo.cur = r.head; wo.curAt = Date.now(); wo.manual = null;
  $('wo-manual').close();
  paintEwoWork();
  toast('Data Loan No ' + m.loan + ' tersimpan.');
}

/* ----- Otomatis: unggah Excel hasil unduhan, isi No. THC / No. TAK menurut Loan No ----- */

// Murni: array baris Excel -> { rows: [[Loan No, No. THC, No. TAK, Total Balance, Simp. Wajib, ..., Simp. Sipadan]], missing: [judul kolom nominal yang tidak ada] }.
// Header wajib di baris 1 (sama seperti file unduhan). Kolom nominal opsional: bila tidak ada atau selnya kosong, nilai di sistem TIDAK diubah.
function woParseAuto(aoa) {
  const head = aoa[0] || [];
  const at = key => head.findIndex(v => woCompact(v) === key);
  const iLoan = at('loanno'), iThc = at('nothc'), iTak = at('notak');
  if (iLoan < 0 || iThc < 0 || iTak < 0) {
    throw new Error('Header tabel harus berada di baris 1 dan memuat "Loan No", "No. THC", dan "No. TAK" seperti pada file hasil unduhan. File ini tidak sesuai sehingga tidak diproses.');
  }
  const nomAt = WO_NOMINAL.map(([, c]) => head.findIndex(v => WO_COLS[c][2].indexOf(woCompact(v)) > -1));
  const missing = WO_NOMINAL.map(([, c], k) => nomAt[k] < 0 ? WO_COLS[c][0] : '').filter(Boolean);
  const rows = [];
  for (let i = 1; i < aoa.length; i++) {
    const r = aoa[i] || [], loan = toTxt(r[iLoan]);
    if (!loan) continue;
    const nom = WO_NOMINAL.map(([, c], k) => {
      const v = nomAt[k] < 0 ? '' : r[nomAt[k]];
      if (blank(v)) return '';
      const x = toNum(typeof v === 'string' ? v.replace(/^Rp\.?/i, '').trim() : v), label = WO_COLS[c][0];
      if (x === '') throw new Error('Nilai angka tidak valid pada baris Excel ' + (i + 1) + ', kolom "' + label + '": ' + String(v).slice(0, 30) + '.');
      if (x < 0) throw new Error('"' + label + '" tidak boleh negatif (baris Excel ' + (i + 1) + ').');
      return x;
    });
    rows.push([loan, toTxt(r[iThc]), toTxt(r[iTak])].concat(nom));
  }
  if (!rows.length) throw new Error('Tidak ada baris data (Loan No terisi) di bawah header.');
  return { rows, missing };
}

function woOpenAuto() {
  if (!wo.cur || !woEditable()) return;
  wo.auto = null;
  const input = h('input', { type: 'file', id: 'wa-file', accept: '.xlsx,.xls' });
  input.addEventListener('change', () => { const f = input.files && input.files[0]; if (f) woReadAuto(f); });
  put($('wa-body'),
    h('h2', { id: 'wa-title', text: 'Otomatis: isi No. THC, No. TAK & nilai' }),
    h('p', { class: 'dlg-sub', text: wo.cur.id + ' · WO ' + wo.cur.jenis + ' ' + wo.cur.y + ' · ' + wo.cur.b }),
    h('ol', { class: 'wo-steps' },
      h('li', { text: 'Tekan Unduh pada ID_EWO ini, lalu isi kolom No. THC dan No. TAK. Total Balance dan ketujuh simpanan di file itu juga boleh dikoreksi.' }),
      h('li', { text: 'Unggah file itu di sini. Data dicocokkan lewat Loan No, lalu No. THC, No. TAK, Total Balance, dan masing-masing simpanan disimpan ke sheet WO (tidak membuat sheet baru).' }),
      h('li', { text: 'Total Simpanan dan PIUTANG YANG DI WO dihitung ulang oleh sistem dari nilai yang baru (isi kedua kolom itu di file diabaikan). Sel yang dikosongkan tidak mengubah isi yang sudah ada.' }),
      h('li', { text: 'Header tabel harus tetap di baris 1 seperti file unduhan. Berlaku untuk seluruh baris di ID_EWO ini, bukan hanya baris yang tombolnya ditekan.' })),
    h('label', { class: 'wo-pick', for: 'wa-file' }, h('b', { text: 'Pilih file Excel hasil unduhan' }), h('span', { text: '.xlsx atau .xls · maksimal ' + UP_MAX_MB + ' MB' }), input),
    h('div', { id: 'wa-result' }),
    h('p', { class: 'note', id: 'wa-msg', role: 'status' }),
    h('div', { class: 'dlg-act' },
      h('button', { class: 'btn ghost', type: 'button', onclick: () => $('wo-auto').close(), text: 'Tutup' }),
      h('button', { class: 'btn fit', type: 'button', id: 'wa-go', disabled: true, onclick: woApplyAuto, text: 'Isi Data' })));
  $('wo-auto').showModal();
}

async function woReadAuto(file) {
  const ewo = wo.cur && wo.cur.id, go = $('wa-go');
  wo.auto = null; go.disabled = true; put($('wa-result')); setNote('wa-msg', 'Membaca file…');
  try {
    const hit = WO_EWO_RE.exec(file.name);
    if (hit && hit[0].toUpperCase() !== ewo) throw new Error('File ini milik ' + hit[0].toUpperCase() + ', sedangkan ID_EWO yang dibuka adalah ' + ewo + '.');
    const parsed = woParseAuto(await woWorkbookRows(file, WO_SHEET[wo.cur.j], true)), rows = parsed.rows;
    if (!wo.cur || wo.cur.id !== ewo) return;
    const send = rows.filter(r => r.slice(1).some(v => v !== ''));
    // Perkiraan jumlah data yang nilainya berbeda dari yang tampil sekarang (server tetap menghitung ulang dan menentukan hasil akhir).
    const diff = send.filter(r => {
      const cur = woRowOf(r[0]);
      return cur && WO_NOMINAL.some(([, c], k) => r[3 + k] !== '' && Math.abs(r[3 + k] - (cur[c] === '' ? 0 : cur[c])) > 1e-9);
    }).length;
    wo.auto = { name: file.name, send };
    put($('wa-result'), h('div', { class: 'wo-result' }, h('div', { class: 'stats' },
      h('span', { class: 'st', text: file.name }), h('span', { class: 'st', text: fmt(rows.length) + ' baris di file' }),
      h('span', { class: 'st' + (send.length ? '' : ' warn'), text: fmt(send.length) + ' baris berisi No. THC / No. TAK / nilai' }),
      h('span', { class: 'st', text: '± ' + fmt(diff) + ' baris nilainya berbeda dari data sekarang' })),
      parsed.missing.length ? h('p', { text: 'Kolom tidak ada di file sehingga nilainya tidak diubah: ' + parsed.missing.join(', ') + '.' }) : null));
    if (!send.length) setNote('wa-msg', 'Tidak ada baris yang memuat No. THC, No. TAK, atau nilai. Isi kolomnya di file Excel terlebih dahulu.', 'bad');
    else setNote('wa-msg', hit ? 'File siap diproses. Tekan Isi Data untuk menyimpan.' : 'Nama file tidak memuat ID_EWO; pencocokan hanya memakai Loan No. Tekan Isi Data untuk menyimpan.');
    go.disabled = !send.length;
  } catch (e) { setNote('wa-msg', e.message || 'File tidak bisa dibaca.', 'bad'); }
}

async function woApplyAuto() {
  const a = wo.auto, go = $('wa-go');
  if (!a || !woEditable()) return;
  go.disabled = true; setNote('wa-msg', 'Menyimpan No. THC, No. TAK, dan nilai simpanan…');
  const r = await woCall('webWoFillAuto', { ewo: wo.cur.id, rows: JSON.stringify(a.send) }, { timeoutMs: 180000 });
  if (!r) return;
  if (!r.ok) { go.disabled = false; return setNote('wa-msg', r.error || 'Gagal mengisi data.', 'bad'); }
  wo.rows = r.rows; wo.cur = r.head; wo.curAt = Date.now(); wo.auto = null;
  paintEwoWork();
  const x = r.result, chip = (t, warn) => h('span', { class: 'st' + (warn ? ' warn' : ''), text: t });
  put($('wa-result'), h('div', { class: 'wo-result' },
    h('div', { class: 'stats' }, chip(fmt(x.updated) + ' data diperbarui'), chip(fmt(x.nominal || 0) + ' data nilainya berubah (total dihitung ulang)'), chip(fmt(x.unchanged) + ' tanpa perubahan'),
      x.notFound ? chip(fmt(x.notFound) + ' Loan No tidak ditemukan', true) : null, x.skippedBatal ? chip(fmt(x.skippedBatal) + ' dilewati (Batal)', true) : null,
      x.invalid ? chip(fmt(x.invalid) + ' tidak valid (No. THC/TAK diawali = + - @ atau terlalu panjang, atau nilai bukan angka / negatif)', true) : null),
    x.samples && x.samples.length ? h('p', { text: 'Loan No tidak ditemukan, contoh: ' + x.samples.join(', ') }) : null,
    x.invalidSamples && x.invalidSamples.length ? h('p', { text: 'Tidak valid, contoh: ' + x.invalidSamples.join('; ') }) : null));
  setNote('wa-msg', 'Selesai. ' + fmt(x.updated) + ' data diperbarui, ' + fmt(x.nominal || 0) + ' di antaranya nilainya berubah.', 'good');
  toast(fmt(x.updated) + ' data diperbarui dari file Otomatis.');
}

/* ----- Unduh Excel (dipakai juga oleh Otomatis) ----- */

async function woDownload() {
  const c = wo.cur;
  if (!c) return;
  const rows = wo.rows.filter(r => r[WC.STATUS] !== WO_BATAL);
  if (!rows.length) return toast('Tidak ada data aktif untuk diunduh.', true);
  const btn = $('ewo-download');
  btn.disabled = true; setNote('ewo-note', 'Menyiapkan file Excel…');
  try {
    await Promise.all([loadXLSX(), nextFrame()]);
    const cols = WO_COLS.slice(0, WO_NBASE);
    const aoa = [cols.map(x => x[0])].concat(rows.map(r => cols.map((x, k) => x[1] === 'n' ? (r[k] === '' ? '' : Number(r[k])) : String(r[k]))));
    const ws = XLSX.utils.aoa_to_sheet(aoa), L = k => XLSX.utils.encode_col(k);
    const fill = XS.text && { patternType: 'solid', fgColor: { rgb: 'FFF4BD' } };
    const widths = cols.map((x, k) => Math.max(10, x[0].length + 3, ...rows.slice(0, 300).map(r => String(r[k]).length + 2)));
    cols.forEach((x, k) => {
      const head = ws[L(k) + '1'];
      head.s = XS.head;
      for (let i = 0; i < rows.length; i++) {
        const cell = ws[L(k) + (i + 2)];
        if (!cell) continue;
        if (x[1] === 'n') { cell.s = XS.num; cell.z = XL_NUM; }
        else { cell.s = k === WC.THC || k === WC.TAK ? Object.assign({}, XS.text, { fill }) : XS.text; if (k === WC.THC || k === WC.TAK) cell.z = '@'; }
      }
    });
    ws['!cols'] = widths.map(w => ({ wch: Math.min(w, 45) }));
    ws['!rows'] = [{ hpt: 32 }];
    ws['!freeze'] = { ySplit: '1', topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' };
    ws['!autofilter'] = { ref: 'A1:' + L(cols.length - 1) + (rows.length + 1) };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, WO_SHEET[c.j]);
    XLSX.writeFile(wb, 'WO ' + c.jenis + ' ' + c.y + ' ' + c.id + '.xlsx');
    setNote('ewo-note', 'File Excel "WO ' + c.jenis + ' ' + c.y + ' ' + c.id + '.xlsx" berhasil dibuat: ' + fmt(rows.length) + ' data aktif.', 'good');
  } catch (e) { setNote('ewo-note', e.message || 'Gagal membuat file Excel.', 'bad'); }
  btn.disabled = false;
}

/* ----- Tahap C: Eksekusi Data ----- */

async function woExecute() {
  const c = wo.cur;
  if (!c || !woEditable()) return;
  const s = woStatsOf(wo.rows);
  const ok = await askConfirm('Eksekusi data ' + c.id + '?',
    fmt(s.n) + ' data aktif dengan total PIUTANG YANG DI WO ' + rp(s.total) + (s.batal ? ' (' + fmt(s.batal) + ' data dibatalkan tidak ikut)' : '') + '. '
    + (s.noThc || s.noTak ? fmt(s.noThc) + ' data belum memiliki No. THC dan ' + fmt(s.noTak) + ' data belum memiliki No. TAK. ' : '')
    + 'Status menjadi "' + WO_ST_APPROVE + '" dan data hanya bisa dilihat, kecuali statusnya dikembalikan ke "' + WO_ST_CONFIRM + '".', 'Eksekusi Data');
  if (!ok) return;
  const btn = $('ewo-exec');
  btn.disabled = true;
  const r = await woCall('webWoExecute', { ewo: c.id }, { timeoutMs: 120000 });
  if (!r) return;
  if (!r.ok) { btn.disabled = false; return toast(r.error || 'Eksekusi Data gagal.', true); }
  wo.cur = r.head; wo.curAt = Date.now(); wo.awoAt = 0; wo.optAt = 0;
  paintEwoWork();
  toast(c.id + ' berstatus "' + WO_ST_APPROVE + '".');
  woLoadOptions();
}

/* ---------- Menu Approve Write-Off ---------- */

function woEnterApprove() {
  if (!state.user) return;
  if (!wo.awo || Date.now() - wo.awoAt > 30000) woLoadApprove();
  else paintApprove();
}

async function woLoadApprove() {
  if (wo.awoLoading) return;
  wo.awoLoading = true; $('awo-refresh').disabled = true;
  if (!wo.awo) setNote('awo-note', 'Memuat daftar ID_EWO…');
  const r = await woCall('webWoApproveList', {}, { retry: true });
  wo.awoLoading = false; $('awo-refresh').disabled = false;
  if (!r) return;
  if (!r.ok) return setNote('awo-note', r.error || 'Gagal memuat daftar Approve Write-Off.', 'bad');
  wo.awo = r.ewos || []; wo.awoAt = Date.now();
  paintApprove();
}

function paintApprove() {
  const all = wo.awo || [], j = wo.awoJ, jn = j === 'R' ? 'Reguler' : 'Khusus';
  const list = all.filter(e => !j || e.j === j).sort((a, b) => collator.compare(a.b, b.b) || (a.id < b.id ? 1 : a.id > b.id ? -1 : 0));
  setNote('awo-note', !all.length ? 'Belum ada ID_EWO berstatus "' + WO_ST_APPROVE + '" pada cabang Anda.'
    : !list.length ? 'Tidak ada ID_EWO WO ' + jn + ' berstatus "' + WO_ST_APPROVE + '". Ubah filter Jenis WO untuk melihat jenis lain.'
    : fmt(list.length) + (j ? ' ID_EWO WO ' + jn + ' (dari ' + fmt(all.length) + ' ID_EWO)' : ' ID_EWO') + ' berstatus "' + WO_ST_APPROVE + '" pada ' + fmt(new Set(list.map(e => e.b)).size) + ' cabang.');
  const th = (t, num) => h('th', { scope: 'col', class: num ? 'num' : '', text: t });
  const rows = list.map(e => {
    const done = e.ver === WO_VER_OK;
    return h('tr', null,
      h('td', { 'data-label': 'Cabang', text: e.b }), h('td', { class: 'wo-name', 'data-label': 'ID_EWO', text: e.id }), h('td', { 'data-label': 'Jenis', text: e.jenis }),
      h('td', { 'data-label': 'Tahun', text: String(e.y) }),
      h('td', { class: 'num', 'data-label': 'Data', text: fmt(e.n) }), h('td', { class: 'num', 'data-label': 'PIUTANG YANG DI WO', text: fmt(e.total) }),
      h('td', { 'data-label': 'Dieksekusi', text: e.exAt ? fmtStamp(e.exAt) + (e.exBy ? ' · ' + e.exBy : '') : '—' }),
      e.jv ? h('td', { 'data-label': 'No JV', text: e.jv }) : h('td', { class: 'is-dim', 'data-label': 'No JV', text: 'Belum diinput' }),
      done ? h('td', { 'data-label': 'Verifikasi Nilai WO' }, h('span', { class: 'wo-status is-ok', text: 'Sesuai' })) : h('td', { class: 'is-dim', 'data-label': 'Verifikasi Nilai WO', text: 'Belum diverifikasi' }),
      h('td', { class: 'wo-act' }, h('div', { class: 'wo-rowbtns' },
        woBtn('Cek Data', 'ghost', 'detail', { 'data-id': e.id }),
        woBtn('Input JV', 'ghost', 'jv', { 'data-id': e.id, disabled: done }),
        woBtn('Verifikasi Nilai WO', 'ghost', 'verif', { 'data-id': e.id, disabled: done }),
        done ? woBtn('Batalkan Verifikasi', 'warn', 'reopen', { 'data-id': e.id }) : null)));
  });
  $('awo-tbl').className = 'wo-hist wo-cards';
  put($('awo-tbl'), h('thead', null, h('tr', null, th('Cabang'), th('ID_EWO'), th('Jenis'), th('Tahun'), th('Data', true), th('PIUTANG YANG DI WO', true),
    th('Dieksekusi'), th('No JV'), th('Verifikasi Nilai WO'), th('Aksi'))),
  h('tbody', null, rows.length ? rows : h('tr', null, h('td', { colspan: '10', class: 'wo-empty', text: 'Tidak ada data.' }))));
}

function woApproveClick(e) {
  const b = e.target.closest('button[data-act]');
  if (!b || b.disabled) return;
  if (b.dataset.act === 'detail') woOpenDetail(b.dataset.id, b);
  else if (b.dataset.act === 'jv') woOpenJv(b.dataset.id);
  else if (b.dataset.act === 'verif') woOpenVerif(b.dataset.id);
  else if (b.dataset.act === 'reopen') woOpenReopen(b.dataset.id);
}
const woAwoOf = id => (wo.awo || []).find(e => e.id === id) || null;
function woApplyHead(head) {
  const i = (wo.awo || []).findIndex(e => e.id === head.id);
  if (i > -1) wo.awo[i] = head;
  wo.optAt = 0; wo.curAt = 0;                                // menu Eksekusi memuat ulang saat dibuka
}

/* ----- Cek Data: daftar data per ID_EWO (hanya baca) ----- */

// Kolom yang ditampilkan: [indeks kolom, jenis (n = angka)]. Total Balance sampai PIUTANG = nilai yang diverifikasi.
const WO_DET_COLS = [WC.NO, WC.CTR, WC.CLIENT, WC.NAME, WC.LOAN, WC.TB, WC.WAJIB, WC.SUKARELA, WC.PENSIUN, WC.HARIRAYA, WC.KHUSUS, WC.QURBAN, WC.SIPADAN,
  WC.TOTSIMP, WC.PIUTANG, WC.THC, WC.TAK, WC.JV];
const WO_DET_SUM = [WC.TB, WC.WAJIB, WC.SUKARELA, WC.PENSIUN, WC.HARIRAYA, WC.KHUSUS, WC.QURBAN, WC.SIPADAN, WC.TOTSIMP, WC.PIUTANG];

async function woOpenDetail(id, btn) {
  const e = woAwoOf(id);
  if (!e) return;
  if (btn) btn.disabled = true;
  setNote('awo-note', 'Memuat data ' + id + '…');
  const r = await woCall('webWoApproveRows', { ewo: id }, { retry: true, timeoutMs: 120000 });
  if (btn) btn.disabled = false;
  if (!r) return;
  if (!r.ok) { paintApprove(); return toast(r.error || 'Gagal memuat data ID_EWO.', true); }
  paintApprove();                                            // mengembalikan catatan daftar
  wo.det = { id, head: r.head, rows: r.rows || [], q: '', page: 0 };
  woApplyHead(r.head);                                       // ringkasan di daftar ikut segar (status bisa saja sudah berubah)
  if (r.head.st !== WO_ST_APPROVE) { wo.awo = wo.awo.filter(x => x.id !== id); paintApprove(); }
  paintDetail(true);
  $('wo-detail').showModal();
}

function woDetView() {
  const d = wo.det, q = d.q.trim().toLocaleLowerCase('id');
  if (!q) return d.rows;
  return d.rows.filter(r => [r[WC.CTR], r[WC.CLIENT], r[WC.NAME], r[WC.LOAN], r[WC.THC], r[WC.TAK]].some(v => String(v).toLocaleLowerCase('id').includes(q)));
}

function paintDetail(full) {
  const d = wo.det;
  if (!d) return;
  if (full) {
    const c = d.head, s = woStatsOf(d.rows), chip = (t, warn) => h('span', { class: 'st' + (warn ? ' warn' : ''), text: t });
    const search = h('input', { id: 'wd-q', type: 'search', placeholder: 'Cari Ctr ID, Client ID, nama, Loan No, No. THC, atau No. TAK…', autocomplete: 'off' });
    search.addEventListener('input', debounce(() => { d.q = search.value; d.page = 0; paintDetail(false); }, 150));
    put($('wd-body'),
      h('h2', { id: 'wd-title', text: 'Cek Data ' + c.id }),
      h('p', { class: 'dlg-sub', text: 'WO ' + c.jenis + ' · Tahun ' + c.y + ' · Cabang ' + c.b + ' · status ' + c.st + (c.exAt ? ' · dieksekusi ' + fmtStamp(c.exAt) : '') }),
      h('div', { class: 'stats' }, chip(fmt(s.n) + ' data aktif'), s.batal ? chip(fmt(s.batal) + ' dibatalkan', true) : null, chip('PIUTANG YANG DI WO ' + rp(s.total)),
        chip('Belum No. THC: ' + fmt(s.noThc), s.noThc > 0), chip('Belum No. TAK: ' + fmt(s.noTak), s.noTak > 0), chip(c.jv ? 'No JV: ' + c.jv : 'No JV belum diinput', !c.jv)),
      h('div', { class: 'wo-toolbar' }, h('label', { class: 'version-search', for: 'wd-q' }, h('span', { 'aria-hidden': 'true', text: '⌕' }), search)),
      h('div', { class: 'tw tall wo-tw' }, h('table', { id: 'wd-tbl' })),
      h('div', { class: 'wo-pager', id: 'wd-pager' }),
      h('div', { class: 'dlg-act' }, h('button', { class: 'btn ghost', type: 'button', onclick: () => $('wo-detail').close(), text: 'Tutup' })));
  }
  const view = woDetView(), pages = Math.max(1, Math.ceil(view.length / WO_PAGE));
  if (d.page >= pages) d.page = pages - 1;
  const slice = view.slice(d.page * WO_PAGE, (d.page + 1) * WO_PAGE);
  const th = c => h('th', { scope: 'col', class: WO_COLS[c][1] === 'n' && c !== WC.NO ? 'num' : (c === WC.NO ? 'num' : ''), text: WO_COLS[c][0] });
  const cell = (r, c) => {
    const v = r[c], num = WO_COLS[c][1] === 'n';
    if (c === WC.NAME) return h('td', { class: 'wo-name', 'data-label': WO_COLS[c][0] }, v, r[WC.STATUS] === WO_BATAL ? h('span', { class: 'wo-flag', text: 'BATAL' }) : null);
    if (num && c !== WC.NO) return v === '' ? h('td', { class: 'num is-dim', 'data-label': WO_COLS[c][0], text: '—' }) : h('td', { class: 'num', 'data-label': WO_COLS[c][0], text: fmt(v) });
    if (v === '' || v == null) return h('td', { class: (c === WC.NO ? 'num ' : '') + 'is-dim', 'data-label': WO_COLS[c][0], text: '—' });
    return h('td', { class: c === WC.NO ? 'num' : (c === WC.THC || c === WC.TAK ? 'wo-ref' : ''), 'data-label': WO_COLS[c][0], text: String(v) });
  };
  const rows = slice.map(r => h('tr', { class: r[WC.STATUS] === WO_BATAL ? 'is-batal' : '' }, WO_DET_COLS.map(c => cell(r, c))));
  const active = d.rows.filter(r => r[WC.STATUS] !== WO_BATAL), tot = {};
  WO_DET_SUM.forEach(c => { tot[c] = woNum(active.reduce((a, r) => a + (Number(r[c]) || 0), 0)); });
  const foot = h('tfoot', null, h('tr', null,
    h('td', { colspan: String(WO_DET_COLS.indexOf(WC.TB)), class: 'wo-tot-l', text: 'Total ' + fmt(active.length) + ' data aktif (tanpa yang dibatalkan)' }),
    WO_DET_COLS.slice(WO_DET_COLS.indexOf(WC.TB)).map(c => WO_DET_SUM.indexOf(c) > -1 ? h('td', { class: 'num', 'data-label': WO_COLS[c][0], text: fmt(tot[c]) }) : h('td'))));
  $('wd-tbl').className = 'wo-tbl wo-cards wo-det';
  put($('wd-tbl'), h('thead', null, h('tr', null, WO_DET_COLS.map(th))),
    h('tbody', null, rows.length ? rows : h('tr', null, h('td', { colspan: String(WO_DET_COLS.length), class: 'wo-empty', text: d.rows.length ? 'Tidak ada data yang cocok dengan pencarian.' : 'ID_EWO ini tidak memiliki data.' }))),
    d.rows.length ? foot : null);
  const total = view.length, from = total ? d.page * WO_PAGE + 1 : 0, to = Math.min(total, (d.page + 1) * WO_PAGE);
  const go = n => () => { d.page = Math.max(0, Math.min(pages - 1, d.page + n)); paintDetail(false); $('wd-tbl').scrollIntoView({ block: 'nearest' }); };
  put($('wd-pager'), h('span', { text: total ? 'Menampilkan ' + fmt(from) + '–' + fmt(to) + ' dari ' + fmt(total) + ' data' : '' }),
    pages > 1 ? h('div', { class: 'pnav' },
      h('button', { class: 'btn ghost sm', type: 'button', disabled: d.page === 0, onclick: go(-1), text: 'Sebelumnya' }),
      h('span', { text: 'Halaman ' + (d.page + 1) + ' / ' + pages }),
      h('button', { class: 'btn ghost sm', type: 'button', disabled: d.page >= pages - 1, onclick: go(1), text: 'Berikutnya' })) : null);
}

/* ----- Report Write-Off: seluruh data WO (termasuk Batal) + unduhan WO Sesuai Reguler / Khusus ----- */

const WO_REP_CHUNK = 5000, WO_REP_MAX_AGE = 120000;   // baris per permintaan; umur data sebelum dimuat ulang otomatis saat menu dibuka (ms)
const woRepFresh = () => ({ rows: [], at: 0, loaded: false, loading: false, dl: false, f: { j: '', st: '', y: '', b: '' }, q: '', page: 0 });
if (!wo.rep) wo.rep = woRepFresh();

// Status satu baris: Batal | Belum Diproses (belum masuk ID_EWO) | Confirm | Approve Write-Off | Sesuai (ID_EWO terverifikasi "Sesuai")
function woRepState(r, e) {
  if (r[WC.STATUS] === WO_BATAL) return WO_BATAL;
  if (!r[WC.EWO] || !e) return 'Belum Diproses';
  return e.ver === WO_VER_OK ? WO_VER_OK : e.st;
}
const woRepPill = st => h('span', { class: 'wo-status' + (st === WO_VER_OK ? ' is-ok' : st === WO_ST_APPROVE ? ' is-approve' : st === WO_BATAL ? ' is-bad' : st === 'Belum Diproses' ? ' is-muted' : ''), text: st });
const woJenisNama = code => code === 'R' ? 'Reguler' : 'Khusus';

function woRepReset() {
  wo.rep = woRepFresh();
  $('rwo-q').value = '';
  ['rwo-jenis', 'rwo-status'].forEach(id => { $(id).value = ''; });
  ['rwo-tahun', 'rwo-cabang'].forEach(id => $(id).replaceChildren(h('option', { value: '', text: 'Semua' })));
  ['rwo-stats', 'rwo-tbl', 'rwo-pager'].forEach(id => $(id).replaceChildren());
  $('rwo-refresh').disabled = false;
  woRepButtons(); setNote('rwo-note', '');
}

function woEnterReport() {
  if (!state.user) return;
  rsbEnter();
  const R = wo.rep;
  if (R.loading) return;
  if (!R.loaded || Date.now() - R.at > WO_REP_MAX_AGE) woLoadReport();
  else paintReport();
}

async function woLoadReport() {
  const R = wo.rep;
  if (R.loading) return;
  R.loading = true; $('rwo-refresh').disabled = true; woRepButtons();
  const items = [], eMap = Object.create(null), seen = Object.create(null);
  let failed = '';
  try {
    for (const code of ['R', 'K']) {
      let off = 0, total = Infinity;
      while (off < total) {
        setNote('rwo-note', 'Memuat data WO ' + woJenisNama(code) + '… ' + fmt(off) + (isFinite(total) ? ' dari ' + fmt(total) : '') + ' data');
        const r = await woCall('webWoReport', { j: code, off, lim: WO_REP_CHUNK }, { retry: true, timeoutMs: 120000 });
        if (!r) { R.loading = false; woRepButtons(); return; }                 // sesi berakhir / koneksi gagal (sudah diberitahukan) / tampilan sudah diganti
        if (!r.ok) throw new Error(r.error || 'Gagal memuat Report Write-Off.');
        total = r.total;
        (r.ewos || []).forEach(e => { eMap[e.id] = e; });
        r.rows.forEach(row => { const key = code + '|' + row[WC.LOAN]; if (!seen[key]) { seen[key] = true; items.push({ j: code, r: row }); } });
        if (!r.rows.length) break;
        off += r.rows.length;
      }
    }
  } catch (e) { failed = e.message || 'Gagal memuat Report Write-Off.'; }
  R.loading = false; $('rwo-refresh').disabled = false;
  if (wo.rep !== R) return;                                                   // logout / ganti akun selama memuat
  if (failed) { woRepButtons(); return setNote('rwo-note', failed + (R.loaded ? ' Data yang tampil adalah hasil pemuatan sebelumnya.' : ''), 'bad'); }
  items.forEach(x => {
    x.st = woRepState(x.r, eMap[x.r[WC.EWO]]);
    x.hay = [x.r[WC.EWO], x.r[WC.CTR], x.r[WC.CLIENT], x.r[WC.NAME], x.r[WC.LOAN], x.r[WC.THC], x.r[WC.TAK], x.r[WC.JV]].join('\u0001').toLocaleLowerCase('id');
  });
  items.sort((a, b) => (a.j === b.j ? 0 : a.j === 'R' ? -1 : 1) || (Number(b.r[WC.TAHUN]) || 0) - (Number(a.r[WC.TAHUN]) || 0)
    || collator.compare(a.r[WC.BRANCH], b.r[WC.BRANCH]) || (Number(a.r[WC.NO]) || 0) - (Number(b.r[WC.NO]) || 0));
  R.rows = items; R.at = Date.now(); R.loaded = true; R.page = 0;
  woRepFillFilters();
  paintReport();
}

function woRepFillFilters() {
  const R = wo.rep;
  const years = [...new Set(R.rows.map(x => x.r[WC.TAHUN]).filter(v => v !== ''))].sort((a, b) => b - a).map(String);
  const branches = [...new Set(R.rows.map(x => x.r[WC.BRANCH]).filter(Boolean))].sort(collator.compare);
  const fill = (id, vals, key) => {
    if (vals.indexOf(R.f[key]) < 0) R.f[key] = '';
    $(id).replaceChildren(h('option', { value: '', text: 'Semua' }), ...vals.map(v => h('option', { value: v, text: v })));
    $(id).value = R.f[key];
  };
  fill('rwo-tahun', years, 'y'); fill('rwo-cabang', branches, 'b');
}

function woRepView() {
  const R = wo.rep, f = R.f, q = R.q.trim().toLocaleLowerCase('id');
  return R.rows.filter(x => (!f.j || x.j === f.j) && (!f.st || x.st === f.st) && (!f.y || String(x.r[WC.TAHUN]) === f.y)
    && (!f.b || x.r[WC.BRANCH] === f.b) && (!q || x.hay.includes(q)));
}

function woRepButtons() {
  const R = wo.rep, n = j => R.loaded ? R.rows.filter(x => x.j === j && x.st === WO_VER_OK).length : 0, busy = R.loading || R.dl;
  [['rwo-dl-r', 'R', 'Unduh WO Sesuai Reguler'], ['rwo-dl-k', 'K', 'Unduh WO Sesuai Khusus']].forEach(([id, j, label]) => {
    $(id).textContent = label + (R.loaded ? ' (' + fmt(n(j)) + ' data)' : '');
    $(id).disabled = busy || !n(j);
  });
  $('rwo-dl-view').disabled = busy || !R.loaded || !woRepView().length;
}

function paintReport() {
  const R = wo.rep;
  if (!R.loaded) return;
  const view = woRepView(), pages = Math.max(1, Math.ceil(view.length / WO_PAGE));
  if (R.page >= pages) R.page = pages - 1;
  const active = view.filter(x => x.st !== WO_BATAL), chip = (t, warn) => h('span', { class: 'st' + (warn ? ' warn' : ''), text: t });
  const total = woNum(active.reduce((a, x) => a + (Number(x.r[WC.PIUTANG]) || 0), 0));
  put($('rwo-stats'), chip(fmt(view.length) + ' data'), chip(fmt(active.length) + ' aktif'), chip(fmt(view.length - active.length) + ' dibatalkan', view.length > active.length),
    chip(fmt(view.filter(x => x.st === WO_VER_OK).length) + ' Sesuai'), chip('PIUTANG YANG DI WO ' + rp(total)));
  const COLS = [['No.', x => x.r[WC.NO], 'num'], ['Jenis', x => woJenisNama(x.j)], ['Tahun', x => x.r[WC.TAHUN], 'plain'], ['Cabang', x => x.r[WC.BRANCH]],
    ['ID_EWO', x => x.r[WC.EWO]], ['Status', x => x.st, 'pill'], ['Ctr ID', x => x.r[WC.CTR]], ['Client ID', x => x.r[WC.CLIENT]], ['Client Name', x => x.r[WC.NAME], 'name'],
    ['Loan No', x => x.r[WC.LOAN]], ['Total Balance', x => x.r[WC.TB], 'num'], ['Total Simpanan', x => x.r[WC.TOTSIMP], 'num'], ['PIUTANG YANG DI WO', x => x.r[WC.PIUTANG], 'num'],
    ['No. THC', x => x.r[WC.THC], 'ref'], ['No. TAK', x => x.r[WC.TAK], 'ref'], ['No. JV', x => x.r[WC.JV], 'ref']];
  const cell = (x, [label, get, kind]) => {
    const v = get(x);
    if (kind === 'pill') return h('td', { 'data-label': label }, woRepPill(v));
    if (kind === 'num') return h('td', { class: 'num' + (v === '' ? ' is-dim' : ''), 'data-label': label, text: v === '' ? '—' : fmt(v) });
    if (kind === 'name') return h('td', { class: 'wo-name', 'data-label': label }, v, x.st === WO_BATAL ? h('span', { class: 'wo-flag', text: 'BATAL' }) : null);
    return h('td', { class: (kind === 'ref' ? 'wo-ref ' : '') + (v === '' ? 'is-dim' : ''), 'data-label': label, text: v === '' || v == null ? '—' : String(v) });
  };
  const slice = view.slice(R.page * WO_PAGE, (R.page + 1) * WO_PAGE);
  $('rwo-tbl').className = 'wo-tbl wo-cards rwo-det';
  put($('rwo-tbl'), h('thead', null, h('tr', null, COLS.map(c => h('th', { scope: 'col', class: c[2] === 'num' ? 'num' : '', text: c[0] })))),
    h('tbody', null, slice.length ? slice.map(x => h('tr', { class: x.st === WO_BATAL ? 'is-batal' : '' }, COLS.map(c => cell(x, c))))
      : h('tr', null, h('td', { colspan: String(COLS.length), class: 'wo-empty', text: R.rows.length ? 'Tidak ada data yang cocok dengan filter atau pencarian.' : 'Belum ada data WO pada cabang Anda.' }))));
  const from = view.length ? R.page * WO_PAGE + 1 : 0, to = Math.min(view.length, (R.page + 1) * WO_PAGE);
  const go = n => () => { R.page = Math.max(0, Math.min(pages - 1, R.page + n)); paintReport(); $('rwo-tbl').scrollIntoView({ block: 'nearest' }); };
  put($('rwo-pager'), h('span', { text: view.length ? 'Menampilkan ' + fmt(from) + '–' + fmt(to) + ' dari ' + fmt(view.length) + ' data' : '' }),
    pages > 1 ? h('div', { class: 'pnav' },
      h('button', { class: 'btn ghost sm', type: 'button', disabled: R.page === 0, onclick: go(-1), text: 'Sebelumnya' }),
      h('span', { text: 'Halaman ' + (R.page + 1) + ' / ' + pages }),
      h('button', { class: 'btn ghost sm', type: 'button', disabled: R.page >= pages - 1, onclick: go(1), text: 'Berikutnya' })) : null);
  if (!R.loading) {
    const t = new Date(R.at);
    setNote('rwo-note', 'Data dimuat pukul ' + String(t.getHours()).padStart(2, '0') + '.' + String(t.getMinutes()).padStart(2, '0') + ' (' + fmt(R.rows.length) + ' data WO pada cabang Anda). Tekan Muat ulang untuk memperbarui.');
  }
  woRepButtons();
}

// kind: 'R' / 'K' = WO Sesuai jenis itu (seluruh data aktif, tidak terpengaruh filter); 'view' = data yang sedang tampil (termasuk Batal)
async function woRepDownload(kind) {
  const R = wo.rep;
  if (!R.loaded || R.dl) return;
  const sesuai = kind !== 'view';
  const items = sesuai ? R.rows.filter(x => x.j === kind && x.st === WO_VER_OK) : woRepView();
  if (!items.length) return toast(sesuai ? 'Belum ada WO Sesuai ' + woJenisNama(kind) + '.' : 'Tidak ada data untuk diunduh.', true);
  R.dl = true; woRepButtons(); setNote('rwo-note', 'Menyiapkan file Excel…');
  try {
    await Promise.all([loadXLSX(), nextFrame()]);
    // [judul, jenis (n = angka, z = format Excel), pengambil nilai]; sesuai = 25 kolom dokumen (sampai ID_EWO) + Alasan Batal, tampilan = + Jenis WO & Status
    const cols = (sesuai ? [] : [['Jenis WO', 't', x => woJenisNama(x.j)], ['Status', 't', x => x.st]])
      .concat(WO_COLS.slice(0, WO_NBASE + 1).map((c, k) => [c[0], c[1] === 'n' ? (k === WC.NO || k === WC.TAHUN ? 'i' : 'n') : 't', x => x.r[k], k]))
      .concat([['Alasan Batal', 't', x => x.r[WC.ALASAN_BATAL] || '', WC.ALASAN_BATAL]]);   // kolom terakhir, baik Unduh WO Sesuai maupun Unduh Data Yang Tampil
    const aoa = [cols.map(c => c[0])].concat(items.map(x => cols.map(c => { const v = c[2](x); return c[1] === 't' ? String(v) : (v === '' ? '' : Number(v)); })));
    const ws = XLSX.utils.aoa_to_sheet(aoa), L = k => XLSX.utils.encode_col(k);
    const isRef = c => c[3] === WC.THC || c[3] === WC.TAK || c[3] === WC.JV;
    const fill = XS.text && { patternType: 'solid', fgColor: { rgb: 'FFF4BD' } };
    cols.forEach((c, k) => {
      ws[L(k) + '1'].s = XS.head;
      for (let i = 0; i < items.length; i++) {
        const cell = ws[L(k) + (i + 2)];
        if (!cell) continue;
        if (c[1] === 'i') { cell.s = XS.num; cell.z = '0'; }
        else if (c[1] === 'n') { cell.s = XS.num; cell.z = XL_NUM; }
        else { cell.s = isRef(c) ? Object.assign({}, XS.text, { fill }) : XS.text; if (isRef(c)) cell.z = '@'; }
      }
    });
    ws['!cols'] = cols.map((c, k) => ({ wch: Math.min(45, Math.max(10, c[0].length + 3, ...items.slice(0, 300).map(x => String(c[2](x)).length + 2))) }));
    ws['!rows'] = [{ hpt: 32 }];
    ws['!freeze'] = { ySplit: '1', topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' };
    ws['!autofilter'] = { ref: 'A1:' + L(cols.length - 1) + (items.length + 1) };
    const d = new Date(), ymd = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    const name = sesuai ? 'Report WO Sesuai ' + woJenisNama(kind) + ' ' + ymd + '.xlsx' : 'Report WO ' + ymd + '.xlsx';
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sesuai ? 'WO ' + woJenisNama(kind) + ' Sesuai' : 'Report WO');
    XLSX.writeFile(wb, name);
    const batal = items.filter(x => x.st === WO_BATAL).length;
    setNote('rwo-note', 'File Excel "' + name + '" berhasil dibuat: ' + fmt(items.length) + ' data' + (batal ? ' (termasuk ' + fmt(batal) + ' dibatalkan)' : '') + '.', 'good');
  } catch (e) { setNote('rwo-note', e.message || 'Gagal membuat file Excel.', 'bad'); }
  R.dl = false; woRepButtons();
}

function woBindReport() {
  $('rwo-refresh').onclick = () => { woLoadReport(); rsbLoad(); };
  $('rsb-tbl').addEventListener('click', rsbClick);
  $('rsb-q').addEventListener('input', debounce(() => { wo.rsb.q = $('rsb-q').value; rsbPaint(); }, 150));
  $('wt-form').addEventListener('submit', rsbSaveTtd);
  $('wo-ttd').addEventListener('close', () => { if (wo.rsb) wo.rsb.ttd = null; });
  $('rwo-dl-r').onclick = () => woRepDownload('R');
  $('rwo-dl-k').onclick = () => woRepDownload('K');
  $('rwo-dl-view').onclick = () => woRepDownload('view');
  ['rwo-jenis', 'rwo-status', 'rwo-tahun', 'rwo-cabang'].forEach(id => $(id).addEventListener('change', () => {
    wo.rep.f = { j: $('rwo-jenis').value, st: $('rwo-status').value, y: $('rwo-tahun').value, b: $('rwo-cabang').value };
    wo.rep.page = 0; paintReport();
  }));
  $('rwo-q').addEventListener('input', debounce(() => { wo.rep.q = $('rwo-q').value; wo.rep.page = 0; paintReport(); }, 150));
}

/* ----- Surat Berita Acara (menu Report Write-Off) -----
   Daftar cabang yang punya ID_EWO "Sesuai": TTD MSA / TTD BM (unggah gambar), Generate (susun surat PDF di browser), dan tombol unduh per Jenis WO.
   Data surat, nama penanda tangan, dan gambar TTD datang dari server (webWoSuratList / webWoSuratTtd / webWoSuratData di WriteOff.gs);
   PDF disusun oleh surat-wo.js memakai surat-kop.js dan surat-font.js (dimuat sesuai kebutuhan, satu folder dengan index.html). */

const RSB_ROLES = { MSA: 'MIS Support Administration Staff', BM: 'Branch Manager' };
const RSB_SCRIPTS = ['surat-kop.js', 'surat-font.js', 'surat-wo.js'];
const RSB_TTD_MAX = 42000;                                    // panjang data URL TTD; server menerima sampai 45.000 karakter
const rsbFresh = () => ({ items: null, at: 0, loading: false, q: '', busy: {}, cache: {}, ttd: null });
if (!wo.rsb) wo.rsb = rsbFresh();

function rsbReset() {
  wo.rsb = rsbFresh();
  $('rsb-q').value = '';
  $('rsb-tbl').replaceChildren(); $('rsb-count').textContent = '';
  setNote('rsb-note', '');
}

let suratLoading = null;
function loadSuratWO() {
  if (window.SuratWO && window.SURAT_KOP && window.SURAT_FONTS) return Promise.resolve();
  return suratLoading || (suratLoading = Promise.all(RSB_SCRIPTS.map(src => new Promise((ok, no) => {
    const s = h('script', { src });
    s.onload = ok; s.onerror = () => no(new Error(src));
    document.head.append(s);
  }))).then(() => {
    if (!window.SuratWO || !window.SURAT_KOP || !window.SURAT_FONTS) throw new Error('isi berkas tidak sesuai');
  }).catch(e => {
    suratLoading = null;
    throw new Error('Berkas pembuat surat gagal dimuat (' + (e.message || 'tidak diketahui') + '). Pastikan ' + RSB_SCRIPTS.join(', ') + ' berada satu folder dengan index.html, lalu muat ulang halaman.');
  }));
}

const rsbEwo = id => { for (const br of (wo.rsb.items || [])) { const e = br.ewos.find(x => x.id === id); if (e) return e; } return null; };
const rsbBranchOf = id => (wo.rsb.items || []).find(br => br.ewos.some(x => x.id === id)) || null;
// Satu cabang boleh punya beberapa ID_EWO untuk jenis + tahun yang sama (ID_EWO berikutnya setelah yang sebelumnya "Sesuai").
// Bila begitu, label tombol diberi nomor urut ID_EWO dan nama file diberi ID_EWO-nya agar tidak kembar; bila tidak, tampilan tetap seperti semula.
const rsbDup = (br, e) => !!br && br.ewos.filter(x => x.j === e.j && x.y === e.y).length > 1;
const rsbLabel = (br, e) => e.jenis + ' ' + e.y + (rsbDup(br, e) ? ' #' + (Number(String(e.id).slice(-9)) || e.id) : '');
const rsbFileName = (L, tag) => ('Berita Acara WO ' + L.jenis + ' Tahun ' + L.y + ' Cabang ' + L.b + (tag ? ' ' + tag : '')).replace(/[\\/:*?"<>|\u0000-\u001F]+/g, ' ').replace(/\s+/g, ' ').trim() + '.pdf';
const rsbNowStamp = () => { const d = new Date(); return today() + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds()); };
function rsbSave(blob, name) {
  const url = URL.createObjectURL(blob), a = h('a', { href: url, download: name });
  document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}

function rsbEnter() {
  const R = wo.rsb;
  if (R.loading) return;
  if (!R.items || Date.now() - R.at > WO_REP_MAX_AGE) rsbLoad();
  else rsbPaint();
}

async function rsbLoad(quiet) {
  const R = wo.rsb;
  if (R.loading) return;
  R.loading = true;
  if (!quiet && !R.items) setNote('rsb-note', 'Memuat daftar Surat Berita Acara…');
  const r = await woCall('webWoSuratList', {}, { retry: true, timeoutMs: 120000 });
  if (wo.rsb !== R) return;                                    // logout / ganti akun selama memuat
  R.loading = false;
  if (!r) return;
  if (!r.ok) return setNote('rsb-note', (r.error || 'Gagal memuat daftar Surat Berita Acara.') + (R.items ? ' Data yang tampil adalah hasil pemuatan sebelumnya.' : ''), 'bad');
  R.items = r.branches || []; R.at = Date.now();
  if (!quiet || /^Memuat/.test($('rsb-note').textContent)) setNote('rsb-note', '');
  rsbPaint();
}

const rsbStaff = (key, p) => h('span', { class: p.nama ? '' : 'wo-miss', text: key + ': ' + (p.nama || 'belum ditemukan') });
function rsbTtdBtn(br, role) {
  const p = role === 'MSA' ? br.msa : br.bm;
  return h('button', { class: 'btn sm ghost rsb-ttd' + (p.ttd ? ' is-set' : ''), type: 'button', 'data-act': 'ttd', 'data-b': br.b, 'data-role': role,
    title: RSB_ROLES[role] + (p.nama ? ' · ' + p.nama : '') + (p.ttd ? ' · TTD diunggah' + (p.at ? ' ' + fmtStamp(p.at) : '') + (p.by ? ' oleh ' + p.by : '') : ' · belum ada TTD'),
    text: 'TTD ' + role + (p.ttd ? ' ✓' : '') });
}
function rsbJenisCell(br) {
  const R = wo.rsb, shown = br.ewos.filter(e => e.state !== 'none');
  const surat = !shown.length ? h('span', { class: 'rsb-none', text: 'Belum digenerate' }) : h('div', { class: 'rsb-dls' }, shown.map(e => {
    const label = rsbLabel(br, e), info = e.id + ' · ' + fmt(e.n) + ' data diproses, ' + fmt(e.nb) + ' dibatalkan · digenerate ' + fmtStamp(e.genAt) + (e.genBy ? ' oleh ' + e.genBy : '');
    if (e.state === 'ready') return h('button', { class: 'btn sm ghost', type: 'button', 'data-act': 'dl', 'data-id': e.id, disabled: !!R.busy[e.id], title: info, text: '↓ ' + label });
    return h('span', { class: 'wo-status is-muted rsb-stale', title: info + '. Data atau TTD berubah sejak digenerate.', text: label + ' · perlu generate ulang' });
  }));
  // Jurnal: satu tombol per tahun (WO Reguler + WO Khusus cabang itu), tidak menunggu surat digenerate
  const years = [...new Set(br.ewos.map(e => e.y))].sort((a, b) => b - a);
  const jurnal = h('div', { class: 'rsb-dls rsb-jr' }, years.map(y => h('button', {
    class: 'btn sm ghost rsb-jbtn', type: 'button', 'data-act': 'jurnal', 'data-b': br.b, 'data-y': String(y), disabled: !!R.busy[rsbJKey(br.b, y)],
    title: 'Unduh Jurnal WO Reguler dan WO Khusus tahun ' + y + ' (ZIP, satu PDF per nomor transaksi)', text: '↓ Jurnal ' + y })));
  return h('div', { class: 'rsb-cellstack' }, surat, jurnal);
}

function rsbPaint() {
  const R = wo.rsb, tbl = $('rsb-tbl');
  if (!R.items) { tbl.replaceChildren(); $('rsb-count').textContent = ''; return; }
  const q = R.q.trim().toLocaleUpperCase('id');
  const list = R.items.filter(x => !q || x.b.toLocaleUpperCase('id').includes(q));
  $('rsb-count').textContent = fmt(list.length) + (list.length === R.items.length ? '' : ' dari ' + fmt(R.items.length)) + ' cabang';
  tbl.className = 'wo-cards';
  const th = t => h('th', { scope: 'col', text: t });
  if (!list.length) {
    return put(tbl, h('tbody', null, h('tr', null, h('td', { class: 'wo-empty', colspan: '3',
      text: R.items.length ? 'Tidak ada cabang yang cocok dengan pencarian.' : 'Belum ada cabang dengan ID_EWO yang verifikasinya “' + WO_VER_OK + '”.' }))));
  }
  put(tbl, h('thead', null, h('tr', null, th('Nama Cabang'), th('Aksi'), th('Jenis WO'))),
    h('tbody', null, list.map(br => {
      const busy = !!R.busy[br.b];
      return h('tr', null,
        h('td', { class: 'wo-name', 'data-label': 'Nama Cabang' }, h('b', { text: br.b }), h('div', { class: 'rsb-staff' }, rsbStaff('MSA', br.msa), ' · ', rsbStaff('BM', br.bm))),
        h('td', { class: 'wo-act', 'data-label': 'Aksi' }, h('div', { class: 'wo-rowbtns' }, rsbTtdBtn(br, 'MSA'), rsbTtdBtn(br, 'BM'),
          h('button', { class: 'btn sm', type: 'button', 'data-act': 'gen', 'data-b': br.b, disabled: busy, text: busy ? 'Membuat…' : 'Generate' }))),
        h('td', { 'data-label': 'Jenis WO' }, rsbJenisCell(br)));
    })));
}

function rsbClick(ev) {
  const b = ev.target.closest('button[data-act]');
  if (!b || b.disabled) return;
  if (b.dataset.act === 'ttd') rsbOpenTtd(b.dataset.b, b.dataset.role);
  else if (b.dataset.act === 'gen') rsbGenerate(b.dataset.b);
  else if (b.dataset.act === 'dl') rsbDownload(b.dataset.id);
  else if (b.dataset.act === 'jurnal') rsbJurnal(b.dataset.b, Number(b.dataset.y));
}

/* ----- TTD: unggah gambar -> dirapikan di browser (latar putih bersih, dipotong sebatas goresan, dikecilkan, JPEG) -> disimpan server ----- */

async function rsbLoadBitmap(file) {
  if (window.createImageBitmap) { try { return await createImageBitmap(file); } catch (e) { /* lanjut ke elemen gambar */ } }
  const url = URL.createObjectURL(file);
  try {
    return await new Promise((ok, no) => {
      const im = new Image();
      im.onload = () => ok(im); im.onerror = () => no(new Error('Gambar tidak dapat dibaca. Gunakan berkas JPG atau PNG.'));
      im.src = url;
    });
  } finally { setTimeout(() => URL.revokeObjectURL(url), 10000); }
}

async function rsbPrepareSig(file) {
  if (!/^image\//.test(file.type)) throw new Error('Berkas harus berupa gambar (JPG atau PNG).');
  if (file.size > 10 * 1024 * 1024) throw new Error('Ukuran gambar maksimal 10 MB.');
  const src = await rsbLoadBitmap(file), w0 = src.width, h0 = src.height;
  if (!w0 || !h0) throw new Error('Gambar tidak dapat dibaca.');
  const k0 = Math.min(1, 1400 / Math.max(w0, h0)), cw = Math.max(1, Math.round(w0 * k0)), ch = Math.max(1, Math.round(h0 * k0));
  const c = document.createElement('canvas'); c.width = cw; c.height = ch;
  const g = c.getContext('2d', { willReadFrequently: true });
  g.fillStyle = '#fff'; g.fillRect(0, 0, cw, ch); g.drawImage(src, 0, 0, cw, ch);   // latar transparan menjadi putih
  const img = g.getImageData(0, 0, cw, ch), px = img.data;
  let x0 = cw, y0 = ch, x1 = -1, y1 = -1;
  for (let y = 0; y < ch; y++) for (let x = 0; x < cw; x++) {
    const i = (y * cw + x) * 4, lum = 0.299 * px[i] + 0.587 * px[i + 1] + 0.114 * px[i + 2];
    if (lum > 215) { px[i] = px[i + 1] = px[i + 2] = 255; continue; }                // warna kertas -> putih bersih
    if (lum < 170) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
  }
  if (x1 < 0) throw new Error('Goresan tanda tangan tidak ditemukan. Gunakan foto atau pindaian dengan tinta yang jelas di atas kertas putih.');
  g.putImageData(img, 0, 0);
  const mx = Math.round(Math.max(x1 - x0, y1 - y0) * 0.04) + 2;
  x0 = Math.max(0, x0 - mx); y0 = Math.max(0, y0 - mx); x1 = Math.min(cw - 1, x1 + mx); y1 = Math.min(ch - 1, y1 + mx);
  const sw = x1 - x0 + 1, sh = y1 - y0 + 1;
  let k = Math.min(1, 520 / sw, 220 / sh);
  for (let attempt = 0; attempt < 5; attempt++, k *= 0.8) {
    const ow = Math.max(8, Math.round(sw * k)), oh = Math.max(8, Math.round(sh * k));
    const o = document.createElement('canvas'); o.width = ow; o.height = oh;
    const og = o.getContext('2d');
    og.fillStyle = '#fff'; og.fillRect(0, 0, ow, oh); og.drawImage(c, x0, y0, sw, sh, 0, 0, ow, oh);
    for (const q of [0.88, 0.78, 0.66, 0.55]) {
      const url = o.toDataURL('image/jpeg', q);
      if (/^data:image\/jpeg;base64,/.test(url) && url.length <= RSB_TTD_MAX) return url;
    }
  }
  throw new Error('Gambar tidak dapat dikecilkan hingga batas ukuran. Gunakan gambar yang lebih sederhana.');
}

function rsbOpenTtd(branch, role) {
  const R = wo.rsb, br = R.items && R.items.find(x => x.b === branch), who = br && (role === 'MSA' ? br.msa : role === 'BM' ? br.bm : null);
  if (!who) return;
  R.ttd = { b: branch, role, img: '', busy: false };
  put($('wt-form'),
    h('h2', { id: 'wt-title', text: 'TTD ' + role + ' · ' + RSB_ROLES[role] }),
    h('p', { class: 'dlg-sub', text: 'Cabang ' + branch }),
    h('p', { class: 'wm-hint', text: 'Nama pada surat: ' + (who.nama || 'belum ditemukan (periksa kolom CABANG dan JABATAN di Update_Karyawan)') + '.' }),
    h('p', { class: 'wm-hint', text: who.ttd ? 'TTD sudah diunggah' + (who.at ? ' ' + fmtStamp(who.at) : '') + (who.by ? ' oleh ' + who.by : '') + '. Pilih gambar baru untuk menggantinya.' : 'Belum ada TTD yang diunggah.' }),
    h('div', { class: 'wt-pick' }, h('label', { for: 'wt-file', text: 'Gambar tanda tangan (JPG / PNG)' }),
      h('input', { id: 'wt-file', type: 'file', accept: 'image/png,image/jpeg,image/webp' })),
    h('div', { class: 'wt-prev', id: 'wt-prev', text: 'Pratinjau tanda tangan akan tampil di sini.' }),
    h('p', { class: 'wm-err', id: 'wt-err', role: 'alert' }),
    h('div', { class: 'dlg-act' },
      who.ttd ? h('button', { class: 'btn danger wt-act-l', type: 'button', id: 'wt-del', text: 'Hapus TTD' }) : null,
      h('button', { class: 'btn ghost', type: 'button', id: 'wt-cancel', text: 'Batal' }),
      h('button', { class: 'btn fit', type: 'submit', id: 'wt-save', disabled: true, text: 'Simpan' })));
  $('wt-file').addEventListener('change', rsbPickFile);
  $('wt-cancel').onclick = () => $('wo-ttd').close();
  if (who.ttd) $('wt-del').onclick = rsbDeleteTtd;
  $('wo-ttd').showModal();
}

async function rsbPickFile() {
  const T = wo.rsb.ttd, f = $('wt-file').files[0];
  if (!T || !f) return;
  const err = $('wt-err'), prev = $('wt-prev'), save = $('wt-save'), empty = 'Pratinjau tanda tangan akan tampil di sini.';
  err.textContent = ''; T.img = ''; save.disabled = true; put(prev, 'Memproses gambar…');
  try {
    const url = await rsbPrepareSig(f);
    if (wo.rsb.ttd !== T) return;
    T.img = url; save.disabled = false;
    put(prev, h('img', { src: url, alt: 'Pratinjau tanda tangan' }));
  } catch (e) {
    if (wo.rsb.ttd !== T) return;
    put(prev, empty); err.textContent = e.message || 'Gambar tidak dapat diproses.';
  }
}

// Perbarui tampilan seketika; surat yang sudah digenerate langsung ditandai "perlu generate ulang", lalu muat ulang dari server untuk nilai pastinya.
function rsbApplyTtd(branch, role, on, r) {
  const R = wo.rsb, br = R.items && R.items.find(x => x.b === branch);
  if (br) {
    const p = role === 'MSA' ? br.msa : br.bm;
    p.ttd = on; p.by = on ? (r.by || '') : ''; p.at = on ? (r.at || '') : '';
    br.ewos.forEach(e => { if (e.state === 'ready') e.state = 'stale'; });
  }
  rsbPaint(); rsbLoad(true);
}

async function rsbSaveTtd(ev) {
  ev.preventDefault();
  const R = wo.rsb, T = R.ttd, err = $('wt-err'), save = $('wt-save');
  if (!T || T.busy) return;
  if (!T.img) { err.textContent = 'Pilih gambar tanda tangan terlebih dahulu.'; return; }
  T.busy = true; save.disabled = true; err.textContent = '';
  const r = await woCall('webWoSuratTtd', { b: T.b, role: T.role, img: T.img }, { timeoutMs: 60000 });
  if (wo.rsb !== R) return;
  T.busy = false;
  if (!r) return;
  if (!r.ok) { if (R.ttd === T) { err.textContent = r.error || 'TTD gagal disimpan.'; save.disabled = false; } else toast(r.error || 'TTD gagal disimpan.', true); return; }
  if (R.ttd === T) $('wo-ttd').close();
  rsbApplyTtd(T.b, T.role, true, r);
  toast('TTD ' + T.role + ' cabang ' + T.b + ' tersimpan.');
}

async function rsbDeleteTtd() {
  const R = wo.rsb, T = R.ttd;
  if (!T || T.busy) return;
  const ok = await askConfirm('Hapus TTD ' + T.role + '?', 'TTD ' + RSB_ROLES[T.role] + ' cabang ' + T.b + ' dihapus. Surat yang sudah digenerate perlu digenerate ulang.', 'Hapus');
  if (!ok || wo.rsb !== R || R.ttd !== T) return;
  T.busy = true; $('wt-del').disabled = true; $('wt-save').disabled = true; $('wt-err').textContent = '';
  const r = await woCall('webWoSuratTtd', { b: T.b, role: T.role, hapus: '1' }, { timeoutMs: 60000 });
  if (wo.rsb !== R) return;
  T.busy = false;
  if (!r) return;
  if (!r.ok) { if (R.ttd === T) { $('wt-err').textContent = r.error || 'TTD gagal dihapus.'; $('wt-del').disabled = false; } else toast(r.error || 'TTD gagal dihapus.', true); return; }
  if (R.ttd === T) $('wo-ttd').close();
  rsbApplyTtd(T.b, T.role, false, r);
  toast('TTD ' + T.role + ' cabang ' + T.b + ' dihapus.');
}

/* ----- Generate & unduh ----- */

async function rsbGenerate(branch) {
  const R = wo.rsb, br = R.items && R.items.find(x => x.b === branch);
  if (!br || R.busy[branch]) return;
  const noName = [br.msa.nama ? '' : RSB_ROLES.MSA, br.bm.nama ? '' : RSB_ROLES.BM].filter(Boolean);
  if (noName.length) return setNote('rsb-note', 'Nama staf ' + noName.join(' dan ') + ' cabang ' + branch + ' belum ditemukan di data karyawan (Update_Karyawan atau Akses_2). Lengkapi kolom CABANG dan JABATAN-nya terlebih dahulu.', 'bad');
  const noTtd = [br.msa.ttd ? '' : 'TTD MSA', br.bm.ttd ? '' : 'TTD BM'].filter(Boolean);
  if (noTtd.length && !(await askConfirm('Generate tanpa ' + noTtd.join(' dan ') + '?',
    'Surat tetap dibuat, tetapi ruang tanda tangan yang belum diunggah dikosongkan. Unggah TTD lalu Generate ulang untuk menempelkannya.', 'Tetap Generate'))) return;
  if (wo.rsb !== R || R.busy[branch]) return;
  R.busy[branch] = true; rsbPaint();
  const made = [];
  let failed = '';
  try {
    await loadSuratWO();
    for (const e of br.ewos) {
      const label = rsbLabel(br, e);
      setNote('rsb-note', 'Membuat Surat Berita Acara ' + label + ' cabang ' + branch + '…');
      const r = await woCall('webWoSuratData', { ewo: e.id, gen: '1' }, { timeoutMs: 120000 });
      if (wo.rsb !== R || !r) return;
      if (!r.ok) throw new Error(r.error || 'Gagal membuat surat ' + label + '.');
      const pdf = await SuratWO.build(r.letter);
      R.cache[e.id] = { fp: r.fp, blob: pdf.blob(), name: rsbFileName(r.letter, rsbDup(br, e) ? e.id : '') };
      Object.assign(e, { state: 'ready', fp: r.fp, genAt: rsbNowStamp(), genBy: state.user.nik || '' });
      made.push(label);
      rsbPaint();
    }
  } catch (e) { failed = e.message || 'Gagal membuat surat.'; }
  if (wo.rsb !== R) return;
  R.busy[branch] = false; rsbPaint();
  if (failed) setNote('rsb-note', (made.length ? 'Surat ' + made.join(', ') + ' berhasil dibuat, tetapi proses terhenti: ' : '') + failed, 'bad');
  else { setNote('rsb-note', 'Surat Berita Acara cabang ' + branch + ' dibuat: ' + made.join(', ') + '. Unduh lewat tombol pada kolom Jenis WO.', 'good'); toast('Surat Berita Acara ' + branch + ' dibuat.'); }
  rsbLoad(true);
}

async function rsbDownload(id) {
  const R = wo.rsb, e = rsbEwo(id);
  if (!e || R.busy[id]) return;
  const c = R.cache[id];
  if (c && c.fp === e.fp) { rsbSave(c.blob, c.name); setNote('rsb-note', 'Surat "' + c.name + '" diunduh.', 'good'); return; }
  R.busy[id] = true; rsbPaint();
  try {
    const br = rsbBranchOf(id);
    setNote('rsb-note', 'Menyiapkan surat ' + rsbLabel(br, e) + '…');
    await loadSuratWO();
    const r = await woCall('webWoSuratData', { ewo: id, gen: '0' }, { retry: true, timeoutMs: 120000 });
    if (wo.rsb !== R || !r) return;
    if (!r.ok) { rsbLoad(true); throw new Error(r.error || 'Surat tidak dapat disiapkan.'); }
    const pdf = await SuratWO.build(r.letter);
    R.cache[id] = { fp: r.fp, blob: pdf.blob(), name: rsbFileName(r.letter, rsbDup(br, e) ? id : '') };
    rsbSave(R.cache[id].blob, R.cache[id].name);
    setNote('rsb-note', 'Surat "' + R.cache[id].name + '" diunduh.', 'good');
  } catch (err) { setNote('rsb-note', err.message || 'Gagal menyiapkan surat.', 'bad'); }
  finally { if (wo.rsb === R) { R.busy[id] = false; rsbPaint(); } }
}

/* ----- Jurnal Write-Off: satu ZIP "Jurnal [Tahun]" per cabang (WO Reguler + WO Khusus), satu PDF Journal Voucher per nomor transaksi -----
   Nomor transaksi: WO Khusus = No JV, WO Reguler = No. TAK. Isi jurnal disusun server (webWoJurnalData); PDF dan ZIP disusun di browser (surat-wo.js). */

const rsbJKey = (b, y) => 'J|' + b + '|' + y;
// Nama berkas PDF = nomor transaksi. Karakter yang dilarang di nama berkas diganti "-"; nama kembar diberi (2), (3), ...
function rsbJurnalFileName(no, used) {
  let base = String(no == null ? '' : no).replace(/[\\/:*?"<>|\u0000-\u001F]+/g, '-').replace(/\s+/g, ' ').trim().replace(/[. ]+$/, '') || 'Jurnal';
  base = base.slice(0, 120);
  let name = base, n = 1;
  while (used[name.toLowerCase()]) name = base + ' (' + (++n) + ')';
  used[name.toLowerCase()] = true;
  return name + '.pdf';
}

async function rsbJurnal(branch, year) {
  const R = wo.rsb, key = rsbJKey(branch, year);
  if (!R || R.busy[key] || !Number.isFinite(year)) return;
  R.busy[key] = true; rsbPaint();
  try {
    setNote('rsb-note', 'Menyiapkan Jurnal ' + year + ' cabang ' + branch + '…');
    await loadSuratWO();
    const r = await woCall('webWoJurnalData', { b: branch, y: String(year) }, { retry: true, timeoutMs: 120000 });
    if (wo.rsb !== R || !r) return;
    if (!r.ok) throw new Error(r.error || 'Jurnal tidak dapat disiapkan.');
    const files = [], used = Object.create(null), list = r.vouchers || [];
    if (!list.length) throw new Error('Tidak ada jurnal untuk cabang ' + branch + ' tahun ' + year + '.');
    for (let i = 0; i < list.length; i++) {
      setNote('rsb-note', 'Membuat PDF jurnal ' + fmt(i + 1) + ' dari ' + fmt(list.length) + ' (' + list[i].no + ')…');
      const pdf = await SuratWO.buildJurnal(list[i]);
      files.push({ name: rsbJurnalFileName(list[i].no, used), bytes: pdf.bytes });
      if (i % 5 === 4) await new Promise(ok => setTimeout(ok, 0));          // beri kesempatan layar menyegarkan teks progres
      if (wo.rsb !== R) return;
    }
    const cab = String(r.cab || '').replace(/[\\/:*?"<>|\u0000-\u001F]+/g, ' ').replace(/\s+/g, ' ').trim();       // sudah Proper Case dari server
    const name = 'Jurnal ' + year + (cab ? ' ' + cab : '') + '.zip';      // contoh: Jurnal 2026 Cikalong Kulon.zip
    rsbSave(SuratWO.zip(files), name);
    setNote('rsb-note', '"' + name + '" diunduh: ' + fmt(files.length) + ' PDF (' + fmt(list.filter(v => v.jenis === 'Reguler').length) + ' nomor TAK Reguler, '
      + fmt(list.filter(v => v.jenis === 'Khusus').length) + ' nomor JV Khusus) cabang ' + r.b + '.', 'good');
  } catch (err) { setNote('rsb-note', err.message || 'Gagal membuat jurnal.', 'bad'); }
  finally { if (wo.rsb === R) { R.busy[key] = false; rsbPaint(); } }
}

/* ----- Batalkan Verifikasi: ID_EWO "Sesuai" kembali ke Approve Write-Off yang belum diverifikasi ----- */

function woOpenReopen(id) {
  const e = woAwoOf(id);
  if (!e || e.ver !== WO_VER_OK) return;
  wo.reopenId = id;
  put($('wr-form'),
    h('h2', { id: 'wr-title', text: 'Batalkan Verifikasi' }),
    h('p', { class: 'dlg-sub', text: id + ' · ' + e.b }),
    h('p', { class: 'wm-hint', text: 'Hasil verifikasi "' + WO_VER_OK + '"' + (e.verAt ? ' (' + fmtStamp(e.verAt) + (e.verBy ? ' · ' + e.verBy : '') + ')' : '')
      + ' akan dikosongkan. Status tetap "' + WO_ST_APPROVE + '", sehingga No JV dapat diubah dan Verifikasi Nilai WO dapat diisi ulang.' }),
    h('div', { class: 'wj-field' }, h('label', { for: 'wr-alasan', text: 'Alasan (wajib)' }),
      h('input', { id: 'wr-alasan', type: 'text', maxlength: '300', autocomplete: 'off', placeholder: 'Contoh: salah input No JV' })),
    h('p', { class: 'wm-err', id: 'wr-err', role: 'alert' }),
    h('div', { class: 'dlg-act' },
      h('button', { class: 'btn ghost', type: 'button', onclick: () => $('wo-reopen').close(), text: 'Batal' }),
      h('button', { class: 'btn fit', type: 'submit', id: 'wr-save', text: 'Batalkan Verifikasi' })));
  $('wo-reopen').showModal();
  $('wr-alasan').focus();
}

async function woReopenSubmit(ev) {
  ev.preventDefault();
  const id = wo.reopenId, input = $('wr-alasan'), err = $('wr-err'), save = $('wr-save');
  if (!id || !input) return;
  const alasan = input.value.trim();
  const msg = !alasan ? 'Alasan wajib diisi.' : alasan.length > 300 ? 'Alasan maksimal 300 karakter.' : '';
  input.classList.toggle('bad', !!msg); err.textContent = msg;
  if (msg) return input.focus();
  const ok = await askConfirm('Batalkan verifikasi ' + id + '?', 'Hasil "' + WO_VER_OK + '" dihapus dan ID_EWO kembali ke "' + WO_ST_APPROVE + '" yang belum diverifikasi. Pembatalan ini dicatat beserta alasannya.', 'Ya, batalkan');
  if (!ok) return;
  save.disabled = true;
  const r = await woCall('webWoReopen', { ewo: id, alasan }, { timeoutMs: 120000 });
  if (!r) return;
  save.disabled = false;
  if (!r.ok) { err.textContent = r.error || 'Gagal membatalkan verifikasi.'; return; }
  woApplyHead(r.head); wo.reopenId = ''; wo.rsb.at = 0;
  $('wo-reopen').close(); paintApprove();
  toast('Verifikasi ' + id + ' dibatalkan. Status tetap "' + WO_ST_APPROVE + '".');
}

function woOpenJv(id) {
  const e = woAwoOf(id);
  if (!e || e.ver === WO_VER_OK) return;
  wo.jvId = id;
  put($('wj-form'),
    h('h2', { id: 'wj-title', text: 'Input JV' }),
    h('p', { class: 'dlg-sub', text: id + ' · ' + e.b }),
    h('div', { class: 'wj-field' }, h('label', { for: 'wj-no', text: 'No JV' }),
      h('input', { id: 'wj-no', type: 'text', maxlength: String(WO_REF_MAX), autocomplete: 'off', spellcheck: 'false', value: e.jv && e.jv !== 'Beragam' ? e.jv : '' })),
    h('p', { class: 'wm-hint', text: 'No JV disimpan pada ' + fmt(e.n) + ' data aktif di sheet WO ' + e.jenis + ' (kolom No JV).' }),
    h('p', { class: 'wm-err', id: 'wj-err', role: 'alert' }),
    h('div', { class: 'dlg-act' },
      h('button', { class: 'btn ghost', type: 'button', onclick: () => $('wo-jv').close(), text: 'Batal' }),
      h('button', { class: 'btn fit', type: 'submit', id: 'wj-save', text: 'Simpan No JV' })));
  $('wo-jv').showModal();
  $('wj-no').focus();
}

async function woJvSubmit(ev) {
  ev.preventDefault();
  const id = wo.jvId, input = $('wj-no'), err = $('wj-err'), save = $('wj-save');
  if (!id || !input) return;
  const jv = input.value.trim();
  const msg = !jv ? 'No JV wajib diisi.' : woRefError(jv, 'No JV');
  input.classList.toggle('bad', !!msg); err.textContent = msg;
  if (msg) return input.focus();
  save.disabled = true;
  const r = await woCall('webWoSetJV', { ewo: id, jv }, { timeoutMs: 90000 });
  if (!r) return;
  save.disabled = false;
  if (!r.ok) { err.textContent = r.error || 'Gagal menyimpan No JV.'; return; }
  woApplyHead(r.head); wo.jvId = '';
  $('wo-jv').close(); paintApprove();
  toast('No JV ' + jv + ' tersimpan untuk ' + id + '.');
}

function woOpenVerif(id) {
  const e = woAwoOf(id);
  if (!e || e.ver === WO_VER_OK) return;
  wo.verifId = id;
  const kv = (k, v) => [h('dt', { text: k }), h('dd', { text: v })];
  put($('wv-body'),
    h('h2', { id: 'wv-title', text: 'Verifikasi Nilai WO' }),
    h('p', { class: 'dlg-sub', text: 'Apakah nilai WO pada ID_EWO ini sudah sesuai?' }),
    h('dl', { class: 'wo-kv' }, kv('ID_EWO', id), kv('Cabang', e.b), kv('Jenis / Tahun', 'WO ' + e.jenis + ' · ' + e.y), kv('Data aktif', fmt(e.n)),
      kv('PIUTANG YANG DI WO', rp(e.total)), kv('No JV', e.jv || 'Belum diinput')),
    h('p', { class: 'wm-hint', text: 'Sesuai: ID_EWO dinyatakan terverifikasi dan No JV terkunci (bisa dibuka lagi lewat Batalkan Verifikasi). Tidak: status kembali ke "' + WO_ST_CONFIRM + '" agar nilai bisa direvisi di menu Eksekusi Write-Off, lalu dieksekusi ulang.' }),
    h('p', { class: 'wm-err', id: 'wv-err', role: 'alert' }),
    h('div', { class: 'wo-choice' },
      h('button', { class: 'btn fit', type: 'button', id: 'wv-ok', onclick: () => woVerify(WO_VER_OK), text: 'Sesuai' }),
      h('button', { class: 'btn danger', type: 'button', id: 'wv-no', onclick: () => woVerify(WO_VER_NO), text: 'Tidak' })),
    h('div', { class: 'dlg-act' }, h('button', { class: 'btn ghost', type: 'button', onclick: () => $('wo-verif').close(), text: 'Tutup' })));
  $('wo-verif').showModal();
}

async function woVerify(hasil) {
  const id = wo.verifId, e = woAwoOf(id);
  if (!e) return;
  const ok = await askConfirm(hasil === WO_VER_OK ? 'Nyatakan nilai WO sesuai?' : 'Nyatakan nilai WO tidak sesuai?',
    hasil === WO_VER_OK ? id + ' akan terverifikasi "' + WO_VER_OK + '" dan No JV terkunci sampai verifikasi dibatalkan.'
      : id + ' akan dikembalikan ke status "' + WO_ST_CONFIRM + '" supaya nilainya direvisi di menu Eksekusi Write-Off, lalu hilang dari daftar ini sampai dieksekusi ulang.',
    hasil === WO_VER_OK ? 'Ya, Sesuai' : 'Ya, Tidak Sesuai');
  if (!ok) return;
  const a = $('wv-ok'), b = $('wv-no'), err = $('wv-err');
  a.disabled = b.disabled = true; err.textContent = '';
  const r = await woCall('webWoVerify', { ewo: id, hasil }, { timeoutMs: 120000 });
  if (!r) return;
  a.disabled = b.disabled = false;
  if (!r.ok) { err.textContent = r.error || 'Verifikasi gagal.'; return; }
  wo.verifId = '';
  if (r.head.st !== WO_ST_APPROVE) wo.awo = wo.awo.filter(x => x.id !== id);   // "Tidak": kembali ke Confirm, keluar dari daftar ini
  else woApplyHead(r.head);
  wo.optAt = 0; wo.curAt = 0; wo.rsb.at = 0;
  $('wo-verif').close(); paintApprove();
  toast(hasil === WO_VER_OK ? id + ' terverifikasi "' + WO_VER_OK + '".' : id + ' dikembalikan ke status "' + WO_ST_CONFIRM + '".');
}

/* ----- Setting > Write-Off: Backup / Clear / Sinkron Backup (ADMINISTRATOR; server memeriksa ulang) ----- */

const WOS_FORMAT = 'dashboard-write-off-backup', WOS_VERSION = 1, WOS_CLEAR_WORD = 'HAPUS';
const WOS_SEND_ROWS = 2000, WOS_SEND_BYTES = 1500000;   // batas isi satu permintaan Sinkron Backup (baris / karakter)

function woSettingPaint() {
  const b = wo.lastBackup;
  $('wos-last').textContent = b ? 'Backup terakhir pada sesi ini: ' + fmtStamp(b.at) + ' · ' + b.name + ' (' + fmt(b.rows) + ' baris di ' + b.tabs + ' sheet).' : 'Belum ada backup pada sesi ini.';
  ['wos-backup', 'wos-clear', 'wos-restore'].forEach(id => { $(id).disabled = wo.setBusy; });
}
function woSettingBusy(on, text, tone) {
  wo.setBusy = on; woSettingPaint();
  if (text != null) setNote('wos-msg', text, tone);
}
const woStamp = () => { const d = new Date(); return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '-' + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds()); };

// Baca seluruh sheet Write-Off dari server (per sheet, per potongan). Mengembalikan { tabs: [{ name, head, rows }], total } atau null bila gagal / dibatalkan.
async function woBackupRead(onStep) {
  const tabs = [];
  let count = 6, total = 0;
  for (let t = 0; t < count; t++) {
    let off = 0, grand = Infinity, tab = null;
    while (off < grand) {
      onStep('Membaca sheet ' + (tab ? tab.name : (t + 1) + '/' + count) + '… ' + fmt(off) + (isFinite(grand) ? ' dari ' + fmt(grand) : '') + ' baris');
      const r = await woCall('webWoBackup', { t, off }, { retry: true, timeoutMs: 180000 });
      if (!r) return null;
      if (!r.ok) throw new Error(r.error || 'Gagal membaca data Write-Off.');
      if (!tab) { tab = { name: r.name, head: r.head, exists: r.exists, rows: [] }; if (r.tabs) count = r.tabs.length; }
      grand = r.total;
      if (!r.rows.length) break;
      for (const row of r.rows) tab.rows.push(row);
      off += r.rows.length;
    }
    if (isFinite(grand) && tab.rows.length !== grand) throw new Error('Jumlah baris sheet ' + tab.name + ' tidak sesuai (terbaca ' + tab.rows.length + ' dari ' + grand + '). Coba lagi.');
    tabs.push(tab);
    total += tab.rows.length;
  }
  return { tabs, total };
}

async function woSettingBackup() {
  if (wo.setBusy) return;
  woSettingBusy(true, 'Menyiapkan backup…');
  try {
    const data = await woBackupRead(t => setNote('wos-msg', t));
    if (!data) return woSettingBusy(false);
    const u = state.user || {};
    const file = { format: WOS_FORMAT, version: WOS_VERSION, createdAt: new Date().toISOString(), createdBy: { nik: u.nik || '', nama: u.nama || '' },
      counts: Object.fromEntries(data.tabs.map(t => [t.name, t.rows.length])), tabs: data.tabs.map(t => ({ name: t.name, head: t.head, rows: t.rows })) };
    const name = 'Backup Write-Off ' + woStamp() + '.json';
    const url = URL.createObjectURL(new Blob([JSON.stringify(file)], { type: 'application/json' })), a = h('a', { href: url, download: name });
    document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 4000);
    wo.lastBackup = { at: new Date().toISOString(), name, rows: data.total, tabs: data.tabs.length };
    woSettingBusy(false, 'Backup selesai: "' + name + '" berisi ' + fmt(data.total) + ' baris dari ' + data.tabs.length + ' sheet (' + data.tabs.map(t => t.name + ' ' + fmt(t.rows.length)).join(', ') + ').', 'good');
    toast('Backup Write-Off berhasil dibuat.');
  } catch (e) { woSettingBusy(false, e.message || 'Backup gagal.', 'bad'); }
}

function woOpenClear() {
  if (wo.setBusy) return;
  put($('wc-form'),
    h('h2', { id: 'wc-title', text: 'Clear data Write-Off?' }),
    h('p', { class: 'wm-hint', text: 'Seluruh isi data Write-Off (WO Reguler, WO Khusus, EWO, EWO_Log, EWO_TTD, EWO_Surat) di Google Sheet dan cache layar akan dihapus. Tab dan header sheet tetap ada. '
      + 'Penghapusan tidak dapat dibatalkan, kecuali Anda punya file backup untuk Sinkron Backup.' + (wo.lastBackup ? '' : ' Pada sesi ini Anda BELUM membuat backup.') }),
    h('div', { class: 'wj-field' }, h('label', { for: 'wc-word', text: 'Ketik ' + WOS_CLEAR_WORD + ' untuk melanjutkan' }),
      h('input', { id: 'wc-word', type: 'text', autocomplete: 'off', spellcheck: 'false' })),
    h('p', { class: 'wm-err', id: 'wc-err', role: 'alert' }),
    h('div', { class: 'dlg-act' },
      h('button', { class: 'btn ghost', type: 'button', onclick: () => $('wo-clear').close(), text: 'Kembali' }),
      h('button', { class: 'btn danger', type: 'submit', id: 'wc-go', text: 'Clear sekarang' })));
  $('wo-clear').showModal();
  $('wc-word').focus();
}

async function woClearSubmit(ev) {
  ev.preventDefault();
  const word = $('wc-word').value.trim();
  if (word !== WOS_CLEAR_WORD) { $('wc-err').textContent = 'Ketik ' + WOS_CLEAR_WORD + ' (huruf besar) untuk melanjutkan.'; return $('wc-word').focus(); }
  $('wo-clear').close();
  woSettingBusy(true, 'Membersihkan data Write-Off…');
  const r = await woCall('webWoClear', { confirm: word }, { timeoutMs: 180000 });
  if (!r) return woSettingBusy(false);
  if (!r.ok) return woSettingBusy(false, r.error || 'Clear gagal.', 'bad');
  const n = (r.tabs || []).reduce((a, t) => a + t.cleared, 0);
  woReset();                              // cache layar (pilihan filter, daftar ID_EWO, data yang sedang dibuka, Report, Surat) ikut bersih
  woSettingBusy(false, 'Clear selesai: ' + fmt(n) + ' baris dihapus dari ' + (r.tabs || []).filter(t => t.exists).length + ' sheet, termasuk log. Tab dan header sheet tetap ada.', 'good');
  toast('Data Write-Off dibersihkan.');
}

// Validasi isi file backup di sisi klien (server memeriksa lagi tiap baris). Mengembalikan teks galat atau ''.
function woRestoreCheck(f) {
  if (!f || typeof f !== 'object' || f.format !== WOS_FORMAT) return 'Ini bukan file backup Write-Off dari aplikasi ini.';
  if (f.version !== WOS_VERSION) return 'Versi file backup (' + f.version + ') tidak dikenali.';
  if (!Array.isArray(f.tabs) || !f.tabs.length) return 'File backup tidak memuat sheet.';
  for (const t of f.tabs) {
    if (!t || typeof t.name !== 'string' || !Array.isArray(t.head) || !Array.isArray(t.rows)) return 'Struktur sheet pada file backup tidak lengkap.';
    if (f.counts && Number.isInteger(f.counts[t.name]) && f.counts[t.name] !== t.rows.length) return 'Jumlah baris sheet "' + t.name + '" tidak cocok dengan catatan di file (file rusak atau terpotong).';
  }
  return '';
}

async function woSettingRestoreFile(file) {
  if (!file || wo.setBusy) return;
  let data;
  try { data = JSON.parse(await file.text()); } catch (e) { return setNote('wos-msg', 'File tidak dapat dibaca sebagai JSON.', 'bad'); }
  const bad = woRestoreCheck(data);
  if (bad) return setNote('wos-msg', bad, 'bad');
  const total = data.tabs.reduce((a, t) => a + t.rows.length, 0);
  const ok = await askConfirm('Sinkron Backup?', 'File "' + file.name + '" (dibuat ' + (data.createdAt ? fmtStamp(data.createdAt) : '—') + ') berisi ' + fmt(total) + ' baris dari ' + data.tabs.length
    + ' sheet. Isi sheet Write-Off yang ada sekarang akan DIGANTI dengan isi file ini.', 'Ya, sinkronkan');
  if (!ok) return;
  woSettingBusy(true, 'Memulihkan data Write-Off…');
  try {
    const known = ['WO Reguler', 'WO Khusus', 'EWO', 'EWO_Log', 'EWO_TTD', 'EWO_Surat'];
    let written = 0;
    for (const tab of data.tabs) {
      const t = known.indexOf(tab.name);
      if (t < 0) throw new Error('Sheet "' + tab.name + '" pada file backup tidak dikenali.');
      let i = 0, first = true;
      while (first || i < tab.rows.length) {
        let n = 0, size = 0;
        while (i + n < tab.rows.length && n < WOS_SEND_ROWS && (n === 0 || size + JSON.stringify(tab.rows[i + n]).length <= WOS_SEND_BYTES)) { size += JSON.stringify(tab.rows[i + n]).length; n++; }
        setNote('wos-msg', 'Memulihkan sheet ' + tab.name + '… ' + fmt(i) + ' dari ' + fmt(tab.rows.length) + ' baris');
        const r = await woCall('webWoRestore', { t, start: first ? '1' : '0', offset: i, rows: JSON.stringify(tab.rows.slice(i, i + n)) }, { timeoutMs: 180000 });
        if (!r) return woSettingBusy(false);
        if (!r.ok) throw new Error((r.error || 'Sinkron Backup gagal.') + ' Sheet "' + tab.name + '" mungkin baru terisi sebagian; ulangi Sinkron Backup dengan file yang sama.');
        written += r.written; i += n; first = false;
      }
    }
    woReset();
    woSettingBusy(false, 'Sinkron Backup selesai: ' + fmt(written) + ' baris dipulihkan ke ' + data.tabs.length + ' sheet.', 'good');
    toast('Data Write-Off dipulihkan dari backup.');
  } catch (e) { woSettingBusy(false, e.message || 'Sinkron Backup gagal.', 'bad'); }
}

function woBindSetting() {
  $('wos-backup').onclick = woSettingBackup;
  $('wos-clear').onclick = woOpenClear;
  $('wos-restore').onclick = () => { if (!wo.setBusy) { $('wos-file').value = ''; $('wos-file').click(); } };
  $('wos-file').addEventListener('change', () => woSettingRestoreFile($('wos-file').files[0]));
  $('wc-form').addEventListener('submit', woClearSubmit);
}

function bindWriteOff() {
  $('ewo-refresh').onclick = woRefreshAll;
  ['ewo-branch', 'ewo-jenis', 'ewo-tahun'].forEach(id => $(id).addEventListener('change', () => woFilterChanged(id)));
  $('ewo-proses').onclick = woProses;
  $('ewo-download').onclick = woDownload;
  $('ewo-close').onclick = woCloseWork;
  $('ewo-exec').onclick = woExecute;
  $('ewo-q').addEventListener('input', debounce(() => { wo.q = $('ewo-q').value; wo.page = 0; paintEwoTable(); }, 150));
  $('ewo-tbl').addEventListener('click', woRowClick);
  $('ewo-hist').addEventListener('click', woHistClick);
  $('wm-form').addEventListener('input', woManualCalc);
  $('wm-form').addEventListener('submit', woManualSubmit);
  $('wj-form').addEventListener('submit', woJvSubmit);
  $('awo-refresh').onclick = () => { wo.awoAt = 0; woLoadApprove(); };
  woBindReport();
  $('wr-form').addEventListener('submit', woReopenSubmit);
  $('wb-form').addEventListener('submit', woBatalSubmit);
  woBindSetting();
  $('awo-jenis').addEventListener('change', () => { wo.awoJ = $('awo-jenis').value; paintApprove(); });
  $('awo-tbl').addEventListener('click', woApproveClick);
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
  TABS.forEach(k => { $('t-' + k).onclick = NAV_GROUPS.indexOf(k) > -1 ? () => navParentClick(k) : () => selectTab(k); });
  bindNavGroups();
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
    if (!e.target.classList || !e.target.classList.contains('tab')) return;
    const vis = TABS.filter(k => !$('t-' + k).hidden), i = vis.findIndex(k => $('t-' + k).getAttribute('aria-selected') === 'true');
    const to = e.key === 'ArrowRight' ? vis[(i + 1) % vis.length] : e.key === 'ArrowLeft' ? vis[(i - 1 + vis.length) % vis.length]
      : e.key === 'Home' ? vis[0] : e.key === 'End' ? vis[vis.length - 1] : null;
    if (to) { e.preventDefault(); selectTab(to, true); }
  });

  // Write-Off: Eksekusi & Approve
  bindWriteOff();

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
  $('pola-form').addEventListener('submit', e => { e.preventDefault(); savePola(); });
  $('pola-list').addEventListener('change', () => { state.polaDirty = !!state.pola && polaSelected() !== state.pola; setNote('pola-msg', ''); syncPolaForm(); });
  $('image-download-form').addEventListener('submit', e => { e.preventDefault(); saveImageDownloadAccess(); });
  $('image-download-level-list').addEventListener('change', () => {
    state.imageDownloadDirty = imageDownloadChecked().join('|') !== (state.imageDownloadLevels || []).join('|');
    paintImageDownloadSettings();
    setNote('image-download-msg', '');
  });
  $('wjlist').addEventListener('change', () => { state.wjDirty = !!state.data && wajibChecked() !== state.data.wajib; setNote('wjmsg', ''); syncWajibForm(); });
  $('wilayah-user-q').addEventListener('input', debounce(renderWilayahUsers, 120));
  $('wilayah-branch-q').addEventListener('input', debounce(renderWilayahEditor, 120));
  $('wilayah-all').onclick = () => setWilayahAll(true);
  $('wilayah-none').onclick = () => setWilayahAll(false);
  $('wilayah-save').onclick = saveWilayah;

  // Manajemen (show/hide menu per tingkat)
  $('mg-wrap').addEventListener('change', mgChange);
  $('mg-wrap').addEventListener('click', mgClick);
  $('mg-save').onclick = mgSave;
  $('mg-undo').onclick = mgUndo;
  $('mg-default').onclick = mgDefault;
  $('mg-refresh').onclick = mgRefresh;

  // Unggah
  $('file').onchange = e => pickFile(e.target.files[0]);
  const drop = $('drop');
  ['dragover', 'dragenter'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => pickFile(e.dataTransfer.files[0]));
  $('upgo').onclick = sendUpload;
  $('refresh-all').onclick = refreshAllData;
  // Filter Kabar Berita: layar dan gambar unduhan memakai pilihan yang sama.
  $('kb-branch').onchange = e => {
    state.kb.branch = e.target.value || '';
    state.newsLoadSeq++; state.newsPromise = null; state.newsLoading = false; state.newsReady = false;   // batalkan alur lama; hasil yang sudah tiba tetap disimpan
    kbFilterPaint(); paintKabarBerita(); prepareRegionalNewsData();
  };
  $('kb-refresh').onclick = () => prepareRegionalNewsData({ force: true });

  // Pintasan: "/" langsung ke kolom cari; tab dibuka kembali setelah lama = cek pembaruan data
  document.addEventListener('keydown', e => {
    if (e.key === '/' && state.user && !$('p-dashboard').hidden && state.dashboardPage === 'dash' && !$('dbody').hidden && !e.target.closest('input,select,textarea,dialog')) { e.preventDefault(); $('dq').focus(); }
  });
  document.addEventListener('visibilitychange', () => {
    // Begitu aplikasi kembali aktif, periksa versi/ruang lingkup di server tanpa menunggu interval.
    // Pemicu ini menggunakan penghitung yang sama dengan interval berkala.
    if (!document.hidden && state.user) {
      menuRefresh();
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
