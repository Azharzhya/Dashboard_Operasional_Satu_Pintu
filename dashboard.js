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
  angsuran: { label: 'Jumlah Angsuran', get: (r, k) => r[F.CICILAN] > 0 ? Math.floor((r[F.SUKARELA] + wajibOf(r, k)) / r[F.CICILAN]) : 0, num: true },
  sisa: { label: 'Sisa Pelunasan', get: (r, k) => r[F.TBADJ] - r[F.KUM] - wajibOf(r, k), num: true },
  sisaTm: { label: 'Total Balance - Simpanan', get: (r, k) => r[F.BALANCE] - r[F.KUM] - wajibOf(r, k), num: true },
  note: { label: 'Catatan', get: (r, k) => noteText(r[F.NOTE], k), num: false },
  ya: { label: 'Keterangan', get: () => 'YA' },    // hanya di file Excel (tetap seperti versi lama)
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
    const ctl = new AbortController(), timer = setTimeout(() => ctl.abort(), API_TIMEOUT_MS);
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
  user: null, loadSeq: 0, loading: false, lastCheck: 0,
  data: null, byCat: {}, branches: new Map(), groups: {}, clientIdCounts: new Map(),
  cat: 'a', q: '', branch: '', staf: '', sort: { label: '', dir: 1 }, page: 0, size: 50, view: [],
  sumBy: 'branch', sumSort: { key: 'n', dir: -1 }, sumQ: '', wjDirty: false,
  wilayah: null, wilayahSelected: '', wilayahDraft: [], wilayahDrafts: Object.create(null), wilayahLoading: false, wilayahLoadedAt: 0,
};

function showView(name, msg) {
  $('splash').hidden = name !== 'splash';
  $('login').hidden = name !== 'login';
  $('app').hidden = name !== 'app';
  if (name === 'login') { showErr(msg); $(msg ? 'pw' : 'nik').focus(); }
}
function showErr(msg) { $('err').hidden = !msg; $('err').textContent = msg || ''; }

const canSetting = u => !!u && (u.tingkat === 'ADMINISTRATOR' || u.tingkat === 'USER 0');   // menu Setting
const TABS = ['dash', 'sum', 'set', 'up'];
function selectTab(which, focus) {
  if (which === 'up' && (!state.user || state.user.tingkat !== 'ADMINISTRATOR')) which = 'dash';
  if (which === 'set' && !canSetting(state.user)) which = 'dash';
  TABS.forEach(k => {
    const on = k === which;
    $('t-' + k).setAttribute('aria-selected', String(on));
    $('t-' + k).tabIndex = on ? 0 : -1;
    $('p-' + k).hidden = !on;
  });
  if (which === 'sum') drawSummary();
  if (focus) $('t-' + which).focus();
}

function paintUser(user) {
  const name = user.nama || user.nik || '';
  $('uname').textContent = name;
  $('ulevel').textContent = user.tingkat || '';
  $('uavatar').textContent = name.split(/\s+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('').toUpperCase();
  $('t-set').hidden = !canSetting(user);
  if (!canSetting(user) && !$('p-set').hidden) selectTab('dash');
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

function startAutoSync() {
  stopAutoSync();
  const poll = () => {
    // Polling hanya saat aplikasi aktif dan data awal sudah tersedia.
    // Jika tidak ada perubahan, server menjawab same:true dengan respons ringan.
    if (!state.user || document.hidden || state.loading || !state.data) return;
    loadDash({ automatic: true });
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

function logout(msg) {
  stopAutoSync();
  state.loadSeq++; state.loading = false; state.user = null; state.data = null; state.byCat = {}; state.groups = {}; state.clientIdCounts = new Map(); state.wjDirty = false;
  session.clear(); cache.clear(); state.wilayah = null; state.wilayahDraft = []; state.wilayahDrafts = Object.create(null); state.wilayahSelected = ''; setNote('wjmsg', '');
  clearDash(); resetUpload();
  showView('login', msg || '');
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
  session.set({ token: r.token, expires: r.expires, user: r.user });
  $('pw').value = '';
  enterApp(r.user);
}

/* ---------------------------------------------------------------------
   6. Dashboard
   --------------------------------------------------------------------- */
function enterApp(user) {
  state.user = user;
  paintUser(user);
  state.cat = 'a'; state.q = ''; state.branch = ''; state.staf = ''; state.page = 0; state.sort = { label: '', dir: 1 };
  resetUpload(); setNote('setmsg', ''); setNote('wjmsg', ''); state.wjDirty = false; selectTab('dash'); showView('app');
  startAutoSync();
  loadDash();
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
  if (r.same) return loadDashFresh();
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

async function loadDashFresh() { state.data = null; state.loading = false; await cache.clear(); return loadDash(); }

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
  selectTab('dash');
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

async function readExcel(file) {
  if (file.size > UP_MAX_MB * 1048576) throw new Error('Ukuran file melebihi ' + UP_MAX_MB + ' MB.');
  const buf = await file.arrayBuffer(), b = new Uint8Array(buf.slice(0, 4));
  const zip = b[0] === 0x50 && b[1] === 0x4B, ole = b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0;
  if (!/\.(xlsx|xlsm|xls)$/i.test(file.name) || !(zip || ole)) throw new Error('File harus berformat Excel (.xlsx atau .xls).');
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
   10. Unggah
   --------------------------------------------------------------------- */
let pending = null;

function upBar(f) { $('upbar').hidden = f == null; if (f != null) $('upfill').style.width = Math.round(f * 100) + '%'; }
function resetUpload() {
  pending = null; $('upinfo').hidden = true; $('file').value = ''; $('upstats').replaceChildren(); $('uptbl').replaceChildren();
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

async function pickFile(file) {
  if (!file) return;
  resetUpload(); $('upinfo').hidden = false; $('upgo').disabled = true; setNote('upmsg', 'Membaca file…');
  await nextFrame();                                       // beri kesempatan layar menampilkan status
  try {
    pending = await readExcel(file);
    renderPreview(pending);
    setNote('upmsg', 'Periksa pratinjau, lalu kirim. Data lama di tab Data_Upload akan diganti.');
    $('upgo').disabled = false;
  } catch (e) {
    pending = null; $('upstats').replaceChildren(); $('uptbl').replaceChildren();
    setNote('upmsg', e.message || 'File tidak bisa dibaca.', 'bad');
  }
}

function upFail(r, sent) {
  if (r.code === 'SESSION_REQUIRED') return logout(r.error);
  upBar(null); $('upgo').disabled = false;
  setNote('upmsg', (r.error || 'Gagal mengirim data.') + (sent ? ' Data di Sheet baru terisi ' + fmt(sent) + ' baris, ulangi unggah.' : ''), 'bad');
}

async function sendUpload() {
  if (!pending) return;
  const total = pending.rows.length, tok = (session.get() || {}).token;
  const ok = await askConfirm('Ganti seluruh data?', 'Seluruh isi tab Data_Upload akan diganti dengan ' + fmt(total) + ' baris dari file ini. Data lama tidak bisa dikembalikan.', 'Ganti data');
  if (!ok) return;
  $('upgo').disabled = true; upBar(0); setNote('upmsg', 'Menyiapkan tab Data_Upload…');
  let r = await api('webUploadStart', { token: tok, total, file: pending.name });
  if (!r.ok) return upFail(r, 0);
  const id = r.uploadId; let sent = 0;
  for (let seq = 0; sent < total; seq++) {
    const part = pending.rows.slice(sent, sent + UP_CHUNK); let good = false, res;
    for (let t = 0; t < 3 && !good; t++) {                 // kiriman ulang aman: server mengabaikan potongan yang sudah masuk
      res = await api('webUploadChunk', { token: tok, uploadId: id, seq, rows: JSON.stringify(part) });
      good = res.ok;
      if (!good) { if (res.code === 'SESSION_REQUIRED') break; await sleep(1500 * (t + 1)); }
    }
    if (!good) return upFail(res, sent);
    sent += part.length; upBar(sent / total);
    setNote('upmsg', 'Mengirim ' + fmt(sent) + ' dari ' + fmt(total) + ' baris…');
  }
  setNote('upmsg', 'Menyimpan dan menghitung ulang dashboard…');
  r = await api('webUploadFinish', { token: tok, uploadId: id, total });
  if (!r.ok) return upFail(r, total);
  upBar(1); setNote('upmsg', 'Selesai: ' + fmt(r.rows) + ' baris tersimpan di tab Data_Upload.', 'good');
  toast('Data diperbarui. Dashboard dimuat ulang.');
  pending = null; loadDash();                              // segarkan dashboard dengan data baru
}

/* ---------------------------------------------------------------------
   11. Unduh Excel
   --------------------------------------------------------------------- */
// Satu lembar dari daftar kolom + baris. Angka diberi format ribuan & ada filter di judul.
function sheetOf(cols, rows, k) {
  const ws = XLSX.utils.aoa_to_sheet([cols.map(c => c.label)].concat(rows.map(r => cols.map(c => c.get(r, k)))));
  ws['!cols'] = cols.map(c => ({ wch: Math.max(12, c.label.length + 2) }));
  // Warnai sel Client ID yang berulang dengan merah pada file Excel hasil ekspor.
  const clientCol = cols.findIndex(c => c.label === 'Client ID');
  if (clientCol >= 0) {
    for (let i = 0; i < rows.length; i++) {
      const id = clientIdKey(rows[i][F.CLIENT]);
      const count = id ? (state.clientIdCounts.get(id) || 0) : 0;
      if (count > 1) {
        const addr = XLSX.utils.encode_cell({ r: i + 1, c: clientCol });
        const cell = ws[addr];
        if (cell) cell.s = {
          fill: { patternType: 'solid', fgColor: { rgb: 'FFFFD6D6' } },
          font: { bold: true, color: { rgb: 'FFB42318' } },
        };
      }
    }
  }
  cols.forEach((c, j) => {
    if (!c.num) return;
    for (let i = 1; i <= rows.length; i++) { const cell = ws[XLSX.utils.encode_cell({ r: i, c: j })]; if (cell) cell.z = '#,##0'; }
  });
  if (rows.length) ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: cols.length - 1 } }) };
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

function exportAll() {
  if (!state.data) return;
  download($('ddl'), wb => {
    const sumCols = [{ label: 'Kategori', get: c => c.title }, { label: 'Jumlah Baris', get: c => state.byCat[c.k].length, num: true },
      { label: 'Total Balance', get: c => state.byCat[c.k].reduce((a, r) => a + r[F.BALANCE], 0), num: true },
      { label: 'Simp. Wajib Dihitung', get: c => inc(c.k) ? 'YA' : 'TIDAK' }];
    XLSX.utils.book_append_sheet(wb, sheetOf(sumCols, CATS), 'Ringkasan');
    const bg = groupsOf('branch').slice().sort((a, b) => collator.compare(a.name, b.name));
    XLSX.utils.book_append_sheet(wb, sheetOf([{ label: 'Cabang', get: g => g.name }, { label: 'Pinjaman', get: g => g.n, num: true },
      ...CATS.map(c => ({ label: c.title, get: g => g[c.k], num: true })), { label: 'Total Balance', get: g => g.balance, num: true }], bg), 'Per Cabang');
    CATS.forEach(c => XLSX.utils.book_append_sheet(wb, sheetOf(excelCols(c), state.byCat[c.k], c.k), c.sheet));
  }, 'Dashboard_' + today() + '.xlsx');
}

function exportView() {   // kategori yang sedang dibuka, lengkap dengan filter & urutan saat ini
  if (!state.data || !state.view.length) return;
  const c = catOf(state.cat);
  download($('dxl'), wb => XLSX.utils.book_append_sheet(wb, sheetOf(excelCols(c), state.view, c.k), c.sheet), 'Dashboard_' + c.sheet.replace(/\s+/g, '_') + '_' + today() + '.xlsx');
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
  $('app').querySelector('.tabs').addEventListener('keydown', e => {
    const vis = TABS.filter(k => !$('t-' + k).hidden), i = vis.findIndex(k => $('t-' + k).getAttribute('aria-selected') === 'true');
    const to = e.key === 'ArrowRight' ? vis[(i + 1) % vis.length] : e.key === 'ArrowLeft' ? vis[(i - 1 + vis.length) % vis.length]
      : e.key === 'Home' ? vis[0] : e.key === 'End' ? vis[vis.length - 1] : null;
    if (to) { e.preventDefault(); selectTab(to, true); }
  });

  // Dashboard
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
    if (e.key === '/' && state.user && !$('p-dash').hidden && !$('dbody').hidden && !e.target.closest('input,select,textarea,dialog')) { e.preventDefault(); $('dq').focus(); }
  });
  document.addEventListener('visibilitychange', () => {
    // Begitu aplikasi kembali aktif, periksa versi/ruang lingkup di server tanpa menunggu interval.
    // loadDash memakai ver yang sedang dimiliki: jika sama, server hanya mengirim respons kecil (same:true).
    if (!document.hidden && state.user && !state.loading) {
      loadDash({ automatic: true });
    }
  });
}

async function init() {
  bind();
  try { indexedDB.deleteDatabase(CACHE_DB_OLD); } catch (e) { /* abaikan */ }   // salinan data versi lama tidak dipakai lagi
  const s = session.get();
  if (!s || !s.token || !(s.expires > Date.now())) { session.clear(); return showView('login'); }
  if (s.user) { enterApp(s.user); return; }                // pengguna tersimpan: langsung tampil, server memverifikasi sambil memuat data
  const r = await api('webMe', { token: s.token }, { retry: true });   // sesi dari versi lama (belum menyimpan pengguna)
  if (r.ok) { session.set(Object.assign(s, { user: r.user })); return enterApp(r.user); }
  if (r.code === 'SESSION_REQUIRED') session.clear();
  showView('login', r.error);
}
init();
})();
