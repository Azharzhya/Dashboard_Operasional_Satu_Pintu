/* surat-wo.js -- penyusun PDF "Surat Berita Acara" Write Off (menu Report Writte Off).
 *
 * Mandiri: tanpa pustaka PDF dari luar. Memakai dua berkas yang sudah ada di folder yang sama:
 *   surat-kop.js   -> window.SURAT_KOP    { x, y, w, h, data }  gambar kop (PNG base64) dan posisinya pada halaman A4 (poin)
 *   surat-font.js  -> window.SURAT_FONTS  { family, files: { normal, bold, italic } }  Liberation Serif (TrueType, base64)
 * Kedua berkas itu dimuat sesuai kebutuhan oleh dashboard.js (loadSuratWO) bersama berkas ini.
 *
 * Pemakaian:  const pdf = await SuratWO.build(letter);   // Blob application/pdf
 * letter  = isi surat dari server (webWoSuratData):
 *   { id, jenis, y, b, hari, tanggal, active: [[nama, client6, ctr, produk, total, voucher]...], batal: [[nama, client6, ctr, produk, total]...],
 *     totalActive, totalBatal, msa: { nama, img }, bm: { nama, img } }       img = data URL JPEG TTD ('' = ruang tanda tangan dikosongkan)
 *
 * Tata letak mengikuti Berita_Acara_WO_Reguler_or_Khusus_*.docx: A4, Times (Liberation Serif), kop di setiap halaman, judul, paragraf
 * pembuka rata kiri-kanan, dua tabel (baris judul diulang di halaman berikutnya), penutup, dua kolom tanda tangan, dan Tembusan.
 * Tanda tangan tidak pernah terpotong antar halaman: bila ruangnya kurang, seluruh blok penutup pindah ke halaman baru.
 */
(function (root) {
  'use strict';

  const enc = new TextEncoder();
  const PAGE_W = 595.28, PAGE_H = 841.89;
  const ML = 65, MR = 66, TW = PAGE_W - ML - MR;       // margin kiri / kanan seperti templat Word (1300 / 1320 twip)
  const TOP = 126, BOTTOM = 796;                       // area isi di bawah kop; batas bawah isi
  const BODY = 10, BODY_LEAD = 15;                     // paragraf
  const CELL = 9, CELL_LEAD = 11, PAD_X = 4, PAD_Y = 3.5;
  const SIG_BOX_W = 150, SIG_BOX_H = 62;

  /* ---------------------------------------------------------------------
     Berkas biner
     --------------------------------------------------------------------- */
  function b64ToBytes(b64) {
    const bin = atob(b64), out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  const dataUrlBytes = url => b64ToBytes(String(url).slice(String(url).indexOf(',') + 1));

  async function deflate(u8) {
    if (typeof CompressionStream === 'undefined') return null;
    try {
      const cs = new CompressionStream('deflate'), w = cs.writable.getWriter();
      w.write(u8); w.close();
      return new Uint8Array(await new Response(cs.readable).arrayBuffer());
    } catch (e) { return null; }
  }

  /* ---------------------------------------------------------------------
     TrueType: hanya yang dibutuhkan (cmap, lebar huruf, ukuran dasar)
     --------------------------------------------------------------------- */
  function parseTTF(buf) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    const n = dv.getUint16(4), T = {};
    for (let i = 0; i < n; i++) {
      const o = 12 + i * 16;
      T[String.fromCharCode(buf[o], buf[o + 1], buf[o + 2], buf[o + 3])] = { off: dv.getUint32(o + 8), len: dv.getUint32(o + 12) };
    }
    if (!T.head || !T.hhea || !T.maxp || !T.hmtx || !T.cmap) throw new Error('Font surat tidak lengkap (tabel TrueType hilang).');
    const upm = dv.getUint16(T.head.off + 18);
    const f = {
      bytes: buf, upm: upm,
      bbox: [dv.getInt16(T.head.off + 36), dv.getInt16(T.head.off + 38), dv.getInt16(T.head.off + 40), dv.getInt16(T.head.off + 42)],
      ascent: dv.getInt16(T.hhea.off + 4), descent: dv.getInt16(T.hhea.off + 6),
      nHM: dv.getUint16(T.hhea.off + 34), numGlyphs: dv.getUint16(T.maxp.off + 4),
      italicAngle: T.post ? dv.getInt32(T.post.off + 4) / 65536 : 0, capHeight: Math.round(upm * 0.66)
    };
    if (T['OS/2'] && T['OS/2'].len >= 90 && dv.getUint16(T['OS/2'].off) >= 2) f.capHeight = dv.getInt16(T['OS/2'].off + 88) || f.capHeight;
    f.adv = new Array(f.numGlyphs);
    for (let g = 0; g < f.numGlyphs; g++) {
      const i = Math.min(g, f.nHM - 1);
      f.adv[g] = dv.getUint16(T.hmtx.off + i * 4);
    }
    // cmap Unicode: format 12 lebih dulu, lalu format 4
    const c0 = T.cmap.off, nt = dv.getUint16(c0 + 2);
    let best = null, bestScore = -1;
    for (let i = 0; i < nt; i++) {
      const pid = dv.getUint16(c0 + 4 + i * 8), eid = dv.getUint16(c0 + 6 + i * 8), off = c0 + dv.getUint32(c0 + 8 + i * 8), fmt = dv.getUint16(off);
      const uni = (pid === 3 && (eid === 1 || eid === 10)) || pid === 0;
      const score = !uni ? -1 : fmt === 12 ? 2 : fmt === 4 ? 1 : -1;
      if (score > bestScore) { bestScore = score; best = { off, fmt }; }
    }
    if (!best) throw new Error('Font surat tidak memiliki tabel cmap Unicode.');
    f.gid = cp => {
      const o = best.off;
      if (best.fmt === 12) {
        const ng = dv.getUint32(o + 12);
        for (let i = 0; i < ng; i++) {
          const g = o + 16 + i * 12, s = dv.getUint32(g), e = dv.getUint32(g + 4);
          if (cp >= s && cp <= e) return dv.getUint32(g + 8) + (cp - s);
        }
        return 0;
      }
      const segX2 = dv.getUint16(o + 6), endO = o + 14, startO = endO + segX2 + 2, deltaO = startO + segX2, rangeO = deltaO + segX2;
      for (let i = 0; i < segX2 / 2; i++) {
        if (cp > dv.getUint16(endO + i * 2)) continue;
        const s = dv.getUint16(startO + i * 2);
        if (cp < s) return 0;
        const delta = dv.getInt16(deltaO + i * 2), ro = dv.getUint16(rangeO + i * 2);
        if (!ro) return (cp + delta) & 0xFFFF;
        const g = dv.getUint16(rangeO + i * 2 + ro + (cp - s) * 2);
        return g ? (g + delta) & 0xFFFF : 0;
      }
      return 0;
    };
    return f;
  }

  /* Satu gaya font: lebar teks, pengodean ke ID glyph, catatan karakter terpakai (untuk ToUnicode) */
  class Face {
    constructor(name, bytes, flags, stemV) {
      this.name = name; this.f = parseTTF(bytes); this.flags = flags; this.stemV = stemV;
      this.used = new Map(); this.cache = new Map(); this.id = 0;
    }
    glyph(ch) {
      let g = this.cache.get(ch);
      if (g !== undefined) return g;
      let cp = ch.codePointAt(0);
      if (cp === 0xA0) cp = 0x20;
      g = this.f.gid(cp);
      if (!g) { g = this.f.gid(0x3F); cp = 0x3F; }          // karakter di luar font -> "?"
      this.cache.set(ch, g);
      if (g) this.used.set(g, cp);
      return g;
    }
    width(s, size) {
      let w = 0;
      for (const ch of s) w += this.f.adv[this.glyph(ch)] || 0;
      return w * size / this.f.upm;
    }
    hex(s) {
      let o = '';
      for (const ch of s) o += this.glyph(ch).toString(16).padStart(4, '0');
      return '<' + o + '>';
    }
  }

  /* ---------------------------------------------------------------------
     Gambar: JPEG (TTD) dan PNG (kop) dimasukkan apa adanya bila memungkinkan; selain itu dirasterkan lewat canvas
     --------------------------------------------------------------------- */
  function jpegInfo(b) {
    if (b[0] !== 0xFF || b[1] !== 0xD8) throw new Error('Gambar TTD bukan JPEG.');
    let i = 2;
    while (i + 9 < b.length) {
      if (b[i] !== 0xFF) { i++; continue; }
      const m = b[i + 1];
      if (m === 0xFF) { i++; continue; }
      if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) return { w: (b[i + 7] << 8) | b[i + 8], h: (b[i + 5] << 8) | b[i + 6], comps: b[i + 9] };
      i += 2 + ((b[i + 2] << 8) | b[i + 3]);
    }
    throw new Error('Ukuran gambar TTD tidak terbaca.');
  }

  function pngInfo(b) {
    const sig = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
    if (sig.some((v, i) => b[i] !== v)) return null;
    const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
    let i = 8, ihdr = null, plte = null, trns = false;
    const idat = [];
    while (i + 8 <= b.length) {
      const len = dv.getUint32(i), type = String.fromCharCode(b[i + 4], b[i + 5], b[i + 6], b[i + 7]), d = i + 8;
      if (type === 'IHDR') ihdr = { w: dv.getUint32(d), h: dv.getUint32(d + 4), depth: b[d + 8], ctype: b[d + 9], interlace: b[d + 12] };
      else if (type === 'PLTE') plte = b.subarray(d, d + len);
      else if (type === 'tRNS') trns = true;
      else if (type === 'IDAT') idat.push(b.subarray(d, d + len));
      else if (type === 'IEND') break;
      i += 12 + len;
    }
    if (!ihdr) return null;
    return Object.assign(ihdr, { plte, trns, idat });
  }

  async function rasterize(dataUrl) {
    if (typeof document === 'undefined') throw new Error('Format gambar kop tidak didukung di lingkungan ini.');
    const img = await new Promise((ok, no) => { const im = new Image(); im.onload = () => ok(im); im.onerror = () => no(new Error('Gambar tidak dapat dibaca.')); im.src = dataUrl; });
    const c = document.createElement('canvas');
    c.width = img.naturalWidth; c.height = img.naturalHeight;
    const g = c.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(img, 0, 0);
    const px = g.getImageData(0, 0, c.width, c.height).data, rgb = new Uint8Array(c.width * c.height * 3);
    for (let i = 0, j = 0; i < px.length; i += 4) { rgb[j++] = px[i]; rgb[j++] = px[i + 1]; rgb[j++] = px[i + 2]; }
    return { w: c.width, h: c.height, kind: 'raw', data: rgb };
  }

  async function loadImage(dataUrl) {
    const bytes = dataUrlBytes(dataUrl);
    if (bytes[0] === 0xFF && bytes[1] === 0xD8) {
      const j = jpegInfo(bytes);
      if (j.comps !== 1 && j.comps !== 3) return rasterize(dataUrl);
      return { w: j.w, h: j.h, kind: 'jpeg', comps: j.comps, data: bytes };
    }
    const p = pngInfo(bytes);
    if (p && !p.interlace && !p.trns && p.depth <= 8 && ((p.ctype === 3 && p.plte) || (p.ctype === 2 && p.depth === 8) || (p.ctype === 0))) {
      let total = 0; p.idat.forEach(c => { total += c.length; });
      const data = new Uint8Array(total); let o = 0;
      p.idat.forEach(c => { data.set(c, o); o += c.length; });
      return { w: p.w, h: p.h, kind: 'png', ctype: p.ctype, depth: p.depth, plte: p.plte, data };
    }
    return rasterize(dataUrl);
  }

  /* ---------------------------------------------------------------------
     Penulis PDF (objek, aliran, xref)
     --------------------------------------------------------------------- */
  class Pdf {
    constructor() { this.objs = []; }
    reserve() { this.objs.push(null); return this.objs.length; }
    set(id, chunks) { this.objs[id - 1] = chunks; }
    add(chunks) { this.objs.push(chunks); return this.objs.length; }
    // dict tanpa << >> ; data = Uint8Array; compress = coba Flate
    async stream(dict, data, compress) {
      let body = data, filter = '';
      if (compress) { const z = await deflate(data); if (z && z.length < data.length) { body = z; filter = ' /Filter /FlateDecode'; } }
      return [`<< ${dict}${filter} /Length ${body.length} >>\nstream\n`, body, '\nendstream'];
    }
    bytes(rootId, infoId) {
      const parts = [new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2D, 0x31, 0x2E, 0x34, 0x0A, 0x25, 0xE2, 0xE3, 0xCF, 0xD3, 0x0A])];
      let pos = parts[0].length;
      const offs = [];
      const push = c => { const u = typeof c === 'string' ? enc.encode(c) : c; parts.push(u); pos += u.length; };
      this.objs.forEach((chunks, i) => {
        if (!chunks) throw new Error('Objek PDF ' + (i + 1) + ' belum diisi.');
        offs.push(pos);
        push(`${i + 1} 0 obj\n`); chunks.forEach(push); push('\nendobj\n');
      });
      const xref = pos;
      let x = `xref\n0 ${this.objs.length + 1}\n0000000000 65535 f \n`;
      offs.forEach(o => { x += String(o).padStart(10, '0') + ' 00000 n \n'; });
      push(x + `trailer\n<< /Size ${this.objs.length + 1} /Root ${rootId} 0 R /Info ${infoId} 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
      const out = new Uint8Array(pos); let o = 0;
      parts.forEach(u => { out.set(u, o); o += u.length; });
      return out;
    }
  }

  const num = v => (Math.round(v * 100) / 100).toString();
  const utf16hex = s => { let o = 'FEFF'; for (let i = 0; i < s.length; i++) o += s.charCodeAt(i).toString(16).padStart(4, '0'); return '<' + o + '>'; };

  /* Halaman: perintah gambar dalam koordinat "atas-kiri" (y turun); diubah ke koordinat PDF saat ditulis. */
  class Page {
    constructor(faces) { this.ops = []; this.faces = faces; }
    text(s, x, base, face, size) {
      if (!s) return;
      this.ops.push(`BT /F${face.id} ${num(size)} Tf 1 0 0 1 ${num(x)} ${num(PAGE_H - base)} Tm ${face.hex(s)} Tj ET`);
    }
    line(x1, y1, x2, y2, lw) { this.ops.push(`${num(lw)} w ${num(x1)} ${num(PAGE_H - y1)} m ${num(x2)} ${num(PAGE_H - y2)} l S`); }
    rect(x, y, w, h, lw) { this.ops.push(`${num(lw)} w ${num(x)} ${num(PAGE_H - y - h)} ${num(w)} ${num(h)} re S`); }
    image(name, x, y, w, h) { this.ops.push(`q ${num(w)} 0 0 ${num(h)} ${num(x)} ${num(PAGE_H - y - h)} cm /${name} Do Q`); }
  }

  /* ---------------------------------------------------------------------
     Alat teks
     --------------------------------------------------------------------- */
  // "KOMIDA CABANG CIREBON" / "CAB. CIREBON" -> "CIREBON" (kalimat surat sudah memuat "KOMIDA Cabang")
  const branchName = v => String(v == null ? '' : v).replace(/\s+/g, ' ').trim().replace(/^(KOMIDA\s+)?(CABANG\s+|CAB\.?\s+)?/i, '');
  const clean = v => String(v == null ? '' : v).replace(/[\u0000-\u001F\u007F]/g, ' ').replace(/\s+/g, ' ').trim();
  const proper = v => clean(v).toLowerCase().replace(/(^|[\s\-.\/(])([a-z])/g, (m, a, b) => a + b.toUpperCase());
  function rupiah(v) {
    const n = Math.round((Number(v) || 0) * 100) / 100, neg = n < 0, a = Math.abs(n);
    const int = Math.floor(a), frac = Math.round((a - int) * 100);
    const s = String(int).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return (neg ? '-' : '') + s + (frac ? ',' + String(frac).padStart(2, '0') : '');
  }

  // Pecah teks menjadi baris selama muat dalam maxW. Kata yang lebih panjang dari lebar kolom dipotong per huruf.
  function wrap(text, face, size, maxW) {
    const words = clean(text).split(' ').filter(Boolean), lines = [];
    let cur = '';
    const push = w => {
      if (face.width(w, size) <= maxW) { cur = w; return; }
      let part = '';
      for (const ch of w) {
        if (part && face.width(part + ch, size) > maxW) { lines.push(part); part = ch; } else part += ch;
      }
      cur = part;
    };
    for (const w of words) {
      if (!cur) { push(w); continue; }
      if (face.width(cur + ' ' + w, size) <= maxW) cur += ' ' + w;
      else { lines.push(cur); cur = ''; push(w); }
    }
    if (cur || !lines.length) lines.push(cur);
    return lines;
  }

  /* ---------------------------------------------------------------------
     Penyusunan surat
     --------------------------------------------------------------------- */
  function compose(L, F, kop, sig) {
    const pages = [];
    let pg = null, y = 0;
    const newPage = () => { pg = new Page(F); pages.push(pg); pg.image('Kop', kop.x, kop.y, kop.w, kop.h); y = TOP; };
    const room = h => y + h <= BOTTOM;
    const need = h => { if (!room(h)) newPage(); };
    const base = (top, lead, size) => top + lead / 2 + size * 0.33;          // garis dasar agar teks berada di tengah baris
    const centerText = (s, cx, top, lead, face, size) => pg.text(s, cx - face.width(s, size) / 2, base(top, lead, size), face, size);

    const jenisUp = L.jenis.toUpperCase();
    newPage();

    // Judul
    const t1 = 'BERITA ACARA';
    centerText(t1, ML + TW / 2, y, 22, F.bold, 14);
    const tb = base(y, 22, 14), tw1 = F.bold.width(t1, 14);
    pg.line(ML + TW / 2 - tw1 / 2, tb + 2, ML + TW / 2 + tw1 / 2, tb + 2, 0.8);
    y += 22;
    centerText('TRANSAKSI WRITTE OFF (WO) ' + jenisUp + ' PERIODE ' + L.y, ML + TW / 2, y, 18, F.bold, 10.5);
    y += 18 + 10;

    // Paragraf pembuka (rata kiri-kanan; nilai isian dicetak tebal)
    const N = F.normal, B = F.bold;
    const tok = [['Pada hari ini ', N], [L.hari, B], [' tanggal, ', N], [L.tanggal, B], [' KOMIDA Cabang ', N], [branchName(L.b), B],
      [' telah melakukan eksekusi transaksi anggota Written Off (WO) ' + L.jenis + ' periode tahun ' + L.y + ', adapun data anggota yang kami lakukan Written Off (WO) sebagai berikut :', N]];
    // Pemisahan kata dengan mempertahankan gaya; potongan yang tidak diawali spasi menyambung ke kata sebelumnya
    const words = [];
    let glue = false;
    tok.forEach(([t, face]) => {
      const lead = /^\s/.test(t), trail = /\s$/.test(t);
      t.split(/\s+/).filter(Boolean).forEach((w, i) => words.push({ w: w, face: face, join: i === 0 && glue && !lead }));
      glue = !trail;
    });
    const lines = [];
    let cur = [], curW = 0;
    const sp = N.width(' ', BODY);
    words.forEach(it => {
      const w = it.face.width(it.w, BODY), gap = cur.length ? (it.join ? 0 : sp) : 0;
      if (cur.length && curW + gap + w > TW) { lines.push(cur); cur = []; curW = 0; }
      const g = cur.length ? (it.join ? 0 : sp) : 0;
      cur.push(Object.assign({ width: w, gap: g }, it)); curW += g + w;
    });
    if (cur.length) lines.push(cur);
    lines.forEach((ln, li) => {
      need(BODY_LEAD);
      const last = li === lines.length - 1, nGaps = ln.filter((x, i) => i > 0 && x.gap > 0).length;
      const natural = ln.reduce((a, x) => a + x.width + x.gap, 0), extra = !last && nGaps ? (TW - natural) / nGaps : 0;
      let x = ML;
      ln.forEach((it, i) => {
        if (i > 0) x += it.gap > 0 ? it.gap + extra : 0;
        pg.text(it.w, x, base(y, BODY_LEAD, BODY), it.face, BODY);
        x += it.width;
      });
      y += BODY_LEAD;
    });
    y += 4;

    // Tabel
    // Kolom No Voucher dilebarkan (+/- 108 poin) supaya satu nomor voucher (mis. THC-060-26-11-013667) muat dalam satu baris
    const W1 = [24, 104, 48, 38, 70, 72]; W1.push(TW - W1.reduce((a, b) => a + b, 0));
    const HEAD = ['No', 'Nama Anggota', 'Client ID', 'Ctr ID', 'Product Name', 'Total WO', 'No Voucher'];
    const ALIGN = ['c', 'l', 'c', 'c', 'l', 'r', 'l'];
    const heading = (no, text) => { need(16 + 80); pg.text(no, ML + 18, base(y, 16, BODY), B, BODY); pg.text(text, ML + 36, base(y, 16, BODY), B, BODY); y += 16 + 4; };

    // '\n' = pindah baris yang disengaja (dipakai kolom No Voucher: satu nomor voucher per baris)
    const cellBlock = (s, w, face, size) => String(s).split('\n').reduce((a, part) => a.concat(wrap(part, face, size, w - 2 * PAD_X)), []);
    const drawCell = (x, top, w, h, lines, face, size, align) => {
      size = Math.min(size, ...lines.map(s => (s ? (w - 2 * PAD_X) / Math.max(face.width(s, 1), 1e-6) : size)));
      pg.rect(x, top, w, h, 0.5);
      const bh = lines.length * CELL_LEAD;
      let ty = top + (h - bh) / 2;
      lines.forEach(s => {
        const tw = face.width(s, size);
        const tx = align === 'c' ? x + (w - tw) / 2 : align === 'r' ? x + w - PAD_X - tw : x + PAD_X;
        pg.text(s, tx, base(ty, CELL_LEAD, size), face, size);
        ty += CELL_LEAD;
      });
    };
    const rowHeight = (cells, widths, face) => Math.max(...cells.map((c, i) => cellBlock(c, widths[i], face, CELL).length)) * CELL_LEAD + 2 * PAD_Y;

    function table(cols, rows, totalValue, withVoucher, emptyText) {
      const widths = W1.slice(0, cols), heads = HEAD.slice(0, cols);
      const hh = rowHeight(heads, widths, B);
      const drawHead = () => {
        let x = ML;
        heads.forEach((t, i) => { drawCell(x, y, widths[i], hh, cellBlock(t, widths[i], B, CELL), B, CELL, 'c'); x += widths[i]; });
        y += hh;
      };
      const totalH = CELL_LEAD + 2 * PAD_Y;
      const firstH = rows.length ? rowHeight(rows[0].map(String), widths, N) : totalH;
      need(hh + firstH + (rows.length === 1 ? totalH : 0));
      drawHead();
      if (!rows.length) {
        const w = widths.reduce((a, b) => a + b, 0);
        drawCell(ML, y, w, totalH, [emptyText], F.italic, CELL, 'c'); y += totalH;
        return;
      }
      rows.forEach((r, ri) => {
        const rh = rowHeight(r, widths, N), last = ri === rows.length - 1;
        if (!room(rh + (last ? totalH : 0))) { newPage(); drawHead(); }
        let x = ML;
        r.forEach((t, i) => { drawCell(x, y, widths[i], rh, cellBlock(t, widths[i], N, CELL), N, CELL, ALIGN[i]); x += widths[i]; });
        y += rh;
      });
      // baris Total: label menyatu di lima kolom pertama
      const labelW = widths.slice(0, 5).reduce((a, b) => a + b, 0);
      drawCell(ML, y, labelW, totalH, ['Total'], B, CELL, 'c');
      const tv = rupiah(totalValue), tFit = Math.min(CELL, (widths[5] - 2 * PAD_X) / B.width(tv, 1));
      drawCell(ML + labelW, y, widths[5], totalH, [tv], B, tFit, 'r');
      if (withVoucher) drawCell(ML + labelW + widths[5], y, widths[6], totalH, [''], N, CELL, 'l');
      y += totalH;
    }

    // "THC-..., TAK-..." (dipisah koma / titik koma / baris baru) -> satu nomor voucher per baris
    const voucherLines = v => String(v == null ? '' : v).split(/[,;\r\n]+/).map(clean).filter(Boolean).join('\n');

    heading('1.', 'Yang berhasil di proses:');
    table(7, L.active.map((r, i) => [String(i + 1), r[0], r[1], r[2], r[3], rupiah(r[4]), voucherLines(r[5])]), L.totalActive, true, '');
    y += 16;
    heading('2.', 'Yang tidak dapat di proses:');
    table(6, L.batal.map((r, i) => [String(i + 1), r[0], r[1], r[2], r[3], rupiah(r[4])]), L.totalBatal, false, 'Tidak ada data yang tidak dapat diproses');
    y += 18;

    // Penutup + tanda tangan + tembusan: satu blok, tidak dipecah antar halaman
    const colW = TW / 2, nameLines = who => wrap(who.nama, B, BODY, colW - 12);
    const nl = Math.max(nameLines(L.msa).length, nameLines(L.bm).length);
    const blockH = BODY_LEAD + 6 + BODY_LEAD + 6 + BODY_LEAD + 4 + SIG_BOX_H + 4 + nl * BODY_LEAD + 13 + 14 + BODY_LEAD + 3 * 14 + 4;
    need(blockH);
    pg.text('Demikian Berita Acara ini dibuat agar dapat dipergunakan sebagaimana mestinya.', ML, base(y, BODY_LEAD, BODY), N, BODY);
    y += BODY_LEAD + 6;
    pg.text(proper(branchName(L.b)) + ', ' + L.tanggal, ML, base(y, BODY_LEAD, BODY), N, BODY);
    y += BODY_LEAD + 6;
    const cols2 = [
      { cx: ML + colW / 2, label: 'Yang melakukan penyesuaian,', who: L.msa, sig: sig.msa, jab: 'Mis Support Administration Staff' },
      { cx: ML + colW * 1.5, label: 'Mengetahui,', who: L.bm, sig: sig.bm, jab: 'Branch Manager' }];
    cols2.forEach(c => centerText(c.label, c.cx, y, BODY_LEAD, N, BODY));
    y += BODY_LEAD + 4;
    cols2.forEach(c => {
      if (!c.sig) return;
      const k = Math.min(SIG_BOX_W / c.sig.w, SIG_BOX_H / c.sig.h), w = c.sig.w * k, h = c.sig.h * k;
      pg.image(c.sig.name, c.cx - w / 2, y + (SIG_BOX_H - h) / 2, w, h);
    });
    y += SIG_BOX_H + 4;
    cols2.forEach(c => {
      nameLines(c.who).forEach((s, i) => {
        const top = y + i * BODY_LEAD, tw = B.width(s, BODY), b = base(top, BODY_LEAD, BODY);
        pg.text(s, c.cx - tw / 2, b, B, BODY);
        pg.line(c.cx - tw / 2, b + 1.8, c.cx + tw / 2, b + 1.8, 0.7);
      });
      centerText(c.jab, c.cx, y + nl * BODY_LEAD, 13, F.italic, 9);
    });
    y += nl * BODY_LEAD + 13 + 14;
    pg.text('Tembusan:', ML, base(y, BODY_LEAD, BODY), N, BODY);
    y += BODY_LEAD;
    ['Koordinator Wilayah', 'Manager Regional', 'Koordinator Cabang'].forEach((s, i) => {
      pg.text((i + 1) + '.', ML + 18, base(y, 14, BODY), N, BODY);
      pg.text(s, ML + 36, base(y, 14, BODY), N, BODY);
      y += 14;
    });
    return pages;
  }

  /* ---------------------------------------------------------------------
     Rakit PDF
     --------------------------------------------------------------------- */
  async function build(letter, opts) {
    const o = opts || {}, kopSrc = o.kop || root.SURAT_KOP, fontSrc = o.fonts || root.SURAT_FONTS;
    if (!kopSrc || !kopSrc.data) throw new Error('Gambar kop surat (surat-kop.js) belum dimuat.');
    if (!fontSrc || !fontSrc.files || !fontSrc.files.normal || !fontSrc.files.bold || !fontSrc.files.italic) throw new Error('Font surat (surat-font.js) belum dimuat.');
    const L = letter;
    if (!L || !Array.isArray(L.active) || !Array.isArray(L.batal)) throw new Error('Isi surat tidak lengkap.');

    const fam = String(fontSrc.family || 'LiberationSerif').replace(/[^A-Za-z0-9]/g, '');
    const faces = {
      normal: new Face(fam, b64ToBytes(fontSrc.files.normal), 34, 80),
      bold: new Face(fam + '-Bold', b64ToBytes(fontSrc.files.bold), 34, 140),
      italic: new Face(fam + '-Italic', b64ToBytes(fontSrc.files.italic), 98, 80)
    };
    ['normal', 'bold', 'italic'].forEach((k, i) => { faces[k].id = i + 1; });

    const pdf = new Pdf();
    const catalog = pdf.reserve(), pagesId = pdf.reserve(), info = pdf.reserve();

    // Gambar: kop + TTD
    const images = [];
    const addImage = async (name, dataUrl) => {
      const im = await loadImage(dataUrl);
      let dict = `/Type /XObject /Subtype /Image /Width ${im.w} /Height ${im.h}`, compress = false;
      if (im.kind === 'jpeg') { dict += ` /ColorSpace ${im.comps === 1 ? '/DeviceGray' : '/DeviceRGB'} /BitsPerComponent 8 /Filter /DCTDecode`; }
      else if (im.kind === 'png') {
        const colors = im.ctype === 2 ? 3 : 1;
        let cs = im.ctype === 2 ? '/DeviceRGB' : '/DeviceGray';
        if (im.ctype === 3) {
          let hex = ''; im.plte.forEach(v => { hex += v.toString(16).padStart(2, '0'); });
          cs = `[/Indexed /DeviceRGB ${im.plte.length / 3 - 1} <${hex}>]`;
        }
        dict += ` /ColorSpace ${cs} /BitsPerComponent ${im.depth} /Filter /FlateDecode /DecodeParms << /Predictor 15 /Colors ${colors} /BitsPerComponent ${im.depth} /Columns ${im.w} >>`;
      } else { dict += ' /ColorSpace /DeviceRGB /BitsPerComponent 8'; compress = true; }
      const id = pdf.add(await pdf.stream(dict, im.data, compress));
      images.push({ name: name, id: id });
      return { name: name, w: im.w, h: im.h };
    };
    await addImage('Kop', kopSrc.data);
    const sig = { msa: null, bm: null };
    if (L.msa && L.msa.img) sig.msa = await addImage('SigMsa', L.msa.img);
    if (L.bm && L.bm.img) sig.bm = await addImage('SigBm', L.bm.img);

    // Isi halaman (menyusun surat sekaligus mencatat huruf yang dipakai)
    // Kop dilebarkan tepat selebar area isi (margin kiri = margin kanan) supaya garis kop sejajar dengan teks surat dan simetris.
    // Gambar kop memenuhi seluruh lebarnya, jadi cukup x = ML dan w = TW; tinggi mengikuti rasio asli gambar.
    const kopW = TW, kopH = kopW * kopSrc.h / kopSrc.w;
    const pages = compose(L, faces, { x: ML, y: kopSrc.y, w: kopW, h: kopH }, sig);

    // Font
    const fontRefs = {};
    for (const key of ['normal', 'bold', 'italic']) {
      const fc = faces[key], f = fc.f;
      const fileId = pdf.add(await pdf.stream(`/Length1 ${f.bytes.length}`, f.bytes, true));
      const desc = pdf.add([`<< /Type /FontDescriptor /FontName /${fc.name} /Flags ${fc.flags} /FontBBox [${f.bbox.map(v => Math.round(v * 1000 / f.upm)).join(' ')}] `
        + `/ItalicAngle ${num(f.italicAngle)} /Ascent ${Math.round(f.ascent * 1000 / f.upm)} /Descent ${Math.round(f.descent * 1000 / f.upm)} `
        + `/CapHeight ${Math.round(f.capHeight * 1000 / f.upm)} /StemV ${fc.stemV} /FontFile2 ${fileId} 0 R >>`]);
      const widths = f.adv.map(a => Math.round(a * 1000 / f.upm)).join(' ');
      const cid = pdf.add([`<< /Type /Font /Subtype /CIDFontType2 /BaseFont /${fc.name} /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> `
        + `/FontDescriptor ${desc} 0 R /CIDToGIDMap /Identity /DW 1000 /W [0 [${widths}]] >>`]);
      const entries = [...fc.used.entries()].sort((a, b) => a[0] - b[0]);
      let cmap = '/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n'
        + '1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n';
      for (let i = 0; i < entries.length; i += 100) {
        const part = entries.slice(i, i + 100);
        cmap += part.length + ' beginbfchar\n' + part.map(([g, cp]) => '<' + g.toString(16).padStart(4, '0') + '> <' + cp.toString(16).padStart(4, '0') + '>').join('\n') + '\nendbfchar\n';
      }
      cmap += 'endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend';
      const tu = pdf.add(await pdf.stream('', enc.encode(cmap), true));
      fontRefs[key] = pdf.add([`<< /Type /Font /Subtype /Type0 /BaseFont /${fc.name} /Encoding /Identity-H /DescendantFonts [${cid} 0 R] /ToUnicode ${tu} 0 R >>`]);
    }

    // Halaman
    const res = `<< /Font << /F1 ${fontRefs.normal} 0 R /F2 ${fontRefs.bold} 0 R /F3 ${fontRefs.italic} 0 R >> /XObject << ${images.map(i => `/${i.name} ${i.id} 0 R`).join(' ')} >> >>`;
    const pageIds = [];
    for (const pg of pages) {
      const content = pdf.add(await pdf.stream('', enc.encode(pg.ops.join('\n')), true));
      pageIds.push(pdf.add([`<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${num(PAGE_W)} ${num(PAGE_H)}] /Resources ${res} /Contents ${content} 0 R >>`]));
    }
    pdf.set(pagesId, [`<< /Type /Pages /Kids [${pageIds.map(i => i + ' 0 R').join(' ')}] /Count ${pageIds.length} >>`]);
    pdf.set(catalog, [`<< /Type /Catalog /Pages ${pagesId} 0 R >>`]);
    const d = new Date(), z = n => String(n).padStart(2, '0');
    const stamp = `D:${d.getUTCFullYear()}${z(d.getUTCMonth() + 1)}${z(d.getUTCDate())}${z(d.getUTCHours())}${z(d.getUTCMinutes())}${z(d.getUTCSeconds())}Z`;
    pdf.set(info, [`<< /Title ${utf16hex('Berita Acara WO ' + L.jenis + ' ' + L.y + ' Cabang ' + clean(L.b))} /Subject ${utf16hex(L.id)} /Creator ${utf16hex('Dashboard Operasional Satu Pintu')} /CreationDate (${stamp}) >>`]);
    const bytes = pdf.bytes(catalog, info);
    return { bytes: bytes, pages: pages.length, blob: () => new Blob([bytes], { type: 'application/pdf' }) };
  }

  root.SuratWO = { build: build, rupiah: rupiah, proper: proper };
})(typeof window !== 'undefined' ? window : globalThis);
