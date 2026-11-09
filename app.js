/* Tracking Perpindahan Barang - plain JavaScript (tanpa build step) */
(() => {
  'use strict';

  // ---------- Util ----------
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmtNum = (n) => Number(n).toLocaleString('id-ID', { maximumFractionDigits: 2 });
  const todayStr = () => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  };
  const fmtDate = (s) => { const [y, m, d] = s.split('-'); return `${d}/${m}/${y}`; };
  let toastTimer;
  function toast(msg, isErr) {
    const t = $('toast');
    t.textContent = msg; t.className = 'toast' + (isErr ? ' err' : ''); t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(() => (t.hidden = true), 3500);
  }

  // ---------- Penyimpanan ----------
  // Data stok (besar, 20rb+ baris) -> IndexedDB. Riwayat transaksi -> localStorage.
  const DB_NAME = 'wms-move', STORE = 'kv';
  function idb() {
    return new Promise((res, rej) => {
      const r = indexedDB.open(DB_NAME, 1);
      r.onupgradeneeded = () => r.result.createObjectStore(STORE);
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  }
  async function idbGet(key) {
    const db = await idb();
    return new Promise((res, rej) => { const q = db.transaction(STORE).objectStore(STORE).get(key); q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error); });
  }
  async function idbSet(key, val) {
    const db = await idb();
    return new Promise((res, rej) => { const tx = db.transaction(STORE, 'readwrite'); tx.objectStore(STORE).put(val, key); tx.oncomplete = res; tx.onerror = () => rej(tx.error); });
  }
  const HIST_KEY = 'wms_move_history_v1';
  const loadHistory = () => { try { return JSON.parse(localStorage.getItem(HIST_KEY)) || []; } catch { return []; } };
  const saveHistory = () => localStorage.setItem(HIST_KEY, JSON.stringify(history));

  // ---------- State ----------
  let stock = [];            // [{lot, code, motif, container, loc, qty, uom}]
  let lotMap = new Map();    // lot -> row
  let locations = [];        // daftar lokasi unik
  let stockMeta = null;      // {name, count, at}
  let history = loadHistory();
  let current = null;        // hasil generate {groups, missing}
  let selectedDest = '';

  // ---------- Upload & parse Excel ----------
  const COLS = {
    lot: 'Lot/Serial Number', code: 'Product Code', motif: 'Product/Name', container: 'Nomor Container',
    loc: 'Location', qty: 'Quantity', uom: 'Unit of Measure'
  };
  function indexStock(rows) {
    stock = rows;
    lotMap = new Map(rows.map((r) => [r.lot, r]));
    locations = [...new Set(rows.map((r) => r.loc).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  }
  function renderStatus() {
    const el = $('dataStatus');
    if (stockMeta) {
      el.className = 'chip chip-ok';
      el.textContent = `${stockMeta.name} · ${stockMeta.count.toLocaleString('id-ID')} lot · ${locations.length} lokasi`;
    } else { el.className = 'chip chip-warn'; el.textContent = 'Belum ada data stok'; }
  }
  $('fileInput').addEventListener('change', async (e) => {
    const f = e.target.files[0]; if (!f) return;
    try {
      if (typeof XLSX === 'undefined') throw new Error('Library Excel gagal dimuat. Periksa koneksi internet.');
      toast('Membaca file...');
      const wb = XLSX.read(await f.arrayBuffer(), { type: 'array', cellDates: true });
      const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '', raw: true });
      if (!json.length) throw new Error('File kosong.');
      const missingCols = Object.values(COLS).filter((c) => !(c in json[0]));
      if (missingCols.length) throw new Error('Kolom tidak ditemukan: ' + missingCols.join(', '));
      const rows = [];
      for (const r of json) {
        const lot = String(r[COLS.lot]).trim();
        if (!lot) continue;
        rows.push({
          lot, code: String(r[COLS.code]).trim(), motif: String(r[COLS.motif]).trim(),
          container: String(r[COLS.container]).trim(), loc: String(r[COLS.loc]).trim(),
          qty: Number(r[COLS.qty]) || 0, uom: String(r[COLS.uom]).trim()
        });
      }
      indexStock(rows);
      stockMeta = { name: f.name, count: rows.length, at: Date.now() };
      await idbSet('stock', { rows, meta: stockMeta });
      renderStatus(); resetResult();
      toast(`Berhasil: ${rows.length.toLocaleString('id-ID')} lot dimuat.`);
    } catch (err) { toast(err.message || 'Gagal membaca file', true); }
    e.target.value = '';
  });

  // ---------- Parse lot ----------
  function parseLots(text) {
    return [...new Set(text.split(/[\s,;]+/).map((s) => s.trim()).filter(Boolean))];
  }
  $('lotInput').addEventListener('input', () => {
    const n = parseLots($('lotInput').value).length;
    $('lotCount').textContent = `${n} lot terdeteksi`;
  });
  $('btnClear').addEventListener('click', () => { $('lotInput').value = ''; $('lotCount').textContent = '0 lot terdeteksi'; resetResult(); });

  // ---------- Generate ----------
  function resetResult() {
    current = null; selectedDest = '';
    $('resultCard').hidden = true; $('saveCard').hidden = true;
    $('destInput').value = ''; $('docInput').value = ''; $('destMsg').textContent = '';
  }
  $('btnGenerate').addEventListener('click', () => {
    if (!stock.length) return toast('Upload file stok terlebih dahulu.', true);
    const lots = parseLots($('lotInput').value);
    if (!lots.length) return toast('Masukkan minimal satu lot / serial number.', true);

    const groups = new Map(); const missing = [];
    for (const lot of lots) {
      const r = lotMap.get(lot);
      if (!r) { missing.push(lot); continue; }
      const key = [r.motif, r.container, r.loc].join('||');
      if (!groups.has(key)) groups.set(key, { motif: r.motif, code: r.code, container: r.container, from: r.loc, uom: r.uom, rolls: 0, qty: 0, lots: [] });
      const g = groups.get(key); g.rolls++; g.qty += r.qty; g.lots.push(lot);
    }
    current = { groups: [...groups.values()].sort((a, b) => a.motif.localeCompare(b.motif) || a.container.localeCompare(b.container)), missing };
    renderResult();
  });

  function renderResult() {
    const { groups, missing } = current;
    $('resultCard').hidden = false;
    $('saveCard').hidden = !groups.length;
    const mb = $('missingBox');
    if (missing.length) {
      mb.hidden = false;
      mb.innerHTML = `<b>${missing.length} lot tidak ditemukan</b> di data stok dan tidak akan disimpan:<div class="lotchips">${missing.map((l) => `<span>${esc(l)}</span>`).join('')}</div>`;
    } else mb.hidden = true;

    const body = $('resultTable').querySelector('tbody');
    body.innerHTML = groups.length ? groups.map((g) => `
      <tr>
        <td><b>${esc(g.motif)}</b><div class="muted">${esc(g.code)}</div>
          <details><summary>Lihat ${g.lots.length} lot</summary><div class="lots">${g.lots.map(esc).join(', ')}</div></details></td>
        <td>${esc(g.container) || '-'}</td>
        <td>${esc(g.from)}</td>
        <td class="num">${g.rolls}</td>
        <td class="num">${fmtNum(g.qty)} ${esc(g.uom)}</td>
      </tr>`).join('') : '<tr><td colspan="5" class="empty">Tidak ada lot yang cocok.</td></tr>';

    const totalRolls = groups.reduce((s, g) => s + g.rolls, 0);
    $('resultTable').querySelector('tfoot').innerHTML = groups.length ?
      `<tr><td colspan="3">Total</td><td class="num">${totalRolls}</td><td></td></tr>` : '';
  }

  // ---------- Combobox lokasi tujuan ----------
  const dest = $('destInput'), list = $('destList');
  let activeIdx = -1, shown = [];
  function renderCombo() {
    const q = dest.value.trim().toLowerCase();
    const terms = q.split(/\s+/).filter(Boolean);
    const all = terms.length ? locations.filter((l) => { const x = l.toLowerCase(); return terms.every((t) => x.includes(t)); }) : locations;
    shown = all.slice(0, 60); activeIdx = shown.length ? 0 : -1;
    if (!shown.length) list.innerHTML = '<div class="combo-empty">Lokasi tidak ditemukan.</div>';
    else {
      const hl = (l) => { let h = esc(l); terms.forEach((t) => { h = h.replace(new RegExp('(' + t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'ig'), '<mark>$1</mark>'); }); return h; };
      list.innerHTML = shown.map((l, i) => `<div class="combo-item${i === 0 ? ' active' : ''}" data-i="${i}">${terms.length ? hl(l) : esc(l)}</div>`).join('') +
        (all.length > shown.length ? `<div class="combo-empty">+${all.length - shown.length} lokasi lain, persempit pencarian...</div>` : '');
    }
    list.hidden = false;
  }
  function pick(i) {
    if (i < 0 || !shown[i]) return;
    selectedDest = shown[i]; dest.value = selectedDest; list.hidden = true;
    $('destMsg').textContent = '';
  }
  function moveActive(d) {
    const items = list.querySelectorAll('.combo-item'); if (!items.length) return;
    items[activeIdx]?.classList.remove('active');
    activeIdx = (activeIdx + d + items.length) % items.length;
    items[activeIdx].classList.add('active'); items[activeIdx].scrollIntoView({ block: 'nearest' });
  }
  dest.addEventListener('focus', () => { if (locations.length) renderCombo(); });
  dest.addEventListener('input', () => { selectedDest = ''; renderCombo(); });
  dest.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown') { e.preventDefault(); list.hidden ? renderCombo() : moveActive(1); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); moveActive(-1); }
    else if (e.key === 'Enter') { if (!list.hidden) { e.preventDefault(); pick(activeIdx); } }
    else if (e.key === 'Escape') list.hidden = true;
  });
  list.addEventListener('mousedown', (e) => { const it = e.target.closest('.combo-item'); if (it) { e.preventDefault(); pick(+it.dataset.i); } });
  document.addEventListener('click', (e) => { if (!e.target.closest('.combo')) list.hidden = true; });
  dest.addEventListener('blur', () => {
    setTimeout(() => {
      list.hidden = true;
      if (!selectedDest && dest.value.trim()) {
        const m = locations.find((l) => l.toLowerCase() === dest.value.trim().toLowerCase());
        if (m) { selectedDest = m; dest.value = m; }
        else $('destMsg').textContent = 'Pilih salah satu lokasi dari daftar.';
      }
    }, 120);
  });

  // ---------- Simpan ----------
  $('btnSave').addEventListener('click', () => {
    if (!current || !current.groups.length) return toast('Belum ada hasil generate.', true);
    if (!selectedDest) {
      const m = locations.find((l) => l.toLowerCase() === dest.value.trim().toLowerCase());
      if (m) { selectedDest = m; dest.value = m; } else return toast('Pilih lokasi tujuan dari daftar.', true);
    }
    const doc = $('docInput').value.trim();
    if (!doc) return toast('Isi source dokumen transaksi.', true);

    const lotsNow = current.groups.flatMap((g) => g.lots);
    const used = new Set(history.flatMap((t) => t.groups.flatMap((g) => g.lots)));
    const dupes = lotsNow.filter((l) => used.has(l));
    if (dupes.length && !confirm(`${dupes.length} lot sudah pernah masuk transaksi sebelumnya. Tetap simpan?`)) return;

    history.unshift({
      id: Date.now().toString(36), date: todayStr(), ts: Date.now(), dest: selectedDest, doc,
      groups: current.groups.map((g) => ({ ...g, lots: [...g.lots] }))
    });
    saveHistory();
    $('lotInput').value = ''; $('lotCount').textContent = '0 lot terdeteksi';
    resetResult(); renderHistory();
    toast('Transaksi tersimpan.');
  });

  // ---------- Riwayat ----------
  function renderHistory() {
    const q = $('histSearch').value.trim().toLowerCase();
    const items = history.filter((t) => !q || JSON.stringify(t).toLowerCase().includes(q));
    $('histList').innerHTML = items.length ? items.map((t) => {
      const rolls = t.groups.reduce((s, g) => s + g.rolls, 0);
      return `<div class="tx">
        <div class="tx-head">
          <div><b>${esc(t.doc)}</b> <span class="tx-meta">· ${fmtDate(t.date)} · ${new Date(t.ts).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })} · ${rolls} roll</span>
            <div class="tx-meta">Tujuan: <b>${esc(t.dest)}</b></div></div>
          <div class="row gap" style="margin:0">
            <button class="btn btn-outline btn-sm" data-pdf="${t.id}">PDF</button>
            <button class="btn btn-danger" data-del="${t.id}">Hapus</button>
          </div>
        </div>
        <div class="table-wrap"><table>
          <thead><tr><th>Motif</th><th>No. Container</th><th>Lokasi Awal</th><th class="num">Qty Roll</th></tr></thead>
          <tbody>${t.groups.map((g) => `<tr><td>${esc(g.motif)}<details><summary>${g.lots.length} lot</summary><div class="lots">${g.lots.map(esc).join(', ')}</div></details></td><td>${esc(g.container) || '-'}</td><td>${esc(g.from)}</td><td class="num">${g.rolls}</td></tr>`).join('')}</tbody>
        </table></div></div>`;
    }).join('') : '<div class="empty">Belum ada transaksi.</div>';
  }
  $('histSearch').addEventListener('input', renderHistory);
  $('histList').addEventListener('click', (e) => {
    const pdf = e.target.closest('[data-pdf]'), del = e.target.closest('[data-del]');
    if (pdf) { const t = history.find((x) => x.id === pdf.dataset.pdf); exportPDF([t], `Transaksi ${t.doc}`, `transaksi-${t.doc}`); }
    if (del && confirm('Hapus transaksi ini dari riwayat?')) { history = history.filter((x) => x.id !== del.dataset.del); saveHistory(); renderHistory(); }
  });

  // ---------- Export PDF ----------
  // Setiap transaksi dibuat sebagai blok tersendiri (header + tabel + total).
  function exportPDF(txs, subtitle, filename) {
    if (!txs.length) return toast('Tidak ada transaksi untuk diexport.', true);
    if (!window.jspdf) return toast('Library PDF gagal dimuat. Periksa koneksi internet.', true);
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
    const W = doc.internal.pageSize.getWidth(), H = doc.internal.pageSize.getHeight();
    const M = 40, BLOCK_MIN = 130; // sisa ruang minimum sebelum pindah halaman

    const sorted = [...txs].sort((a, b) => a.ts - b.ts);
    const grandRolls = sorted.reduce((s, t) => s + t.groups.reduce((x, g) => x + g.rolls, 0), 0);

    doc.setFont('helvetica', 'bold'); doc.setFontSize(15);
    doc.text('Laporan Perpindahan Barang', M, 40);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(10);
    doc.text(`${subtitle}  |  ${sorted.length} transaksi  |  ${grandRolls} roll`, M, 57);
    doc.text('Dicetak: ' + new Date().toLocaleString('id-ID'), W - M, 40, { align: 'right' });

    let y = 76;
    sorted.forEach((t, idx) => {
      if (y + BLOCK_MIN > H - 40) { doc.addPage(); y = 40; }

      // Header blok transaksi
      doc.setFillColor(232, 238, 252); doc.setDrawColor(31, 94, 255);
      doc.rect(M, y, W - 2 * M, 42, 'FD');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(20);
      doc.text(`Transaksi ${idx + 1}  -  Source Dokumen: ${t.doc}`, M + 8, y + 16);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5);
      const time = new Date(t.ts).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
      doc.text(`Tanggal: ${fmtDate(t.date)} ${time}`, M + 8, y + 32);
      doc.text(`Tujuan Lokasi Baru: ${t.dest}`, M + 190, y + 32);

      const rolls = t.groups.reduce((s, g) => s + g.rolls, 0);
      doc.autoTable({
        startY: y + 42,
        head: [['No', 'Nama Motif', 'No. Container', 'Lokasi Awal', 'Qty Roll']],
        body: t.groups.map((g, i) => [String(i + 1), g.motif, g.container || '-', g.from, String(g.rolls)]),
        foot: [['', '', '', 'Total Roll', String(rolls)]],
        styles: { fontSize: 8.5, cellPadding: 4, overflow: 'linebreak' },
        headStyles: { fillColor: [31, 94, 255] },
        footStyles: { fillColor: [245, 247, 251], textColor: 20, fontStyle: 'bold' },
        columnStyles: { 0: { cellWidth: 30 }, 2: { cellWidth: 90 }, 3: { cellWidth: 250 }, 4: { cellWidth: 55, halign: 'right' } },
        margin: { left: M, right: M, top: 40, bottom: 40 },
        showFoot: 'lastPage',
        rowPageBreak: 'avoid'
      });
      y = doc.lastAutoTable.finalY + 22;
    });

    const pages = doc.internal.getNumberOfPages();
    for (let p = 1; p <= pages; p++) {
      doc.setPage(p); doc.setFontSize(8); doc.setTextColor(110);
      doc.text(`Halaman ${p} / ${pages}`, W / 2, H - 18, { align: 'center' });
    }
    doc.save(filename + '.pdf');
  }
  $('btnExport').addEventListener('click', () => {
    const d = $('expDate').value || todayStr();
    exportPDF(history.filter((t) => t.date === d), `Tanggal: ${fmtDate(d)}`, `perpindahan-barang-${d}`);
  });
  $('btnExportAll').addEventListener('click', () => exportPDF(history, 'Semua riwayat transaksi', 'perpindahan-barang-semua'));

  // ---------- Tabs ----------
  document.querySelectorAll('.tab').forEach((b) => b.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach((x) => x.classList.toggle('active', x === b));
    $('tab-trx').hidden = b.dataset.tab !== 'trx';
    $('tab-hist').hidden = b.dataset.tab !== 'hist';
    if (b.dataset.tab === 'hist') renderHistory();
  }));

  // ---------- Init ----------
  (async () => {
    $('expDate').value = todayStr();
    try {
      const saved = await idbGet('stock');
      if (saved && saved.rows) { indexStock(saved.rows); stockMeta = saved.meta; }
    } catch (e) { console.warn(e); }
    renderStatus(); renderHistory();
  })();
})();
