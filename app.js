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
    clearTimeout(toastTimer); toastTimer = setTimeout(() => (t.hidden = true), 4000);
  }

  // ---------- Penyimpanan riwayat (localStorage) ----------
  const HIST_KEY = 'wms_move_history_v1';
  const loadHistory = () => { try { return JSON.parse(localStorage.getItem(HIST_KEY)) || []; } catch { return []; } };
  const saveHistory = () => localStorage.setItem(HIST_KEY, JSON.stringify(history));

  // ---------- State ----------
  let history = loadHistory();
  let pending = [];          // transaksi hasil baca PDF, belum disimpan
  const selected = new Set(); // id transaksi yang dipilih untuk export

  // ---------- Baca PDF "Location Transfer" ----------
  const PDFJS_WORKER = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  const DOC_RE = /^[A-Za-z]{2,}\/[A-Za-z]+\/{1,2}[\w/.\-]+$/;          // contoh: GDA/INT//26/11/17/47860
  const DATE_RE = /(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/; // MM/DD/YYYY HH:MM:SS
  // Batas kolom tabel sebagai pecahan lebar halaman: [produk, lot, container, source, destination, done, uom]
  const COL = [0.092, 0.225, 0.362, 0.470, 0.616, 0.798, 0.867];

  function lineJoin(items) {
    // gabungkan item teks per baris visual (atas ke bawah, kiri ke kanan)
    const s = [...items].sort((a, b) => b.y - a.y || a.x - b.x);
    const lines = [];
    for (const it of s) {
      const last = lines[lines.length - 1];
      if (last && Math.abs(last.y - it.y) <= 2.5) last.items.push(it); else lines.push({ y: it.y, items: [it] });
    }
    return lines.map((l) => l.items.sort((a, b) => a.x - b.x).map((i) => i.s).join(' ')).join(' ').replace(/\s+/g, ' ').trim();
  }

  async function parsePdf(file) {
    const pdfjs = window.pdfjsLib;
    if (!pdfjs) throw new Error('Library PDF gagal dimuat. Periksa koneksi internet.');
    pdfjs.GlobalWorkerOptions.workerSrc = PDFJS_WORKER;
    const pdf = await pdfjs.getDocument({ data: await file.arrayBuffer() }).promise;

    const info = { doc: '', dt: null, by: '', total: null };
    const rows = []; let failed = 0;

    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const W = page.getViewport({ scale: 1 }).width;
      const tc = await page.getTextContent();
      const items = tc.items.filter((i) => i.str.trim()).map((i) => ({ s: i.str.trim(), x: i.transform[4], y: i.transform[5] }));

      if (p === 1) {
        const strs = tc.items.map((i) => i.str.trim()).filter(Boolean);
        info.doc = strs.find((s) => DOC_RE.test(s)) || '';
        info.dt = (strs.map((s) => s.match(DATE_RE)).find(Boolean)) || null;
        const bi = strs.indexOf('Created By'); if (bi >= 0) info.by = strs[bi + 1] || '';
        const ti = strs.indexOf('Total Qty');
        if (ti >= 0) { const v = strs.slice(ti + 1, ti + 3).find((s) => /^[\d,]+(\.\d+)?$/.test(s)); if (v) info.total = parseFloat(v.replace(/,/g, '')); }
      }

      // baris tabel dimulai dari angka pada kolom "No."
      const nums = items.filter((i) => /^\d+$/.test(i.s) && i.x < COL[0] * W && i.y > 60).sort((a, b) => b.y - a.y);
      if (!nums.length) continue;
      const ys = nums.map((n) => n.y);
      const buckets = nums.map(() => []);
      for (const it of items) {
        if (it.y < 60 || it.y > ys[0] + 8) continue; // buang footer & kepala halaman
        let idx = -1;
        for (let i = 0; i < ys.length; i++) if (it.y <= ys[i] + 8) idx = i;
        if (idx >= 0) buckets[idx].push(it);
      }
      buckets.forEach((b) => {
        const col = (lo, hi) => lineJoin(b.filter((i) => i.x >= lo * W && i.x < hi * W));
        const product = col(COL[0], COL[1]);
        const lot = col(COL[1], COL[2]).replace(/\s+/g, '');
        const container = col(COL[2], COL[3]);
        const from = col(COL[3], COL[4]);
        const dest = col(COL[4], COL[5]);
        const done = col(COL[5], COL[6]);
        const uom = col(COL[6], 2);
        const pm = product.match(/^\[([^\]]+)\]\s*(.*)$/);
        const code = pm ? pm[1] : '';
        const motif = (pm ? pm[2] : product).replace(/\s*\(.*$/, '').trim();
        const qty = parseFloat(done.replace(/,/g, ''));
        if (!lot || !from || !dest || !motif || Number.isNaN(qty)) { failed++; return; }
        rows.push({ lot, code, motif, container, from, dest, qty, uom });
      });
    }
    return { info, rows, failed };
  }

  // ubah hasil baca PDF -> daftar transaksi (satu per tujuan)
  function buildTransactions(file, parsed) {
    const { info, rows, failed } = parsed;
    const warnings = [];
    let doc = info.doc;
    if (!doc) { doc = file.name.replace(/\.pdf$/i, ''); warnings.push('Nomor dokumen tidak terbaca, memakai nama file.'); }
    let date = todayStr(), ts = Date.now();
    if (info.dt) {
      const [, mo, d, y, h, mi, se] = info.dt;
      date = `${y}-${mo}-${d}`; ts = new Date(+y, +mo - 1, +d, +h, +mi, +se).getTime();
    } else warnings.push('Tanggal transfer tidak terbaca, memakai tanggal hari ini.');
    if (failed) warnings.push(`${failed} baris tidak bisa dibaca dan dilewati.`);
    const sum = rows.reduce((s, r) => s + r.qty, 0);
    if (info.total != null && Math.abs(sum - info.total) > 0.01) {
      warnings.push(`Total Qty di dokumen (${fmtNum(info.total)}) tidak sama dengan jumlah baris terbaca (${fmtNum(sum)}). Periksa hasilnya.`);
    }

    const byDest = new Map();
    for (const r of rows) { if (!byDest.has(r.dest)) byDest.set(r.dest, []); byDest.get(r.dest).push(r); }
    return [...byDest.entries()].map(([dest, rs]) => {
      const groups = new Map();
      for (const r of rs) {
        const key = [r.motif, r.container, r.from].join('||');
        if (!groups.has(key)) groups.set(key, { motif: r.motif, code: r.code, container: r.container, from: r.from, uom: r.uom, rolls: 0, qty: 0, lots: [], qtys: [] });
        const g = groups.get(key); g.rolls++; g.qty += r.qty; g.lots.push(r.lot); g.qtys.push(r.qty);
      }
      return {
        id: '', date, ts, dest, doc, by: info.by,
        groups: [...groups.values()].sort((a, b) => a.motif.localeCompare(b.motif) || a.container.localeCompare(b.container)),
        warnings, file: file.name
      };
    });
  }

  async function handleFiles(fileList) {
    const files = [...fileList].filter((f) => /\.pdf$/i.test(f.name) || f.type === 'application/pdf');
    if (!files.length) return toast('Pilih file PDF Location Transfer.', true);
    $('parseStatus').textContent = 'Membaca file...';
    const out = []; const errs = [];
    for (const f of files) {
      try {
        const parsed = await parsePdf(f);
        if (!parsed.rows.length) { errs.push(`${f.name}: tidak ada baris barang yang terbaca. Pastikan ini PDF "Location Transfer".`); continue; }
        out.push(...buildTransactions(f, parsed));
      } catch (e) { errs.push(`${f.name}: ${e.message || 'gagal dibaca'}`); }
    }
    $('parseStatus').textContent = '';
    if (errs.length) toast(errs.join(' | '), true);
    if (out.length) { pending = out; renderPending(); }
  }

  function renderPending() {
    const card = $('pendingCard');
    card.hidden = !pending.length;
    if (!pending.length) return;
    const totalRolls = pending.reduce((s, t) => s + t.groups.reduce((x, g) => x + g.rolls, 0), 0);
    $('pendingSummary').textContent = `${pending.length} transaksi siap disimpan, ${totalRolls} roll`;
    $('pendingList').innerHTML = pending.map((t) => {
      const rolls = t.groups.reduce((s, g) => s + g.rolls, 0);
      const qty = {}; t.groups.forEach((g) => { qty[g.uom] = (qty[g.uom] || 0) + g.qty; });
      const qtyTxt = Object.entries(qty).map(([u, v]) => `${fmtNum(v)} ${u}`).join(' + ');
      const froms = [...new Set(t.groups.map((g) => g.from))];
      return `<div class="tx">
        <div class="tx-head">
          <div><b>${esc(t.doc)}</b> <span class="tx-meta">· ${fmtDate(t.date)} ${new Date(t.ts).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })}${t.by ? ' · ' + esc(t.by) : ''} · ${rolls} roll · ${qtyTxt}</span></div>
        </div>
        <div class="route">
          <div><span class="tag tag-from">DARI</span> ${froms.map(esc).join('<br>')}</div>
          <div><span class="tag tag-to">KE</span> <b>${esc(t.dest)}</b></div>
        </div>
        ${t.warnings.length ? `<div class="alert alert-warn">${t.warnings.map(esc).join('<br>')}</div>` : ''}
        <div class="table-wrap"><table>
          <thead><tr><th>Nama Motif</th><th>No. Container</th><th>Lokasi Awal</th><th class="num">Qty Roll</th><th class="num">Total Qty</th></tr></thead>
          <tbody>${t.groups.map((g) => `<tr><td><b>${esc(g.motif)}</b><div class="muted">${esc(g.code)}</div><details><summary>Lihat ${g.lots.length} lot</summary><div class="lots">${g.lots.map(esc).join(', ')}</div></details></td><td>${esc(g.container) || '-'}</td><td>${esc(g.from)}</td><td class="num">${g.rolls}</td><td class="num">${fmtNum(g.qty)} ${esc(g.uom)}</td></tr>`).join('')}</tbody>
          <tfoot><tr><td colspan="3">Total</td><td class="num">${rolls}</td><td class="num">${qtyTxt}</td></tr></tfoot>
        </table></div></div>`;
    }).join('');
  }

  // upload: pilih file / drag & drop
  $('fileInput').addEventListener('change', (e) => { handleFiles(e.target.files); e.target.value = ''; });
  const dz = $('dropzone');
  ['dragenter', 'dragover'].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove('over'); }));
  dz.addEventListener('drop', (e) => handleFiles(e.dataTransfer.files));

  $('btnCancel').addEventListener('click', () => { pending = []; renderPending(); });
  $('btnSave').addEventListener('click', () => {
    if (!pending.length) return;
    const dup = [...new Set(pending.map((t) => t.doc).filter((d) => history.some((h) => h.doc === d)))];
    if (dup.length && !confirm(`Dokumen ${dup.join(', ')} sudah pernah disimpan. Tetap simpan lagi?`)) return;
    const base = Date.now();
    pending.slice().reverse().forEach((t, i) => {
      const { warnings, file, ...rec } = t;
      history.unshift({ ...rec, id: (base + i).toString(36) + Math.random().toString(36).slice(2, 5) });
    });
    saveHistory();
    const n = pending.length; pending = []; renderPending(); renderHistory();
    toast(`${n} transaksi tersimpan.`);
  });

  // ---------- Riwayat ----------
  function renderHistory() {
    const q = $('histSearch').value.trim().toLowerCase();
    const items = history.filter((t) => !q || JSON.stringify(t).toLowerCase().includes(q));
    $('histList').innerHTML = items.length ? items.map((t) => {
      const rolls = t.groups.reduce((s, g) => s + g.rolls, 0);
      return `<div class="tx">
        <div class="tx-head">
          <div class="tx-title"><input type="checkbox" class="chk-box" data-sel="${t.id}" ${selected.has(t.id) ? 'checked' : ''} title="Pilih transaksi ini"> <b>${esc(t.doc)}</b> <span class="tx-meta">· ${fmtDate(t.date)} · ${new Date(t.ts).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })} · ${rolls} roll</span>
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
    visibleIds = items.map((t) => t.id);
    updateSelUI();
  }
  let visibleIds = [];
  function updateSelUI() {
    const n = [...selected].filter((id) => history.some((t) => t.id === id)).length;
    $('selCount').textContent = `${n} transaksi dipilih`;
    $('btnExportSel').disabled = n === 0;
    const all = $('selAll');
    all.checked = visibleIds.length > 0 && visibleIds.every((id) => selected.has(id));
    all.indeterminate = !all.checked && visibleIds.some((id) => selected.has(id));
  }
  $('selAll').addEventListener('change', (e) => {
    visibleIds.forEach((id) => (e.target.checked ? selected.add(id) : selected.delete(id)));
    renderHistory();
  });
  $('histList').addEventListener('change', (e) => {
    const c = e.target.closest('[data-sel]'); if (!c) return;
    c.checked ? selected.add(c.dataset.sel) : selected.delete(c.dataset.sel);
    updateSelUI();
  });
  $('btnExportSel').addEventListener('click', () => {
    const txs = history.filter((t) => selected.has(t.id));
    exportPDF(txs, `Transaksi terpilih`, 'perpindahan-barang-terpilih');
  });
  $('histSearch').addEventListener('input', renderHistory);
  $('histList').addEventListener('click', (e) => {
    const pdf = e.target.closest('[data-pdf]'), del = e.target.closest('[data-del]');
    if (pdf) { const t = history.find((x) => x.id === pdf.dataset.pdf); exportPDF([t], `Transaksi ${t.doc}`, `transaksi-${t.doc}`); }
    if (del && confirm('Hapus transaksi ini dari riwayat?')) { history = history.filter((x) => x.id !== del.dataset.del); selected.delete(del.dataset.del); saveHistory(); renderHistory(); }
  });

  // ---------- Export PDF ----------
  // Mode "ringkas": satu baris per motif+container. Mode "detail": satu baris per barcode/lot.
  function exportPDF(txs, subtitle, filename) {
    if (!txs.length) return toast('Tidak ada transaksi untuk diexport.', true);
    if (!window.jspdf) return toast('Library PDF gagal dimuat. Periksa koneksi internet.', true);
    const detail = $('pdfMode').value === 'detail';
    const { jsPDF } = window.jspdf;
    const orientation = $('pdfOrient').value === 'portrait' ? 'portrait' : 'landscape';
    const p = orientation === 'portrait';
    const doc = new jsPDF({ orientation, unit: 'pt', format: 'a4' });
    const W = doc.internal.pageSize.getWidth(), H = doc.internal.pageSize.getHeight();
    const M = 40;
    const AMBER = [217, 119, 6], GREEN = [17, 122, 69];

    const sorted = [...txs].sort((a, b) => a.ts - b.ts);
    const rollsOf = (t) => t.groups.reduce((s, g) => s + g.rolls, 0);
    const grandRolls = sorted.reduce((s, t) => s + rollsOf(t), 0);

    doc.setFont('helvetica', 'bold'); doc.setFontSize(15); doc.setTextColor(20);
    doc.text('Laporan Perpindahan Barang', M, 40);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(10);
    doc.text(`${subtitle}  |  ${detail ? 'Detail semua lot' : 'Ringkas'}  |  ${sorted.length} transaksi  |  ${grandRolls} roll`, M, 57);
    doc.text('Dicetak: ' + new Date().toLocaleString('id-ID'), W - M, 40, { align: 'right' });

    let y = 76;
    sorted.forEach((t, idx) => {
      const froms = [...new Set(t.groups.map((g) => g.from))];
      const badgeW = 34, lh = 12;
      const textX = M + 8 + badgeW + 8, textW = W - M - 8 - textX;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5);
      const fromLines = froms.flatMap((f) => doc.splitTextToSize(f, textW));
      doc.setFont('helvetica', 'bold');
      const destLines = doc.splitTextToSize(t.dest, textW);
      const headerH = 14 + 16 + fromLines.length * lh + 8 + destLines.length * lh + 8;

      if (y + headerH + 60 > H - 40) { doc.addPage(); y = 40; }

      // Kotak header transaksi
      doc.setFillColor(243, 246, 252); doc.setDrawColor(190, 200, 220);
      doc.rect(M, y, W - 2 * M, headerH, 'FD');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(20);
      doc.text(`Transaksi ${idx + 1}  -  Source Dokumen: ${t.doc}`, M + 8, y + 17);
      const time = new Date(t.ts).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5);
      doc.text(`Tanggal: ${fmtDate(t.date)} ${time}`, W - M - 8, y + 17, { align: 'right' });

      const badge = (label, color, by) => {
        doc.setFillColor(...color); doc.rect(M + 8, by - 9, badgeW, 13, 'F');
        doc.setTextColor(255); doc.setFont('helvetica', 'bold'); doc.setFontSize(7.5);
        doc.text(label, M + 8 + badgeW / 2, by + 0.5, { align: 'center' });
      };
      let yy = y + 36;
      badge('DARI', AMBER, yy);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); doc.setTextColor(110, 60, 5);
      fromLines.forEach((l, i) => doc.text(l, textX, yy + i * lh));
      yy += fromLines.length * lh + 8;
      badge('KE', GREEN, yy);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9.5); doc.setTextColor(10, 90, 50);
      destLines.forEach((l, i) => doc.text(l, textX, yy + i * lh));
      doc.setTextColor(20);

      // Tabel. Kolom Lokasi Awal hanya muncul jika transaksi berasal dari lebih dari satu lokasi
      const showFromCol = froms.length > 1;
      const uomTotals = {};
      let head, body, columnStyles, foot;
      const fromCol = showFromCol ? [{ label: 'Lokasi Awal', w: 200 }] : [];
      if (detail) {
        head = [['No', 'Barcode / Lot', 'No. Container', 'Nama Motif', ...fromCol.map((c) => c.label), 'Qty Roll', 'Qty']];
        body = []; let n = 0;
        t.groups.forEach((g) => g.lots.forEach((lot, i) => {
          const q = g.qtys ? g.qtys[i] : undefined;
          if (q != null) uomTotals[g.uom] = (uomTotals[g.uom] || 0) + q;
          body.push([String(++n), lot, g.container || '-', g.motif, ...(showFromCol ? [g.from] : []), '1', q != null ? `${fmtNum(q)} ${g.uom}` : '-']);
        }));
        const qtyTotal = Object.entries(uomTotals).map(([u, v]) => `${fmtNum(v)} ${u}`).join(' + ') || '-';
        foot = [['', '', '', 'Total', ...(showFromCol ? [''] : []), String(rollsOf(t)), qtyTotal]];
        const k = showFromCol ? 1 : 0;
        columnStyles = { 0: { cellWidth: p ? 25 : 30 }, 1: { cellWidth: p ? 88 : 100 }, 2: { cellWidth: p ? 70 : 85 }, [4 + k]: { cellWidth: p ? 38 : 50, halign: 'right' }, [5 + k]: { cellWidth: p ? 62 : 85, halign: 'right' } };
        if (showFromCol) columnStyles[4] = { cellWidth: p ? 130 : 170 };
      } else {
        head = [['No', 'No. Container', 'Nama Motif', ...fromCol.map((c) => c.label), 'Qty Roll']];
        body = t.groups.map((g, i) => [String(i + 1), g.container || '-', g.motif, ...(showFromCol ? [g.from] : []), String(g.rolls)]);
        foot = [['', '', 'Total Roll', ...(showFromCol ? [''] : []), String(rollsOf(t))]];
        columnStyles = { 0: { cellWidth: p ? 25 : 30 }, 1: { cellWidth: p ? 90 : 110 }, [3 + (showFromCol ? 1 : 0)]: { cellWidth: p ? 45 : 60, halign: 'right' } };
        if (showFromCol) columnStyles[3] = { cellWidth: p ? 160 : 230 };
      }
      doc.autoTable({
        startY: y + headerH,
        head, body, foot,
        styles: { fontSize: detail ? 8 : 8.5, cellPadding: detail ? 3 : 4, overflow: 'linebreak' },
        headStyles: { fillColor: [31, 94, 255] },
        footStyles: { fillColor: [245, 247, 251], textColor: 20, fontStyle: 'bold' },
        columnStyles,
        margin: { left: M, right: M, top: 40, bottom: 40 },
        showFoot: 'lastPage',
        rowPageBreak: 'avoid'
      });
      y = doc.lastAutoTable.finalY + 22;
    });

    const pages = doc.internal.getNumberOfPages();
    for (let p = 1; p <= pages; p++) {
      doc.setPage(p); doc.setFontSize(8); doc.setTextColor(110); doc.setFont('helvetica', 'normal');
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
  $('expDate').value = todayStr();
  renderHistory();
})();
