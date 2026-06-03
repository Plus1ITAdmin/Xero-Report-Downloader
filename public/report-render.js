/*
 * report-render.js
 * -------------------------------------------------------------------------
 * Turns the normalised models from report-core.js into PDF (jsPDF + autotable)
 * and Excel (SheetJS) that mirror Xero's own exports.
 *
 * Handles two model kinds:
 *   - financial reports (P&L, Balance Sheet, ...) — portrait, section/row layout
 *   - { kind:'generalledger' } — landscape, grouped-by-account 10-column layout
 *
 * Dependency-injected so the same rendering runs in the browser (window globals)
 * and in Node (require()'d modules), which is how the output is verified.
 *
 *   const R = ReportRender.createRenderer({ jsPDF, XLSX, JSZip, ReportCore });
 * -------------------------------------------------------------------------
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ReportRender = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function createRenderer(deps) {
    const { jsPDF, XLSX, JSZip, ReportCore } = deps;

    const fmtShortDate = (d) => (d instanceof Date && !isNaN(d))
      ? d.toLocaleDateString('en-AU', { day: 'numeric', month: 'short', year: 'numeric' }) : '';
    // GL number format: thousands + 2dp, negatives in parens, blank for 0/empty.
    const glNum = (v) => {
      if (v == null || v === '') return '';
      const num = Number(v);
      if (!isFinite(num) || num === 0) return '';
      const a = Math.abs(num).toLocaleString('en-AU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      return num < 0 ? `(${a})` : a;
    };
    const orientationFor = (n) => (n && n.kind === 'generalledger' ? 'landscape' : 'portrait');
    const newDoc = (orientation) => new jsPDF({ unit: 'pt', format: 'a4', orientation: orientation || 'portrait', compress: true });
    const newPdf = () => newDoc('portrait');

    // ---- Financial report body (draws on the current page) -----------------
    function drawReportBody(doc, n) {
      const M = 42;
      let y = 58;
      doc.setTextColor(33, 41, 51);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(22); doc.text(String(n.titles[0] || 'Report'), M, y); y += 22;
      doc.setFontSize(12);
      for (let i = 1; i < n.titles.length; i++) { doc.text(String(n.titles[i]), M, y); y += 16; }
      y += 6;

      const head = n.periodLabels.length ? [['', ...n.periodLabels]] : undefined;
      const body = [], meta = [];
      for (const line of n.lines) {
        if (line.type === 'spacer') { body.push([' ']); meta.push({ type: 'spacer' }); continue; }
        if (line.type === 'sectionTitle') {
          const r = new Array(n.valueCount + 1).fill(''); r[0] = line.cells[0] || '';
          body.push(r); meta.push({ type: 'sectionTitle', depth: line.depth }); continue;
        }
        const r = [line.cells[0] || ''];
        for (let i = 0; i < n.valueCount; i++) r.push(ReportCore.formatAmount(line.cells[i + 1]));
        body.push(r); meta.push({ type: line.type, depth: line.depth });
      }

      const valueCols = {};
      for (let i = 1; i <= n.valueCount; i++) valueCols[i] = { halign: 'right', cellWidth: 78 };

      doc.autoTable({
        head, body, startY: y, margin: { left: M, right: M }, theme: 'plain',
        styles: { font: 'helvetica', fontSize: 8.5, textColor: [40, 48, 58], cellPadding: { top: 2.6, right: 8, bottom: 2.6, left: 8 }, lineColor: [225, 230, 236], lineWidth: 0 },
        headStyles: { fontStyle: 'bold', fontSize: 8, halign: 'right', textColor: [70, 80, 92], lineWidth: { bottom: 0.7 }, lineColor: [120, 132, 145], cellPadding: { top: 3, right: 8, bottom: 5, left: 8 } },
        columnStyles: Object.assign({ 0: { halign: 'left', cellWidth: 'auto' } }, valueCols),
        didParseCell: (data) => {
          if (data.section !== 'body') return;
          const m = meta[data.row.index];
          if (!m) return;
          if (m.type === 'spacer') { data.cell.styles.minCellHeight = 4; data.cell.styles.fontSize = 4; return; }
          if (data.column.index === 0) {
            const indent = 8 + (m.depth || 0) * 13 + (m.type === 'row' ? 13 : 0);
            data.cell.styles.cellPadding = { top: 2.6, right: 8, bottom: 2.6, left: indent };
          }
          if (m.type === 'sectionTitle') { data.cell.styles.fontStyle = 'bold'; data.cell.styles.lineWidth = { bottom: 0.5 }; data.cell.styles.lineColor = [210, 216, 223]; }
          if (m.type === 'summary') { data.cell.styles.fontStyle = 'bold'; data.cell.styles.lineWidth = { top: 0.5 }; data.cell.styles.lineColor = [150, 160, 170]; }
        },
      });
    }

    // ---- General Ledger body (landscape, grouped by account) ---------------
    function drawGLBody(doc, n) {
      const M = 42;
      let y = 50;
      doc.setTextColor(33, 41, 51);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(18); doc.text(String(n.titles[0] || 'General Ledger Detail'), M, y); y += 18;
      doc.setFontSize(11);
      for (let i = 1; i < n.titles.length; i++) { doc.text(String(n.titles[i]), M, y); y += 14; }
      y += 6;

      const body = [], meta = [];
      for (const acc of n.accounts) {
        body.push([acc.name, '', '', '', '', '', '', '', '', '']); meta.push('account');
        for (const ln of acc.lines) {
          body.push([
            fmtShortDate(ln.date), ln.source, ln.description, ln.reference,
            glNum(ln.debit), glNum(ln.credit), glNum(ln.runningBalance),
            glNum(ln.gst), ln.gstRate ? `${ln.gstRate}%` : '', ln.gstRateName,
          ]);
          meta.push('row');
        }
        body.push([`Total ${acc.name}`, '', '', '', glNum(acc.totalDebit), glNum(acc.totalCredit), '', '', '', '']); meta.push('total');
        const nd = acc.net >= 0 ? acc.net : null, nc = acc.net < 0 ? -acc.net : null;
        body.push(['Net movement', '', '', '', glNum(nd), glNum(nc), '', '', '', '']); meta.push('net');
        body.push(['', '', '', '', '', '', '', '', '', '']); meta.push('spacer');
      }

      doc.autoTable({
        head: [n.columns], body, startY: y, margin: { left: M, right: M }, theme: 'plain',
        styles: { font: 'helvetica', fontSize: 7.5, textColor: [40, 48, 58], cellPadding: { top: 1.8, right: 5, bottom: 1.8, left: 5 }, overflow: 'linebreak' },
        headStyles: { fontStyle: 'bold', fontSize: 7.5, halign: 'left', fillColor: [245, 247, 250], textColor: [70, 80, 92], lineWidth: { bottom: 0.6 }, lineColor: [120, 132, 145] },
        columnStyles: {
          0: { cellWidth: 52 }, 1: { cellWidth: 62 }, 2: { cellWidth: 'auto' }, 3: { cellWidth: 78 },
          4: { halign: 'right', cellWidth: 56 }, 5: { halign: 'right', cellWidth: 56 }, 6: { halign: 'right', cellWidth: 64 },
          7: { halign: 'right', cellWidth: 44 }, 8: { halign: 'right', cellWidth: 46 }, 9: { cellWidth: 72 },
        },
        didParseCell: (data) => {
          if (data.section === 'head') { if ([4, 5, 6, 7, 8].includes(data.column.index)) data.cell.styles.halign = 'right'; return; }
          const m = meta[data.row.index];
          if (!m) return;
          if (m === 'spacer') { data.cell.styles.minCellHeight = 3; data.cell.styles.fontSize = 3; return; }
          if (m === 'account') {
            data.cell.styles.fontStyle = 'bold';
            if (data.column.index === 0) data.cell.styles.fontSize = 8.5;
            data.cell.styles.lineWidth = { bottom: 0.4 }; data.cell.styles.lineColor = [210, 216, 223];
          }
          if (m === 'total' || m === 'net') data.cell.styles.fontStyle = 'bold';
          if (m === 'total' && data.column.index === 0) { data.cell.styles.lineWidth = { top: 0.4 }; data.cell.styles.lineColor = [150, 160, 170]; }
        },
      });
    }

    const drawBody = (doc, n) => (n && n.kind === 'generalledger' ? drawGLBody(doc, n) : drawReportBody(doc, n));

    function stampPdfFooters(doc, org) {
      const pages = doc.getNumberOfPages();
      const today = new Date().toLocaleDateString('en-AU', { day: 'numeric', month: 'short', year: 'numeric' });
      for (let i = 1; i <= pages; i++) {
        doc.setPage(i);
        const w = doc.internal.pageSize.getWidth(), h = doc.internal.pageSize.getHeight();
        doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(150, 160, 170);
        doc.text(`${org}   ·   ${today}`, 42, h - 24);
        doc.text(`Page ${i} of ${pages}`, w - 42, h - 24, { align: 'right' });
      }
    }

    function combinedPdf(results, org) {
      const doc = newDoc(orientationFor(results[0].normalized));
      results.forEach((x, i) => {
        if (i > 0) doc.addPage('a4', orientationFor(x.normalized));
        drawBody(doc, x.normalized);
      });
      stampPdfFooters(doc, org);
      return doc;
    }
    function singlePdf(normalized, org) {
      const doc = newDoc(orientationFor(normalized));
      drawBody(doc, normalized);
      stampPdfFooters(doc, org);
      return doc;
    }

    // ---- Excel -------------------------------------------------------------
    function applyColWidths(ws, aoa) {
      let maxCols = 0;
      aoa.forEach((row) => { maxCols = Math.max(maxCols, row.length); });
      const widths = [];
      for (let c = 0; c < maxCols; c++) {
        let w = 10;
        aoa.forEach((row) => { const v = row[c]; if (v != null && v !== '') w = Math.max(w, String(v).length + 2); });
        widths.push({ wch: Math.min(w, 50) });
      }
      ws['!cols'] = widths;
    }

    function reportWorksheet(n) {
      const { aoa, numberCells } = ReportCore.buildAOA(n);
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      for (const { r, c } of numberCells) {
        const ref = XLSX.utils.encode_cell({ r, c });
        if (ws[ref]) ws[ref].z = ReportCore.EXCEL_NUM_FMT;
      }
      applyColWidths(ws, aoa);
      return ws;
    }

    function glWorksheet(n) {
      const { aoa, dateCells, rowMeta } = ReportCore.glToAOA(n);
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const range = XLSX.utils.decode_range(ws['!ref']);
      const font = (sz, bold) => ({ name: 'Arial', sz, bold: !!bold });
      const blackBottom = { bottom: { style: 'thin', color: { rgb: '000000' } } };

      // Base font Arial 12; client name 14 bold; title/header/category/total/net bold.
      for (let r = range.s.r; r <= range.e.r; r++) {
        const type = rowMeta[r];
        for (let c = range.s.c; c <= range.e.c; c++) {
          const ref = XLSX.utils.encode_cell({ r, c });
          let cell = ws[ref];
          // Create empty cells on category rows so the underline spans all columns.
          if (!cell) { if (type === 'account') cell = ws[ref] = { t: 's', v: '' }; else continue; }
          const s = cell.s || (cell.s = {});
          if (type === 'client') s.font = font(14, true);
          else if (type === 'title' || type === 'header' || type === 'account' || type === 'total' || type === 'net') s.font = font(12, true);
          else s.font = font(12, false);
          if (type === 'account') s.border = blackBottom; // black line under category name
        }
      }
      // Dates -> "d mmm yyyy"
      for (const { r, c } of dateCells) {
        const ref = XLSX.utils.encode_cell({ r, c });
        if (ws[ref]) { ws[ref].t = 'n'; (ws[ref].s || (ws[ref].s = {})).numFmt = 'd mmm yyyy'; }
      }
      // GST Rate (col 8) on data rows -> display as a percentage (value stays the rate).
      for (let r = range.s.r; r <= range.e.r; r++) {
        if (rowMeta[r] !== 'row') continue;
        const ref = XLSX.utils.encode_cell({ r, c: 8 });
        if (ws[ref] && typeof ws[ref].v === 'number') (ws[ref].s || (ws[ref].s = {})).numFmt = '0.##"%"';
      }
      // GL has 10 fixed columns — give them sensible fixed widths.
      ws['!cols'] = [12, 14, 40, 20, 13, 13, 15, 11, 10, 18].map((wch) => ({ wch }));
      return ws;
    }

    const worksheetFor = (n) => (n && n.kind === 'generalledger' ? glWorksheet(n) : reportWorksheet(n));

    function workbookFor(results) {
      const wb = XLSX.utils.book_new();
      const used = new Set();
      for (const { normalized } of results) {
        XLSX.utils.book_append_sheet(wb, worksheetFor(normalized), ReportCore.sanitizeSheetName(normalized.reportName, used));
      }
      return wb;
    }
    const xlsxWrite = (wb) => XLSX.write(wb, { type: 'array', bookType: 'xlsx' });

    return { JSZip, newPdf, stampPdfFooters, combinedPdf, singlePdf, worksheetFor, workbookFor, xlsxWrite };
  }

  return { createRenderer };
}));
