/*
 * report-render.js
 * -------------------------------------------------------------------------
 * Turns the normalised report model (from report-core.js) into PDF (jsPDF +
 * autotable) and Excel (SheetJS) that mirror Xero's own export layout.
 *
 * Dependency-injected so the exact same rendering runs in the browser (passing
 * window globals) and in Node (passing require()'d modules) — which is how the
 * output is verified headlessly.
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

    // ---- PDF ---------------------------------------------------------------
    function newPdf() {
      return new jsPDF({ unit: 'pt', format: 'a4', compress: true });
    }

    /** Render one report onto the doc, starting a new page unless it's first. */
    function addReportToPdf(doc, n, isFirst) {
      const M = 42; // page margin (pt)
      if (!isFirst) doc.addPage();

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
          if (m.type === 'sectionTitle') {
            data.cell.styles.fontStyle = 'bold';
            data.cell.styles.lineWidth = { bottom: 0.5 };
            data.cell.styles.lineColor = [210, 216, 223];
          }
          if (m.type === 'summary') {
            data.cell.styles.fontStyle = 'bold';
            data.cell.styles.lineWidth = { top: 0.5 };
            data.cell.styles.lineColor = [150, 160, 170];
          }
        },
      });
    }

    function stampPdfFooters(doc, org) {
      const pages = doc.getNumberOfPages();
      const w = doc.internal.pageSize.getWidth(), h = doc.internal.pageSize.getHeight();
      const today = new Date().toLocaleDateString('en-AU', { day: 'numeric', month: 'short', year: 'numeric' });
      for (let i = 1; i <= pages; i++) {
        doc.setPage(i);
        doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(150, 160, 170);
        doc.text(`${org}   ·   ${today}`, 42, h - 24);
        doc.text(`Page ${i} of ${pages}`, w - 42, h - 24, { align: 'right' });
      }
    }

    /** Build a combined PDF (one report per page) and return the jsPDF doc. */
    function combinedPdf(results, org) {
      const doc = newPdf();
      results.forEach((x, i) => addReportToPdf(doc, x.normalized, i === 0));
      stampPdfFooters(doc, org);
      return doc;
    }
    /** Build a single-report PDF and return the jsPDF doc. */
    function singlePdf(normalized, org) {
      const doc = newPdf();
      addReportToPdf(doc, normalized, true);
      stampPdfFooters(doc, org);
      return doc;
    }

    // ---- Excel -------------------------------------------------------------
    function worksheetFor(n) {
      const { aoa, numberCells } = ReportCore.buildAOA(n);
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      for (const { r, c } of numberCells) {
        const ref = XLSX.utils.encode_cell({ r, c });
        if (ws[ref]) ws[ref].z = ReportCore.EXCEL_NUM_FMT;
      }
      let maxCols = 0;
      aoa.forEach((row) => { maxCols = Math.max(maxCols, row.length); });
      const widths = [];
      for (let c = 0; c < maxCols; c++) {
        let w = 10;
        aoa.forEach((row) => { const v = row[c]; if (v != null && v !== '') w = Math.max(w, String(v).length + 2); });
        widths.push({ wch: Math.min(w, 48) });
      }
      ws['!cols'] = widths;
      return ws;
    }

    function workbookFor(results) {
      const wb = XLSX.utils.book_new();
      const used = new Set();
      for (const { normalized } of results) {
        XLSX.utils.book_append_sheet(wb, worksheetFor(normalized), ReportCore.sanitizeSheetName(normalized.reportName, used));
      }
      return wb;
    }

    /** Return the .xlsx bytes (Uint8Array); caller wraps in Blob or Buffer. */
    function xlsxWrite(wb) {
      return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    }

    return {
      JSZip,
      newPdf, addReportToPdf, stampPdfFooters, combinedPdf, singlePdf,
      worksheetFor, workbookFor, xlsxWrite,
    };
  }

  return { createRenderer };
}));
