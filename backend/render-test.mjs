/*
 * render-test.mjs — generate sample PDF + Excel from the demo reports using the
 * SAME renderer the browser ships (report-render.js), so the export layout can
 * be eyeballed/verified without a browser. Output -> ../tmp-export/ (git-ignored).
 * Run: npm run render-test
 */
import { createRequire } from 'node:module';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const pub = path.join(__dirname, '..', 'public');

const ReportCore = require(path.join(pub, 'report-core.js'));
const ReportRender = require(path.join(pub, 'report-render.js'));
const { jsPDF } = require('jspdf');
require('jspdf-autotable'); // patches jsPDF.API.autoTable
const XLSX = require('xlsx');
const JSZip = require('jszip');

const R = ReportRender.createRenderer({ jsPDF, XLSX, JSZip, ReportCore });

const outDir = path.join(__dirname, '..', 'tmp-export');
fs.mkdirSync(outDir, { recursive: true });

const results = Object.keys(ReportCore.DEMO).map((type) => {
  const normalized = ReportCore.parseReport(ReportCore.DEMO[type]);
  return { report: { name: normalized.reportName }, normalized };
});

const doc = R.combinedPdf(results, 'Demo Company (AU)');
const pdfPath = path.join(outDir, 'sample-combined.pdf');
fs.writeFileSync(pdfPath, Buffer.from(doc.output('arraybuffer')));

const wb = R.workbookFor(results);
const xlsxPath = path.join(outDir, 'sample-combined.xlsx');
fs.writeFileSync(xlsxPath, Buffer.from(R.xlsxWrite(wb)));

console.log('Wrote', pdfPath);
console.log('Wrote', xlsxPath);
console.log('PDF pages:', doc.getNumberOfPages(), '· Workbook sheets:', wb.SheetNames.join(', '));
