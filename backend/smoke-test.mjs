/*
 * Smoke test for report-core.js — validates the Xero-JSON -> render-model
 * transform and that the result writes to a real .xlsx. Run: npm run smoke
 */
import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import os from 'node:os';
import fs from 'node:fs';

const require = createRequire(import.meta.url);
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ReportCore = require(path.join(__dirname, '..', 'public', 'report-core.js'));
const XLSX = require('xlsx-js-style');

let failures = 0;
const check = (name, cond) => {
  if (cond) { console.log(`  ✓ ${name}`); }
  else { console.error(`  ✗ ${name}`); failures++; }
};

console.log('\nreport-core smoke test\n----------------------');

// 1) Flat report (Profit & Loss) ------------------------------------------------
console.log('Profit and Loss (flat, 2 periods):');
const pl = ReportCore.parseReport(ReportCore.DEMO.ProfitAndLoss);
check('titles captured', pl.titles[0] === 'Profit and Loss');
check('two period labels', pl.valueCount === 2 && pl.periodLabels.length === 2);
check('flat => maxDepth 0', pl.maxDepth === 0);
check('section title present', pl.lines.some((l) => l.type === 'sectionTitle' && l.cells[0] === 'Trading Income'));
check('net profit summary present', pl.lines.some((l) => l.type === 'summary' && l.cells[0] === 'Net Profit'));

const plAoa = ReportCore.buildAOA(pl).aoa;
check('row1 = report name', plAoa[0][0] === 'Profit and Loss');
check('header has "Account"', plAoa[4][0] === 'Account');
check('header has period label', plAoa[4][1] === '31 May 2026');
const salesRow = plAoa.find((r) => r[0] === 'Sales');
check('Sales value is a number', salesRow && typeof salesRow[1] === 'number' && salesRow[1] === 12500);

// 2) Nested report (Balance Sheet) ---------------------------------------------
console.log('Balance Sheet (nested):');
const bs = ReportCore.parseReport(ReportCore.DEMO.BalanceSheet);
check('nested => maxDepth 1', bs.maxDepth === 1);
const { aoa: bsAoa } = ReportCore.buildAOA(bs);
const assetsRow = bsAoa.find((r) => r[0] === 'Assets');
check('top-level "Assets" in column A', !!assetsRow);
const currentAssets = bsAoa.find((r) => r[1] === 'Current Assets');
check('sub-section "Current Assets" indented to column B', !!currentAssets);
const cash = bsAoa.find((r) => r[1] === 'Business Bank Account');
check('leaf item indented to column B', cash && typeof cash[2] === 'number');
check('BS "Account" header in column B', bsAoa[4][1] === 'Account');

// 3) Depth rule on a hand-built nested input -----------------------------------
console.log('Depth rule (synthetic):');
const synthetic = {
  Reports: [{
    ReportName: 'Synthetic', ReportType: 'Synthetic',
    ReportTitles: ['Synthetic', 'Org', 'Today'],
    Rows: [
      { RowType: 'Header', Cells: [{ Value: '' }, { Value: 'This Year' }] },
      { RowType: 'Section', Title: 'Group', Rows: [
        { RowType: 'Section', Title: 'Sub', Rows: [
          { RowType: 'Row', Cells: [{ Value: 'Item' }, { Value: '10.00' }] },
          { RowType: 'SummaryRow', Cells: [{ Value: 'Total Sub' }, { Value: '10.00' }] },
        ] },
        { RowType: 'SummaryRow', Cells: [{ Value: 'Total Group' }, { Value: '10.00' }] },
      ] },
    ],
  }],
};
const syn = ReportCore.parseReport(synthetic);
const getDepth = (label) => {
  const l = syn.lines.find((x) => x.cells[0] === label);
  return l ? l.depth : -1;
};
check('Group at depth 0', getDepth('Group') === 0);
check('Sub at depth 1', getDepth('Sub') === 1);
check('Item at depth 1 (leaf of Sub)', getDepth('Item') === 1);
check('Total Group at depth 0 (leaf of Group)', getDepth('Total Group') === 0);

// 4) Number / amount formatting ------------------------------------------------
console.log('Formatting:');
check('formatAmount thousands', ReportCore.formatAmount('3401917.00') === '3,401,917.00');
check('formatAmount negative -> parens', ReportCore.formatAmount('-78159.00') === '(78,159.00)');
check('formatAmount zero -> dash', ReportCore.formatAmount('0') === '-');
check('formatAmount blank -> dash', ReportCore.formatAmount('') === '-');
check('formatAmount percent passthrough', ReportCore.formatAmount('11.1%') === '11.1%');

// 5) Write a real multi-sheet workbook -----------------------------------------
console.log('Excel write (all demo reports):');
const wb = XLSX.utils.book_new();
const used = new Set();
for (const type of Object.keys(ReportCore.DEMO)) {
  const n = ReportCore.parseReport(ReportCore.DEMO[type]);
  const { aoa, numberCells } = ReportCore.buildAOA(n);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  for (const { r, c } of numberCells) {
    const ref = XLSX.utils.encode_cell({ r, c });
    if (ws[ref]) ws[ref].z = ReportCore.EXCEL_NUM_FMT;
  }
  XLSX.utils.book_append_sheet(wb, ws, ReportCore.sanitizeSheetName(n.reportName, used));
}
const outPath = path.join(os.tmpdir(), 'xero-smoke-test.xlsx');
XLSX.writeFile(wb, outPath);
const size = fs.statSync(outPath).size;
check('workbook written', size > 2000);
check('one sheet per demo report', wb.SheetNames.length === Object.keys(ReportCore.DEMO).length);
console.log(`  (wrote ${size} bytes to ${outPath})`);

console.log('----------------------');
if (failures) { console.error(`${failures} check(s) FAILED\n`); process.exit(1); }
console.log('All checks passed ✅\n');
