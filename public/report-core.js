/*
 * report-core.js
 * -------------------------------------------------------------------------
 * Environment-agnostic core for the Xero Report Downloader.
 *
 * Xero's Reports API does NOT return PDFs/spreadsheets — it returns a generic
 * "rows & cells" JSON structure shared across every report type. This module
 * turns that into a normalised, render-friendly shape that the browser paints
 * into PDF (jsPDF) or Excel (SheetJS), reproducing Xero's own export layout.
 *
 * Indentation model (reverse-engineered from real Xero exports):
 *   - A report is a tree of Sections; leaf Rows/SummaryRows sit at their
 *     section's depth, and only NESTED sub-sections increase depth.
 *     => Profit & Loss is flat (depth 0 everywhere);
 *        Balance Sheet nests (Assets[0] > Current Assets[1] > items[1]).
 *   - In Excel, the label is written into the column matching its depth.
 *   - In PDF, the label is indented by its depth (with line items nudged in
 *     a little further under their section header, like Xero).
 *
 * Loads in the browser via <script src> (sets window.ReportCore) and in Node
 * via require()/import (module.exports) so the same logic is unit-tested.
 * -------------------------------------------------------------------------
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.ReportCore = factory();
  }
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // Excel number format: 1,234.00 / (1,234.00) for negatives / "-" for zero.
  const EXCEL_NUM_FMT = '#,##0.00;(#,##0.00);"-"';

  /**
   * Convert a Xero cell value to a real number where it clearly is one,
   * otherwise return the original string. Lets Excel treat "1,234.50" /
   * "(1,234.50)" / "-78159.0000" as numbers but keeps "11.1%" / labels as text.
   */
  function toNumberOrString(v) {
    if (v === null || v === undefined || v === '') return '';
    const s = String(v).replace(/,/g, '').trim();
    if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
    if (/^\(\d+(\.\d+)?\)$/.test(s)) return -Number(s.slice(1, -1)); // (123.00) => -123
    return v;
  }

  /**
   * Format a raw cell value for PDF display the way Xero does:
   *   1234.5  -> "1,234.50"   |   -78159 -> "(78,159.00)"
   *   0 / ""  -> "-"          |   "11.1%" / text -> unchanged
   */
  function formatAmount(raw) {
    if (raw === null || raw === undefined) return '';
    const s = String(raw).trim();
    if (s === '') return '-';
    const cleaned = s.replace(/[\s,$]/g, '');
    const negative = /^\(.*\)$/.test(cleaned) || /^-/.test(cleaned);
    const num = cleaned.replace(/[()]/g, '').replace(/^-/, '');
    if (!/^\d+(\.\d+)?$/.test(num)) return s; // not a plain number (e.g. percentage)
    const n = Number(num);
    if (n === 0) return '-';
    const formatted = n.toLocaleString('en-AU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    return negative ? `(${formatted})` : formatted;
  }

  /**
   * Normalise a Xero report response into a flat list of typed, depth-tagged
   * lines. Accepts the full API response ({ Reports:[{...}] }) or one report.
   *
   * Returns:
   *   {
   *     reportType, reportName,
   *     titles:       string[],   // ReportTitles (name, org, date range)
   *     periodLabels: string[],   // column headers (one per value column)
   *     valueCount:   number,     // number of value columns
   *     maxDepth:     number,     // deepest label indent
   *     lines: [ { type, cells:string[], depth } ]   // sectionTitle|row|summary|spacer
   *   }
   */
  function parseReport(input) {
    if (!input) throw new Error('parseReport: empty input');
    const report = input.Reports ? input.Reports[0] : input;
    if (!report) throw new Error('parseReport: no report found in response');

    const titles = (report.ReportTitles && report.ReportTitles.length)
      ? report.ReportTitles.slice()
      : [report.ReportName || 'Report'];

    const lines = [];
    let headerCells = null;
    let maxDepth = 0;

    const cellsOf = (row) => (row.Cells || []).map((c) => (c && c.Value != null ? String(c.Value) : ''));

    function emitLeaf(row, depth) {
      const cells = cellsOf(row);
      maxDepth = Math.max(maxDepth, depth);
      if (row.RowType === 'SummaryRow') lines.push({ type: 'summary', cells, depth });
      else lines.push({ type: 'row', cells, depth });
    }

    function walkSection(section, depth) {
      if (section.Title) {
        maxDepth = Math.max(maxDepth, depth);
        lines.push({ type: 'sectionTitle', cells: [section.Title], depth });
      }
      for (const child of (section.Rows || [])) {
        if (child.RowType === 'Section') walkSection(child, depth + 1);
        else if (child.RowType === 'Header') headerCells = cellsOf(child);
        else emitLeaf(child, depth);
      }
    }

    for (const row of (report.Rows || [])) {
      if (row.RowType === 'Section') {
        walkSection(row, 0);
        lines.push({ type: 'spacer', cells: [], depth: 0 });
      } else if (row.RowType === 'Header') {
        headerCells = cellsOf(row);
      } else {
        emitLeaf(row, 0);
      }
    }

    const periodLabels = headerCells ? headerCells.slice(1) : [];
    let valueCount = periodLabels.length;
    for (const l of lines) {
      if (l.type === 'row' || l.type === 'summary') valueCount = Math.max(valueCount, l.cells.length - 1);
    }
    if (valueCount < 1) valueCount = 1;

    return {
      reportType: report.ReportType || '',
      reportName: report.ReportName || (titles[0] || 'Report'),
      titles,
      periodLabels,
      valueCount,
      maxDepth,
      lines,
    };
  }

  /**
   * Build an array-of-arrays for SheetJS aoa_to_sheet, reproducing Xero's Excel
   * layout: title block, blank row, "Account" + period header, then rows with
   * the label placed in its depth column and numeric values right-most.
   * Returns { aoa, numberCells } where numberCells lists {r,c} to number-format.
   */
  function buildAOA(normalized) {
    const { titles, periodLabels, valueCount, maxDepth, lines } = normalized;
    const firstValueCol = maxDepth + 1;
    const aoa = [];
    const numberCells = [];

    for (const t of titles) aoa.push([t]);
    aoa.push([]); // blank under title block

    // Header row: "Account" in the last label column, period labels to the right.
    const header = [];
    header[maxDepth] = 'Account';
    for (let i = 0; i < valueCount; i++) header[firstValueCol + i] = periodLabels[i] || '';
    for (let i = 0; i < header.length; i++) if (header[i] === undefined) header[i] = '';
    aoa.push(header);

    for (const line of lines) {
      if (line.type === 'spacer') { aoa.push([]); continue; }
      const row = [];
      const label = line.cells[0] != null ? line.cells[0] : '';
      row[line.depth] = label;
      if (line.type === 'row' || line.type === 'summary') {
        for (let i = 0; i < valueCount; i++) {
          const val = toNumberOrString(line.cells[i + 1]);
          const col = firstValueCol + i;
          row[col] = val;
          if (typeof val === 'number') numberCells.push({ r: aoa.length, c: col });
        }
      }
      for (let i = 0; i < row.length; i++) if (row[i] === undefined) row[i] = '';
      aoa.push(row);
    }

    return { aoa, numberCells };
  }

  /** Excel sheet names: max 31 chars, no []:*?/\ characters, unique, non-empty. */
  function sanitizeSheetName(name, used) {
    let n = String(name || 'Report').replace(/[\[\]\:\*\?\/\\]/g, ' ').trim().slice(0, 31) || 'Report';
    if (used) {
      const base = n.slice(0, 28);
      let i = 2;
      while (used.has(n)) n = `${base} ${i++}`;
      used.add(n);
    }
    return n;
  }

  // ---------------------------------------------------------------------------
  // General Ledger Detail — built from the Journals API (not a /Reports endpoint)
  // ---------------------------------------------------------------------------
  const SOURCE_TYPES = {
    ACCREC: 'Sales Invoice', ACCPAY: 'Bill',
    ACCRECCREDIT: 'Sales Credit Note', ACCPAYCREDIT: 'Bill Credit Note',
    ACCRECPAYMENT: 'Sales Payment', ACCPAYPAYMENT: 'Bill Payment',
    ARCREDITPAYMENT: 'Sales Credit Note Refund', APCREDITPAYMENT: 'Bill Credit Note Refund',
    CASHREC: 'Receive Money', CASHPAID: 'Spend Money',
    TRANSFER: 'Bank Transfer', TRANSFERPAYMENT: 'Bank Transfer',
    ARPREPAYMENT: 'Sales Prepayment', APPREPAYMENT: 'Bill Prepayment',
    AROVERPAYMENT: 'Sales Overpayment', APOVERPAYMENT: 'Bill Overpayment',
    EXPCLAIM: 'Expense Claim', EXPPAYMENT: 'Expense Claim Payment',
    MANJOURNAL: 'Manual Journal', PAYSLIP: 'Payslip', WAGEPAYABLE: 'Wage Payable',
    INTEGRATEDPAYROLLPE: 'Payroll', INTEGRATEDPAYROLLPT: 'Payroll Payment',
    INTEGRATEDPAYROLLPTPAYMENT: 'Payroll Payment', EXTERNALSPENDMONEY: 'Spend Money',
  };
  function mapSource(t) {
    if (!t) return '';
    if (SOURCE_TYPES[t]) return SOURCE_TYPES[t];
    return String(t).toLowerCase().replace(/(^|\s)\S/g, (c) => c.toUpperCase());
  }

  function parseJournalDate(s) {
    if (!s) return null;
    const m = /\/Date\((-?\d+)/.exec(String(s));
    return m ? new Date(parseInt(m[1], 10)) : new Date(s);
  }
  function isoToLocalDate(s) {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(s || ''));
    return m ? new Date(+m[1], +m[2] - 1, +m[3]) : (s ? new Date(s) : null);
  }
  function excelDateSerial(d) {
    if (!(d instanceof Date) || isNaN(d)) return '';
    const epoch = Date.UTC(1899, 11, 30);
    const day = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
    return Math.round((day - epoch) / 86400000);
  }
  function fmtLongDate(d) {
    return d instanceof Date && !isNaN(d)
      ? d.toLocaleDateString('en-AU', { day: 'numeric', month: 'long', year: 'numeric' })
      : '';
  }
  const round2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;

  const GL_COLUMNS = ['Date', 'Source', 'Description', 'Reference', 'Debit', 'Credit', 'Running Balance', 'GST', 'GST Rate', 'GST Rate Name'];

  /**
   * Build a General Ledger Detail model from sub-ledger data (the Journals API
   * needs a premium plan). `data` = { Accounts, TaxRates, BankTransactions,
   * Invoices, ManualJournals }. opts: { orgName, fromDate, toDate }.
   *
   * Every line item posts directly to its income/expense account; balance-sheet
   * control accounts (bank, AR/AP, GST) are system-generated and not included.
   */
  function buildGeneralLedger(data, opts) {
    const o = opts || {};
    data = data || {};
    const accountName = {};
    let ar = null, ap = null; // Accounts Receivable / Payable system accounts
    for (const a of (data.Accounts || [])) {
      accountName[a.Code] = a.Name;
      if (a.SystemAccount === 'DEBTORS' || a.Type === 'RECEIVABLE') ar = { code: a.Code, name: a.Name };
      if (a.SystemAccount === 'CREDITORS' || a.Type === 'PAYABLE') ap = { code: a.Code, name: a.Name };
    }
    const taxMap = {};
    for (const t of (data.TaxRates || [])) {
      taxMap[t.TaxType] = { name: t.Name, rate: t.EffectiveRate != null ? Number(t.EffectiveRate) : (t.DisplayTaxRate != null ? Number(t.DisplayTaxRate) : 0) };
    }
    const accounts = new Map();
    let seq = 0;

    function add(code, fallbackName, e) {
      const key = code || fallbackName || '(No account)';
      if (!accounts.has(key)) accounts.set(key, { name: accountName[code] || fallbackName || code || '(No account)', code: code || '', lines: [] });
      e.seq = seq++;
      accounts.get(key).lines.push(e);
    }
    function netTax(line, lineAmountTypes) {
      const amt = Number(line.LineAmount) || 0, tax = Number(line.TaxAmount) || 0;
      return { net: lineAmountTypes === 'Inclusive' ? round2(amt - tax) : amt, tax };
    }
    function postLines(lines, lineAmountTypes, meta) {
      for (const line of (lines || [])) {
        const code = line.AccountCode;
        if (!code) continue;
        const { net, tax } = netTax(line, lineAmountTypes);
        const value = round2(net * meta.sign);
        const tm = taxMap[line.TaxType] || {};
        const desc = meta.contact ? meta.contact + (line.Description ? ' - ' + line.Description : '') : (line.Description || '');
        add(code, line.AccountName, {
          date: meta.date, source: meta.source, description: desc, reference: meta.reference || '',
          debit: value > 0 ? value : null, credit: value < 0 ? round2(-value) : null,
          gst: tax ? round2(Math.abs(tax)) : 0,
          gstRate: tm.rate != null ? tm.rate : (net ? round2(Math.abs(tax / net) * 100) : 0),
          gstRateName: tm.name || line.TaxType || '',
        });
      }
    }

    // Bank transactions: SPEND -> debit expense, RECEIVE -> credit income.
    for (const bt of (data.BankTransactions || [])) {
      const isSpend = String(bt.Type || '').indexOf('SPEND') === 0;
      postLines(bt.LineItems, bt.LineAmountTypes, {
        date: parseJournalDate(bt.Date), source: isSpend ? 'Spend Money' : 'Receive Money',
        reference: bt.Reference, contact: (bt.Contact && bt.Contact.Name) || 'Unknown', sign: isSpend ? 1 : -1,
      });
    }
    // Invoices: line items hit income/expense; the gross total hits the control
    // account — AR (sales invoices, debit) or AP (bills, credit).
    for (const inv of (data.Invoices || [])) {
      const isBill = inv.Type === 'ACCPAY';
      const date = parseJournalDate(inv.Date);
      const contact = (inv.Contact && inv.Contact.Name) || '';
      const ref = inv.Reference || inv.InvoiceNumber || '';
      postLines(inv.LineItems, inv.LineAmountTypes, {
        date, source: isBill ? 'Bill' : 'Sales Invoice', reference: ref, contact, sign: isBill ? 1 : -1,
      });
      const total = Number(inv.Total) || 0;
      const ctrl = isBill ? ap : ar;
      if (total && ctrl) add(ctrl.code, ctrl.name, {
        date, source: isBill ? 'Bill' : 'Sales Invoice', description: contact, reference: inv.InvoiceNumber || inv.Reference || '',
        debit: isBill ? null : total, credit: isBill ? total : null, gst: 0, gstRate: 0, gstRateName: '',
      });
    }
    // Payments settle the control accounts: receipts credit AR, bill payments debit AP.
    for (const p of (data.Payments || [])) {
      const amt = Number(p.Amount) || 0;
      if (!amt) continue;
      const inv = p.Invoice || {};
      const contact = (inv.Contact && inv.Contact.Name) || '';
      const date = parseJournalDate(p.Date);
      const ref = p.Reference || inv.InvoiceNumber || '';
      if (p.PaymentType === 'ACCRECPAYMENT' && ar) add(ar.code, ar.name, { date, source: 'Payment', description: contact, reference: ref, debit: null, credit: amt, gst: 0, gstRate: 0, gstRateName: '' });
      else if (p.PaymentType === 'ACCPAYPAYMENT' && ap) add(ap.code, ap.name, { date, source: 'Payment', description: contact, reference: ref, debit: amt, credit: null, gst: 0, gstRate: 0, gstRateName: '' });
    }
    // Manual journals: LineAmount sign sets debit/credit.
    for (const mj of (data.ManualJournals || [])) {
      const date = parseJournalDate(mj.Date);
      for (const line of (mj.JournalLines || [])) {
        if (!line.AccountCode) continue;
        const amt = Number(line.LineAmount) || 0, tax = Number(line.TaxAmount) || 0;
        const tm = taxMap[line.TaxType] || {};
        add(line.AccountCode, line.AccountName, {
          date, source: 'Manual Journal', description: line.Description || mj.Narration || '', reference: '',
          debit: amt > 0 ? amt : null, credit: amt < 0 ? round2(-amt) : null,
          gst: tax ? round2(Math.abs(tax)) : 0,
          gstRate: tm.rate != null ? tm.rate : 0, gstRateName: tm.name || line.TaxType || '',
        });
      }
    }

    const list = [...accounts.values()];
    for (const acc of list) {
      acc.lines.sort((a, b) => (a.date - b.date) || (a.seq - b.seq));
      let bal = 0, td = 0, tc = 0;
      for (const ln of acc.lines) {
        bal += (ln.debit || 0) - (ln.credit || 0);
        ln.runningBalance = round2(bal);
        td += ln.debit || 0; tc += ln.credit || 0;
      }
      acc.totalDebit = round2(td); acc.totalCredit = round2(tc); acc.net = round2(td - tc);
    }
    list.sort((a, b) => String(a.name).localeCompare(String(b.name)));

    const fromD = isoToLocalDate(o.fromDate), toD = isoToLocalDate(o.toDate);
    const period = (fromD && toD) ? `For the period ${fmtLongDate(fromD)} to ${fmtLongDate(toD)}` : '';
    return {
      kind: 'generalledger', reportName: 'General Ledger Detail',
      titles: ['General Ledger Detail', o.orgName || '', period].filter(Boolean),
      columns: GL_COLUMNS.slice(), accounts: list,
    };
  }

  /** GL model -> array-of-arrays for Excel, mirroring Xero's GL Detail layout.
   *  Returns rowMeta (per-row type) so the renderer can style each row. */
  function glToAOA(gl) {
    const aoa = [], dateCells = [], numberCells = [], rowMeta = [];
    const push = (row, type) => { aoa.push(row); rowMeta.push(type); };
    push([gl.titles[0] || 'General Ledger Detail'], 'title');
    push([gl.titles[1] || ''], 'client');
    push([gl.titles[2] || ''], 'period');
    push([], 'blank');
    push(gl.columns.slice(), 'header');
    for (const acc of gl.accounts) {
      push([acc.name, '', '', '', '', '', '', '', '', ''], 'account');
      for (const ln of acc.lines) {
        const r = aoa.length;
        const serial = excelDateSerial(ln.date);
        push([
          serial, ln.source, ln.description, ln.reference,
          ln.debit != null ? ln.debit : 0, ln.credit != null ? ln.credit : 0,
          ln.runningBalance, ln.gst || 0, ln.gstRate || 0, ln.gstRateName,
        ], 'row');
        if (serial !== '') dateCells.push({ r, c: 0 });
        [4, 5, 6, 7, 8].forEach((c) => numberCells.push({ r, c }));
      }
      push([`Total ${acc.name}`, '', '', '', acc.totalDebit, acc.totalCredit, '', '', '', ''], 'total');
      const nd = acc.net >= 0 ? acc.net : 0, nc = acc.net < 0 ? round2(-acc.net) : 0;
      push(['Net movement', '', '', '', nd, nc, '', '', '', ''], 'net');
      push([], 'blank');
    }
    return { aoa, dateCells, numberCells, rowMeta };
  }

  // ---------------------------------------------------------------------------
  // Demo data — realistic Xero-shaped responses so the whole pipeline (select →
  // dates → render → download) works before any backend / Xero connection.
  // ---------------------------------------------------------------------------
  const H = (v) => ({ Value: v });
  const ROW = (label, ...vals) => ({ RowType: 'Row', Cells: [H(label), ...vals.map(H)] });
  const SUM = (label, ...vals) => ({ RowType: 'SummaryRow', Cells: [H(label), ...vals.map(H)] });
  const HDR = (...vals) => ({ RowType: 'Header', Cells: vals.map(H) });
  const SECTION = (title, rows) => ({ RowType: 'Section', Title: title, Rows: rows });
  const wrap = (r) => ({ Reports: [r] });

  const DEMO = {
    ProfitAndLoss: wrap({
      ReportName: 'Profit and Loss', ReportType: 'ProfitAndLoss',
      ReportTitles: ['Profit and Loss', 'Demo Company (AU)', 'For the month ended 31 May 2026'],
      ReportDate: '2 June 2026',
      Rows: [
        HDR('', '31 May 2026', '31 May 2025'),
        SECTION('Trading Income', [
          ROW('Sales', '12500.00', '10800.00'),
          ROW('Other Revenue', '1500.00', '900.00'),
          SUM('Total Trading Income', '14000.00', '11700.00'),
        ]),
        SECTION('Cost of Sales', [
          ROW('Purchases', '5200.00', '4600.00'),
          SUM('Total Cost of Sales', '5200.00', '4600.00'),
        ]),
        SECTION('', [SUM('Gross Profit', '8800.00', '7100.00')]),
        SECTION('Operating Expenses', [
          ROW('Wages & Salaries', '4200.00', '3800.00'),
          ROW('Rent', '1200.00', '1200.00'),
          ROW('Office Expenses', '650.00', '540.00'),
          SUM('Total Operating Expenses', '6050.00', '5540.00'),
        ]),
        SECTION('', [SUM('Net Profit', '2750.00', '1560.00')]),
      ],
    }),

    BalanceSheet: wrap({
      ReportName: 'Balance Sheet', ReportType: 'BalanceSheet',
      ReportTitles: ['Balance Sheet', 'Demo Company (AU)', 'As at 31 May 2026'],
      ReportDate: '2 June 2026',
      Rows: [
        HDR('', '31 May 2026'),
        SECTION('Assets', [
          SECTION('Current Assets', [
            ROW('Business Bank Account', '25000.00'),
            ROW('Accounts Receivable', '8000.00'),
            SUM('Total Current Assets', '33000.00'),
          ]),
          SECTION('Fixed Assets', [
            ROW('Office Equipment', '6400.00'),
            SUM('Total Fixed Assets', '6400.00'),
          ]),
          SUM('Total Assets', '39400.00'),
        ]),
        SECTION('Liabilities', [
          SECTION('Current Liabilities', [
            ROW('Accounts Payable', '4500.00'),
            ROW('GST', '1200.00'),
            SUM('Total Current Liabilities', '5700.00'),
          ]),
          SUM('Total Liabilities', '5700.00'),
        ]),
        SECTION('', [SUM('Net Assets', '33700.00')]),
        SECTION('Equity', [
          ROW('Current Year Earnings', '2750.00'),
          ROW('Retained Earnings', '30950.00'),
          SUM('Total Equity', '33700.00'),
        ]),
      ],
    }),

    TrialBalance: wrap({
      ReportName: 'Trial Balance', ReportType: 'TrialBalance',
      ReportTitles: ['Trial Balance', 'Demo Company (AU)', 'As at 31 May 2026'],
      ReportDate: '2 June 2026',
      Rows: [
        HDR('Account', 'Debit', 'Credit'),
        SECTION('', [
          ROW('Business Bank Account (090)', '25000.00', ''),
          ROW('Accounts Receivable (610)', '8000.00', ''),
          ROW('Office Equipment (710)', '6400.00', ''),
          ROW('Accounts Payable (800)', '', '4500.00'),
          ROW('GST (820)', '', '1200.00'),
          ROW('Sales (200)', '', '12500.00'),
          ROW('Purchases (300)', '5200.00', ''),
          SUM('Total', '44600.00', '44600.00'),
        ]),
      ],
    }),

    BankSummary: wrap({
      ReportName: 'Bank Summary', ReportType: 'BankSummary',
      ReportTitles: ['Bank Summary', 'Demo Company (AU)', 'For the period 1 May 2026 to 31 May 2026'],
      ReportDate: '2 June 2026',
      Rows: [
        HDR('Bank Account', 'Opening Balance', 'Cash Received', 'Cash Spent', 'Closing Balance'),
        SECTION('', [
          ROW('Business Bank Account', '18200.00', '14000.00', '7200.00', '25000.00'),
          ROW('Petty Cash', '300.00', '0.00', '120.00', '180.00'),
          SUM('Total', '18500.00', '14000.00', '7320.00', '25180.00'),
        ]),
      ],
    }),

    ExecutiveSummary: wrap({
      ReportName: 'Executive Summary', ReportType: 'ExecutiveSummary',
      ReportTitles: ['Executive Summary', 'Demo Company (AU)', 'For the month ended 31 May 2026'],
      ReportDate: '2 June 2026',
      Rows: [
        HDR('', 'May 2026', 'Apr 2026', 'Variance'),
        SECTION('Cash', [
          ROW('Cash received', '14000.00', '12600.00', '11.1%'),
          ROW('Cash spent', '7320.00', '6800.00', '7.6%'),
        ]),
        SECTION('Profitability', [
          ROW('Income', '14000.00', '12600.00', '11.1%'),
          ROW('Gross profit', '8800.00', '7900.00', '11.4%'),
          ROW('Net profit', '2750.00', '2100.00', '31.0%'),
        ]),
        SECTION('Balance Sheet', [
          ROW('Debtors', '8000.00', '7400.00', '8.1%'),
          ROW('Creditors', '4500.00', '4200.00', '7.1%'),
        ]),
      ],
    }),

    BudgetSummary: wrap({
      ReportName: 'Budget Summary', ReportType: 'BudgetSummary',
      ReportTitles: ['Budget Summary', 'Demo Company (AU)', 'For the period ending 31 May 2026'],
      ReportDate: '2 June 2026',
      Rows: [
        HDR('', 'Mar 2026', 'Apr 2026', 'May 2026'),
        SECTION('Income', [
          ROW('Sales', '11000.00', '12000.00', '12500.00'),
          SUM('Total Income', '11000.00', '12000.00', '12500.00'),
        ]),
        SECTION('Expenses', [
          ROW('Operating Expenses', '5400.00', '5800.00', '6050.00'),
          SUM('Total Expenses', '5400.00', '5800.00', '6050.00'),
        ]),
        SECTION('', [SUM('Surplus / (Deficit)', '5600.00', '6200.00', '6450.00')]),
      ],
    }),
  };

  const DEMO_TENANTS = [
    { tenantId: 'demo-tenant-001', tenantName: 'Demo Company (AU)', tenantType: 'ORGANISATION' },
    { tenantId: 'demo-tenant-002', tenantName: 'Acme Pty Ltd', tenantType: 'ORGANISATION' },
    { tenantId: 'demo-tenant-003', tenantName: 'Riverside Cafe Trust', tenantType: 'ORGANISATION' },
  ];

  // Synthetic sub-ledger data so Demo mode can render a General Ledger.
  const DEMO_SUBLEDGERS = (function () {
    const ms = (y, mo, d) => `/Date(${Date.UTC(y, mo - 1, d)}+0000)/`;
    const BankTransactions = [];
    function spend(y, mo, d, code, net, taxType, rate, desc, ref, contact) {
      BankTransactions.push({
        Type: 'SPEND', Status: 'AUTHORISED', Date: ms(y, mo, d), Reference: ref || '', LineAmountTypes: 'Exclusive',
        Contact: contact ? { Name: contact } : undefined,
        LineItems: [{ AccountCode: code, Description: desc, LineAmount: net, TaxType: taxType, TaxAmount: round2(net * rate / 100) }],
      });
    }
    spend(2025, 8, 12, '404', 723.34, 'INPUT', 10, 'CHEQUE 0000722', '0000722', null);
    spend(2025, 10, 3, '404', 181.82, 'INPUT', 10, 'CHEQUE 0000730', '', 'Vicpass');
    spend(2025, 11, 18, '404', 651.67, 'INPUT', 10, 'CHEQUE 0000730', '', 'Vicpass');
    spend(2026, 2, 9, '404', 1219.98, 'INPUT', 10, 'CHEQUE 0000750', '0000750- Vicpass', null);
    spend(2025, 7, 30, '402', 10.00, 'EXEMPTEXPENSES', 0, 'ACCOUNT FEES', 'ACCOUNT FEES', null);
    spend(2025, 9, 30, '402', 40.00, 'EXEMPTEXPENSES', 0, 'Service fee', 'SERVICE FEE', null);
    spend(2025, 12, 31, '402', 12.00, 'EXEMPTEXPENSES', 0, 'BANK CHQ ISSUE FEE', '', null);
    spend(2025, 7, 5, '449', 168.20, 'INPUT', 10, 'Fuel & Lubricants', 'PETROGAS', null);
    spend(2025, 7, 8, '449', 59.50, 'INPUT', 10, 'Fuel & Lubricants', 'LIBERTY', null);
    spend(2025, 7, 11, '449', 51.40, 'INPUT', 10, 'Fuel & Lubricants', 'LIBERTY', null);
    const Invoices = [
      { Type: 'ACCREC', Status: 'AUTHORISED', Date: ms(2025, 9, 15), InvoiceNumber: 'INV-1042', Total: 1100, LineAmountTypes: 'Exclusive', Contact: { Name: 'Acme Corp' },
        LineItems: [{ AccountCode: '200', Description: 'Consulting work', LineAmount: 1000, TaxType: 'OUTPUT', TaxAmount: 100 }] },
      { Type: 'ACCPAY', Status: 'AUTHORISED', Date: ms(2025, 10, 2), InvoiceNumber: 'BILL-559', Reference: '559', Total: 550, LineAmountTypes: 'Exclusive', Contact: { Name: 'Detour Transport' },
        LineItems: [{ AccountCode: '449', Description: 'Cartage', LineAmount: 500, TaxType: 'INPUT', TaxAmount: 50 }] },
    ];
    const Payments = [
      { PaymentType: 'ACCRECPAYMENT', Status: 'AUTHORISED', Date: ms(2025, 10, 10), Amount: 1100, Invoice: { InvoiceNumber: 'INV-1042', Contact: { Name: 'Acme Corp' } } },
      { PaymentType: 'ACCPAYPAYMENT', Status: 'AUTHORISED', Date: ms(2025, 11, 5), Amount: 550, Invoice: { InvoiceNumber: 'BILL-559', Contact: { Name: 'Detour Transport' } } },
    ];
    const Accounts = [
      { Code: '404', Name: 'Accounting Fees' }, { Code: '402', Name: 'Bank Fees' }, { Code: '449', Name: 'Fuel & Lubricants' },
      { Code: '200', Name: 'Sales' },
      { Code: '610', Name: 'Accounts Receivable', SystemAccount: 'DEBTORS' },
      { Code: '800', Name: 'Accounts Payable', SystemAccount: 'CREDITORS' },
    ];
    const TaxRates = [
      { TaxType: 'INPUT', Name: 'GST on Expenses', EffectiveRate: 10 },
      { TaxType: 'OUTPUT', Name: 'GST on Income', EffectiveRate: 10 },
      { TaxType: 'EXEMPTEXPENSES', Name: 'GST Free Expenses', EffectiveRate: 0 },
    ];
    return { Accounts, TaxRates, BankTransactions, Invoices, Payments, ManualJournals: [] };
  })();

  const getDemoReport = (type) => DEMO[type] || DEMO.ProfitAndLoss;

  return {
    EXCEL_NUM_FMT,
    GL_COLUMNS,
    parseReport,
    buildAOA,
    toNumberOrString,
    formatAmount,
    sanitizeSheetName,
    mapSource,
    excelDateSerial,
    buildGeneralLedger,
    glToAOA,
    getDemoReport,
    DEMO,
    DEMO_TENANTS,
    DEMO_SUBLEDGERS,
  };
}));
