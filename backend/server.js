/*
 * Xero Report Downloader — local backend proxy
 * -------------------------------------------------------------------------
 * Why this exists: the Xero API does NOT send CORS headers, so a browser can
 * never call it directly. This tiny Express server sits in front of Xero and:
 *   1. Runs the OAuth 2.0 "Authorization Code" flow (keeps the client secret
 *      server-side) and holds the resulting tokens.
 *   2. Proxies the calls the frontend needs (/connections + /Reports/*),
 *      attaching the access token + Xero-tenant-id header.
 *   3. Serves the static frontend in ../public.
 *
 * TESTING-ONLY NOTES:
 *   - Tokens are kept in a single in-memory slot (one signed-in user at a time).
 *     That is fine for local testing; production needs per-user sessions.
 *   - CORS defaults to "*" so you can also open the frontend from elsewhere
 *     during testing. Lock ALLOWED_ORIGIN down before exposing this anywhere.
 * -------------------------------------------------------------------------
 */
import dotenv from 'dotenv';
import express from 'express';
import crypto from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
// Load backend/.env regardless of which directory the process is launched from.
dotenv.config({ path: path.join(__dirname, '.env') });

const {
  XERO_CLIENT_ID,
  XERO_CLIENT_SECRET,
  XERO_REDIRECT_URI = 'http://localhost:3000/callback',
  // Granular report scopes (required for apps created on/after 2 Mar 2026; also
  // valid for older apps). The broad "accounting.reports.read" no longer works
  // for new apps and triggers invalid_scope.
  // The General Ledger is reconstructed from sub-ledgers (bank transactions,
  // invoices/bills, payments, manual journals) because the Journals API requires
  // a premium Xero plan. These read-only granular scopes work on any plan.
  // Payroll scopes (payroll.payruns.read + payroll.settings.read) let the GL
  // include pay-run journals (wages, PAYG, super, net pay). They're optional:
  // if the org has no payroll or the scope isn't granted, the payroll fetch is
  // skipped and the rest of the GL still builds.
  XERO_SCOPES = 'openid profile email offline_access accounting.reports.profitandloss.read accounting.reports.balancesheet.read accounting.reports.trialbalance.read accounting.reports.banksummary.read accounting.reports.executivesummary.read accounting.budgets.read accounting.settings.read accounting.banktransactions.read accounting.invoices.read accounting.manualjournals.read accounting.payments.read payroll.payruns.read payroll.settings.read',
  PORT = 3000,
  ALLOWED_ORIGIN = '*',
} = process.env;

// --- Xero endpoints -------------------------------------------------------
const XERO_AUTHORIZE_URL = 'https://login.xero.com/identity/connect/authorize';
const XERO_TOKEN_URL = 'https://identity.xero.com/connect/token';
const XERO_CONNECTIONS_URL = 'https://api.xero.com/connections';
const XERO_API_BASE = 'https://api.xero.com/api.xro/2.0';
const XERO_PAYROLL_BASE = 'https://api.xero.com/payroll.xro/1.0'; // AU Payroll

// Whitelisted query params per report type. Anything not listed is dropped so
// the frontend can never push unexpected params at Xero.
const REPORT_PARAMS = {
  ProfitAndLoss: ['fromDate', 'toDate', 'periods', 'timeframe', 'trackingCategoryID', 'trackingOptionID', 'standardLayout', 'paymentsOnly'],
  BalanceSheet: ['date', 'periods', 'timeframe', 'trackingOptionID1', 'trackingOptionID2', 'standardLayout', 'paymentsOnly'],
  TrialBalance: ['date', 'paymentsOnly'],
  BankSummary: ['fromDate', 'toDate'],
  ExecutiveSummary: ['date'],
  BudgetSummary: ['date', 'periods', 'timeframe'],
};

if (!XERO_CLIENT_ID || !XERO_CLIENT_SECRET) {
  console.warn(
    '\n⚠  XERO_CLIENT_ID / XERO_CLIENT_SECRET are not set.\n' +
    '   Copy backend/.env.example to backend/.env and fill them in.\n' +
    '   The page still loads in Demo mode, but "Connect to Xero" will not work.\n'
  );
}

// --- In-memory token store (single user, testing only) --------------------
let tokenStore = null; // { access_token, refresh_token, expires_at }
const pendingStates = new Set();

const app = express();
app.use(express.json());

// Minimal CORS. No cookies are used (tokens live server-side), so reflecting a
// configurable origin is safe for a local testing proxy.
app.use((req, res, next) => {
  res.set('Access-Control-Allow-Origin', ALLOWED_ORIGIN);
  res.set('Vary', 'Origin');
  res.set('Access-Control-Allow-Headers', 'Content-Type');
  res.set('Access-Control-Allow-Methods', 'GET,POST,DELETE,OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

const basicAuthHeader = () =>
  'Basic ' + Buffer.from(`${XERO_CLIENT_ID}:${XERO_CLIENT_SECRET}`).toString('base64');

/** POST to Xero's token endpoint and cache the result. */
async function requestToken(params) {
  const res = await fetch(XERO_TOKEN_URL, {
    method: 'POST',
    headers: {
      Authorization: basicAuthHeader(),
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: new URLSearchParams(params).toString(),
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`Xero token request failed (${res.status}): ${text}`);
  const data = JSON.parse(text);
  tokenStore = {
    access_token: data.access_token,
    refresh_token: data.refresh_token || (tokenStore && tokenStore.refresh_token),
    expires_at: Date.now() + (data.expires_in || 1800) * 1000,
  };
  return tokenStore;
}

/** Return a valid access token, transparently refreshing if it has expired. */
async function getValidAccessToken() {
  if (!tokenStore) {
    const err = new Error('Not connected to Xero. Click "Connect to Xero" first.');
    err.status = 401;
    throw err;
  }
  if (Date.now() > tokenStore.expires_at - 30_000) {
    if (!tokenStore.refresh_token) {
      tokenStore = null;
      const err = new Error('Session expired. Please reconnect to Xero.');
      err.status = 401;
      throw err;
    }
    await requestToken({ grant_type: 'refresh_token', refresh_token: tokenStore.refresh_token });
  }
  return tokenStore.access_token;
}

// --- Auth routes ----------------------------------------------------------
app.get('/auth/login', (req, res) => {
  if (!XERO_CLIENT_ID) return res.status(500).send('Server is missing XERO_CLIENT_ID (see backend/.env).');
  const state = crypto.randomBytes(16).toString('hex');
  pendingStates.add(state);
  // Build the query manually so spaces in `scope` are percent-encoded as %20.
  // (URLSearchParams encodes them as "+", which Xero rejects with invalid_scope.)
  const params = {
    response_type: 'code',
    client_id: XERO_CLIENT_ID,
    redirect_uri: XERO_REDIRECT_URI,
    scope: XERO_SCOPES,
    state,
  };
  const qs = Object.entries(params).map(([k, v]) => `${k}=${encodeURIComponent(v)}`).join('&');
  res.redirect(`${XERO_AUTHORIZE_URL}?${qs}`);
});

async function handleCallback(req, res) {
  const { code, state, error, error_description: desc } = req.query;
  if (error) return res.status(400).send(`Xero returned an error: ${error} — ${desc || ''}`);
  if (!code || !state || !pendingStates.has(state)) {
    return res.status(400).send('Invalid or expired sign-in state. Please try "Connect to Xero" again.');
  }
  pendingStates.delete(state);
  try {
    await requestToken({ grant_type: 'authorization_code', code, redirect_uri: XERO_REDIRECT_URI });
    res.redirect('/?connected=1');
  } catch (e) {
    res.status(500).send('Failed to complete Xero sign-in: ' + e.message);
  }
}
// Accept either path so it still works if the redirect URI was registered as
// /callback or /auth/callback in the Xero app.
app.get('/callback', handleCallback);
app.get('/auth/callback', handleCallback);

app.get('/auth/status', (req, res) => {
  res.json({
    connected: !!tokenStore,
    configured: !!(XERO_CLIENT_ID && XERO_CLIENT_SECRET),
  });
});

app.post('/auth/logout', (req, res) => {
  tokenStore = null;
  res.json({ ok: true });
});

// --- API proxy routes -----------------------------------------------------
app.get('/api/connections', async (req, res) => {
  try {
    const token = await getValidAccessToken();
    const r = await fetch(XERO_CONNECTIONS_URL, {
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
    });
    const body = await r.text();
    res.status(r.status).type('application/json').send(body);
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

// Disconnect a tenant from this app (frees a connection slot). `id` is the
// connection id from GET /connections — NOT the tenantId.
app.delete('/api/connections/:id', async (req, res) => {
  try {
    const token = await getValidAccessToken();
    const r = await fetch(`${XERO_CONNECTIONS_URL}/${encodeURIComponent(req.params.id)}`, {
      method: 'DELETE',
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
    });
    if (r.ok || r.status === 204) return res.json({ ok: true });
    const body = await r.text();
    res.status(r.status).json({ error: body || `HTTP ${r.status}` });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

app.get('/api/report', async (req, res) => {
  const { tenantId, type } = req.query;
  if (!tenantId) return res.status(400).json({ error: 'Missing tenantId' });
  const allowed = REPORT_PARAMS[type];
  if (!allowed) return res.status(400).json({ error: `Unknown or unsupported report type: ${type}` });
  try {
    const token = await getValidAccessToken();
    const url = new URL(`${XERO_API_BASE}/Reports/${type}`);
    for (const key of allowed) {
      const v = req.query[key];
      if (v !== undefined && v !== '') url.searchParams.set(key, v);
    }
    const r = await fetch(url.toString(), {
      headers: {
        Authorization: `Bearer ${token}`,
        'Xero-tenant-id': tenantId,
        Accept: 'application/json',
      },
    });
    const body = await r.text();
    res.status(r.status).type('application/json').send(body);
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

// Fetch all pages (100/page) of a list endpoint, optionally with a `where` filter.
async function fetchAllPages(token, tenantId, endpoint, key, where, base = XERO_API_BASE) {
  const out = [];
  for (let page = 1; page <= 100; page++) {
    const url = new URL(`${base}/${endpoint}`);
    url.searchParams.set('page', String(page));
    if (where) url.searchParams.set('where', where);
    const r = await fetch(url.toString(), {
      headers: { Authorization: `Bearer ${token}`, 'Xero-tenant-id': tenantId, Accept: 'application/json' },
    });
    if (!r.ok) { const t = await r.text(); const e = new Error(t || `HTTP ${r.status}`); e.status = r.status; e.endpoint = endpoint; throw e; }
    const arr = (JSON.parse(await r.text())[key]) || [];
    out.push(...arr);
    if (arr.length < 100) break;
  }
  return out;
}

// Build a Xero `where` date clause, e.g. Date>=DateTime(2025,7,1)&&Date<=DateTime(2026,6,30)
function dateClause(fromDate, toDate) {
  const fmt = (s) => { const d = new Date(s); return `DateTime(${d.getUTCFullYear()},${d.getUTCMonth() + 1},${d.getUTCDate()})`; };
  const parts = [];
  if (fromDate) parts.push(`Date>=${fmt(fromDate)}`);
  if (toDate) parts.push(`Date<=${fmt(toDate)}`);
  return parts.join('&&');
}

// The Payroll API doesn't take a `where` date filter on pay runs, so keep only
// POSTED runs whose PaymentDate falls in the period. Dates arrive as MS-JSON
// ("/Date(ms+0000)/") or ISO.
function filterPayRuns(runs, fromDate, toDate) {
  const toDateObj = (s) => { if (!s) return null; const m = /\/Date\((-?\d+)/.exec(String(s)); const d = m ? new Date(parseInt(m[1], 10)) : new Date(s); return isNaN(d.getTime()) ? null : d; };
  const from = fromDate ? new Date(`${fromDate}T00:00:00Z`) : null;
  const to = toDate ? new Date(`${toDate}T23:59:59Z`) : null;
  return (runs || []).filter((r) => {
    if (r.PayRunStatus && r.PayRunStatus !== 'POSTED') return false;
    const d = toDateObj(r.PaymentDate || r.PayRunPeriodEndDate);
    if (!d) return true;
    if (from && d < from) return false;
    if (to && d > to) return false;
    return true;
  });
}

// General Ledger source: the Journals API needs a premium plan, so we
// reconstruct the GL from sub-ledgers. Fetches the chart of accounts + tax rates
// (for names) plus posted bank transactions, invoices/bills, credit notes,
// payments, bank transfers, manual journals and payroll pay runs for the period.
// The frontend assembles these into the GL (see report-core).
app.get('/api/generalledger', async (req, res) => {
  const { tenantId, fromDate, toDate } = req.query;
  if (!tenantId) return res.status(400).json({ error: 'Missing tenantId' });
  const dc = dateClause(fromDate, toDate);
  const withStatus = (status) => [dc, status].filter(Boolean).join('&&');
  try {
    const token = await getValidAccessToken();
    const get1 = async (endpoint, key, base = XERO_API_BASE) => {
      const r = await fetch(`${base}/${endpoint}`, { headers: { Authorization: `Bearer ${token}`, 'Xero-tenant-id': tenantId, Accept: 'application/json' } });
      if (!r.ok) { const t = await r.text(); const e = new Error(t || `HTTP ${r.status}`); e.status = r.status; e.endpoint = endpoint; throw e; }
      return (JSON.parse(await r.text())[key]) || [];
    };
    // Resilient: if one sub-ledger 403s/errors, skip it rather than fail the whole GL.
    const safe = async (label, fn, fallback = []) => { try { return await fn(); } catch (e) { console.warn(`GL: ${label} unavailable -`, e.message); return fallback; } };
    const page = (endpoint, key, where) => fetchAllPages(token, tenantId, endpoint, key, where);

    const [Accounts, TaxRates] = await Promise.all([
      safe('Accounts', () => get1('Accounts', 'Accounts')),
      safe('TaxRates', () => get1('TaxRates', 'TaxRates')),
    ]);
    const BankTransactions = await safe('BankTransactions', () => page('BankTransactions', 'BankTransactions', withStatus('Status=="AUTHORISED"')));
    const Invoices = await safe('Invoices', () => page('Invoices', 'Invoices', withStatus('(Status=="AUTHORISED"||Status=="PAID")')));
    const CreditNotes = await safe('CreditNotes', () => page('CreditNotes', 'CreditNotes', withStatus('(Status=="AUTHORISED"||Status=="PAID")')));
    const Payments = await safe('Payments', () => page('Payments', 'Payments', withStatus('Status=="AUTHORISED"')));
    const BankTransfers = await safe('BankTransfers', () => page('BankTransfers', 'BankTransfers', dc));
    const ManualJournals = await safe('ManualJournals', () => page('ManualJournals', 'ManualJournals', withStatus('Status=="POSTED"')));
    // Payroll (AU): pay-run totals + the account mappings needed to post them.
    // Optional — skipped cleanly if there's no payroll or the scope isn't granted.
    const allPayRuns = await safe('PayRuns', () => fetchAllPages(token, tenantId, 'PayRuns', 'PayRuns', null, XERO_PAYROLL_BASE));
    const PayRuns = filterPayRuns(allPayRuns, fromDate, toDate);
    const PayItems = await safe('PayItems', () => get1('PayItems', 'PayItems', XERO_PAYROLL_BASE), {});
    const PayrollSettings = await safe('PayrollSettings', () => get1('Settings', 'Settings', XERO_PAYROLL_BASE), {});
    res.json({ Accounts, TaxRates, BankTransactions, Invoices, CreditNotes, Payments, BankTransfers, ManualJournals, PayRuns, PayItems, PayrollSettings });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message, endpoint: e.endpoint });
  }
});

// --- Static frontend ------------------------------------------------------
app.use(express.static(path.join(__dirname, '..', 'public')));

app.listen(PORT, () => {
  console.log(`\n✅ Xero Report Downloader running at  http://localhost:${PORT}`);
  console.log(`   Redirect URI in use:  ${XERO_REDIRECT_URI}`);
  console.log(`   Scopes:  ${XERO_SCOPES}`);
  if (!XERO_CLIENT_ID) console.log('   (Demo mode only — add credentials to backend/.env to connect to Xero.)');
  console.log('');
});
