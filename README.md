# Xero Report Downloader

A web tool for accountants to **bulk-download Xero reports**: pick a client, tick
the reports you want, set the dates (with optional prior-period/prior-year
comparisons), and export everything as **one combined PDF/Excel** or a **ZIP of
separate files** — formatted to mirror Xero's own exports.

> **Status:** working testing build. The interface, the OAuth flow, the Xero API
> proxy, and the PDF/Excel/ZIP export engine are all implemented. Connect your
> Xero app credentials (below) to pull live data, or flip on **Demo mode** to try
> the whole interface with sample data — no Xero connection required.

---

## Why there's a backend (and what this means for hosting)

The Xero API **does not send CORS headers**, so a browser can never call it
directly — not from a local file, not from GitHub Pages. Every Xero integration
needs a small server in between. This repo is therefore two parts:

| Part | Folder | What it does | Where it can live |
|------|--------|--------------|-------------------|
| **Frontend** | `public/` | The UI + the PDF/Excel/ZIP builder (runs entirely in the browser) | Static host — **GitHub Pages is fine** |
| **Backend** | `backend/` | Runs the Xero OAuth flow and proxies the report calls | Any Node host (local now; Render/Railway/Fly/Cloudflare later) |

So your goal of adding this as a **tab in the GitHub Pages loan-planner works** —
the page is self-contained and drops in as a tab. You just also need the little
backend running somewhere and point the page at it (one constant — see
[Hosting later](#hosting-it-later)).

A note on report fidelity: Xero's API returns report **data** (a generic
rows/cells JSON), **not** the finished PDF. This tool rebuilds the PDF/Excel from
that data to match Xero's layout closely (see the verified samples via
`npm run render-test`). It's very close, not byte-identical — send more example
exports and the styling can be tuned further.

---

## 1. Create your Xero app (~5 minutes)

1. Go to **<https://developer.xero.com/app/manage>** → **New app**.
2. Integration type: **Web app** *(this is the one that supports many client
   organisations — do **not** use "Custom connection", which is single-org).*
3. **App name:** e.g. `Plus1 Report Downloader` (shown on the consent screen).
4. **Company or application URL:** any URL you own.
5. **Redirect URI:** `http://localhost:3000/callback`
   *(we'll add the production HTTPS URL later — Xero allows several).*
6. Create the app, then open **Configuration** and copy the **Client id**, and
   **Generate a secret** and copy the **Client secret** (shown once).

You do **not** need a separate app per client, and there is **no "tenant key" to
hardcode** — after you sign in, the tool discovers every organisation you've
authorised automatically. Authorise each client org once (the first time, Xero
asks which organisation to connect) and it appears in the dropdown.

---

## 2. Run it locally

Requires **Node.js 18+**.

```bash
cd backend
cp .env.example .env          # then edit .env and paste your Client ID + Secret
npm install
npm start
```

Open **<http://localhost:3000>**.

`.env` is git-ignored, so your credentials never get committed. Minimum `.env`:

```ini
XERO_CLIENT_ID=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
XERO_CLIENT_SECRET=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
XERO_REDIRECT_URI=http://localhost:3000/callback
```

### Using it

1. **Connect to Xero** → sign in and authorise. (Authorise each client org once.)
2. **Select client** from the Xero-style dropdown.
3. **Choose reports & dates** — tick reports, set a master period, optionally
   switch on *individual dates per report*, and pick a comparison
   (previous period / previous year / custom).
4. **Export** — PDF or Excel, as one combined file or a ZIP of separate files
   (combined Excel = one sheet per report).

### Demo mode

Toggle **Demo mode** (top-right) to explore the full interface and try every
export with built-in sample data — handy before your Xero credentials are set
up, or for showing the tool to colleagues.

---

## Reports supported

Profit and Loss · Balance Sheet · Trial Balance · Bank Summary · Executive
Summary · Budget Summary. Comparison periods apply to P&L, Balance Sheet and
Budget Summary (the reports Xero's API accepts `periods`/`timeframe` on).

*(Aged Receivables/Payables are per-contact in the API and are a planned
follow-up — see Roadmap.)*

---

## Project structure

```
public/
  index.html        # the app UI (drop-in tab) + export orchestration
  report-core.js    # Xero JSON -> normalised model (depth/indent, numbers)  [browser + Node]
  report-render.js  # normalised model -> PDF (jsPDF) / Excel (SheetJS)       [browser + Node]
backend/
  server.js         # Express: OAuth flow + /api proxy + serves public/
  .env.example      # copy to .env with your credentials
  smoke-test.mjs    # unit-checks the core transform        (npm run smoke)
  render-test.mjs   # renders sample PDF/Excel to ../tmp-export (npm run render-test)
```

`report-core.js` and `report-render.js` run in both the browser and Node, so the
exact rendering that ships is what the tests exercise:

```bash
cd backend
npm run smoke         # validates parsing/indentation/number-formatting
npm run render-test   # writes sample-combined.pdf / .xlsx to ../tmp-export/
```

---

## Hosting it later

When you're ready to put this next to the loan-planner:

1. **Frontend:** copy `public/` into your site (e.g. the GitHub Pages repo) and
   embed `index.html` as a tab. At the top of its script, set
   `const API_BASE = "https://your-backend-url";`.
2. **Backend:** deploy `backend/` to any Node host (Render, Railway, Fly.io, or
   adapt to Cloudflare Workers). Set the env vars there, register the host's
   `https://.../callback` as a Redirect URI on the Xero app, and set
   `ALLOWED_ORIGIN` to your GitHub Pages origin to lock the proxy down.

---

## Security notes (this is a testing build)

- Tokens are held in memory for a **single signed-in user** — fine for local
  testing; production needs real per-user sessions.
- The proxy currently allows any origin (`ALLOWED_ORIGIN=*`). Restrict it before
  exposing the backend publicly.
- Never commit `.env`. (It's git-ignored.)

---

## Roadmap / known gaps

- **Excel bold styling** — structure, indentation, numbers and number-formats
  match Xero; bold section headers/totals aren't written yet (SheetJS community
  limitation). Easy to add via a styling-capable writer if wanted.
- **Aged Receivables/Payables** — the API is per-contact; a summary view needs
  fan-out across contacts.
- **Comparison semantics** — `periods`/`timeframe` are passed straight to Xero;
  the preset mapping will be tuned against live data.
