# LedgerFlow Pro — Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Microsoft Excel Add-in for everyday accountants that automates reconciliation, month-end close, payroll, VAT, financial reporting, and fixed assets — with all financial data encrypted client-side (AES-256-GCM) and stored exclusively in the user's own OneDrive via Microsoft Graph API.

**Architecture:** Office.js TaskPane add-in (Webpack + vanilla JS) communicates with Excel via `Excel.run()`. All financial data is encrypted in the browser using the Web Crypto API before being written to the user's OneDrive folder (`LedgerFlow/` in their personal OneDrive). The Express server stores only auth tokens and subscription records — it never touches client data.

**Tech Stack:** Office.js, Webpack 5, MSAL.js v3, Microsoft Graph API, Web Crypto API (AES-256-GCM / PBKDF2), Express.js, MS SQL (mssql), JWT, vanilla JS (ES modules)

---

## Repository Structure

```
LedgerFlow/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html         # Main add-in HTML shell
│   │   └── taskpane.js           # Monolith UI — all feature setup*() functions
│   ├── commands/
│   │   └── commands.js           # Ribbon button handlers
│   ├── api/
│   │   └── ledgerflow.js         # API service layer (auth, subscriptions, local helpers)
│   ├── crypto/
│   │   └── crypto.js             # Web Crypto wrapper — encrypt/decrypt/deriveKey
│   └── onedrive/
│       └── onedrive.js           # MSAL + Graph API — read/write encrypted blobs
├── server/
│   ├── server.js                 # Express entry point
│   ├── db.js                     # MS SQL pool + table creation
│   ├── middleware/
│   │   └── auth.js               # verifyToken, signToken
│   └── routes/
│       ├── auth.js               # POST /login, POST /register, GET /me
│       └── subscriptions.js      # GET /plans, GET /current, POST /upgrade
├── assets/icons/                 # Add-in icon PNGs (16, 32, 80px)
├── deploy-pkg/                   # Azure deployment mirror of server/ + dist/
├── SharedManifest/manifest.xml   # Shared manifest variant
├── manifest.xml                  # Office Add-in manifest
├── webpack.config.js
├── package.json
└── web.config                    # IISNode config for Azure App Service
```

---

## Phase 1 — Project Scaffold

### Task 1.1: `package.json` + Dependencies

**Files:**
- Create: `package.json`

- [ ] Create `package.json` with all required dependencies:

```json
{
  "name": "ledgerflow-pro",
  "version": "1.0.0",
  "description": "Excel Add-in for Accountants — encrypted OneDrive storage",
  "scripts": {
    "dev": "concurrently \"npm run server\" \"webpack serve --config webpack.config.js\"",
    "server": "node server/server.js",
    "build": "webpack --config webpack.config.js --mode production",
    "sideload": "office-addin-debugging start manifest.xml",
    "validate": "office-addin-manifest validate manifest.xml"
  },
  "dependencies": {
    "@azure/msal-browser": "^3.x",
    "express": "^4.18.x",
    "mssql": "^10.x",
    "jsonwebtoken": "^9.x",
    "bcryptjs": "^2.x",
    "cors": "^2.x",
    "dotenv": "^16.x"
  },
  "devDependencies": {
    "webpack": "^5.x",
    "webpack-cli": "^5.x",
    "webpack-dev-server": "^4.x",
    "copy-webpack-plugin": "^11.x",
    "html-webpack-plugin": "^5.x",
    "concurrently": "^8.x",
    "office-addin-debugging": "^1.x",
    "office-addin-manifest": "^1.x",
    "office-addin-dev-certs": "^1.x"
  }
}
```

- [ ] Run `npm install` in `c:\Users\DELL\source\repos\LedgerFlow`
- [ ] Verify `node_modules/` created with no errors

---

### Task 1.2: Webpack Configuration

**Files:**
- Create: `webpack.config.js`

- [ ] Create `webpack.config.js` mirroring AuditFlow's config but for LedgerFlow entry points:

```js
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = (env, options) => {
  const devMode = options.mode !== "production";
  return {
    devtool: devMode ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].bundle.js",
      clean: true,
    },
    resolve: { extensions: [".js"] },
    module: { rules: [{ test: /\.js$/, exclude: /node_modules/, use: "babel-loader" }] },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({ patterns: [{ from: "assets", to: "assets" }] }),
    ],
    devServer: {
      https: true,
      port: 3000,
      headers: { "Access-Control-Allow-Origin": "*" },
    },
  };
};
```

---

### Task 1.3: `manifest.xml`

**Files:**
- Create: `manifest.xml`
- Create: `SharedManifest/manifest.xml`

- [ ] Create `manifest.xml`:
  - `DefaultLocale: en-US`
  - `ProviderName: LedgerFlow Pro`
  - `SourceLocation: https://localhost:3000/taskpane.html` (dev) → Azure URL (prod)
  - `RequestedWidth: 350`, `RequestedHeight: 550`
  - Icons referencing `assets/icons/icon-16.png`, `icon-32.png`, `icon-80.png`
  - Permissions: `ReadWriteDocument` (needed for Excel interaction)

- [ ] Copy to `SharedManifest/manifest.xml` as the shared variant
- [ ] Run `npm run validate` — expected: no errors

---

### Task 1.4: `server/db.js` — MS SQL Connection Pool

**Files:**
- Create: `server/db.js`

- [ ] Create `server/db.js` (copy pattern from AuditFlow's db.js, rename tables to `lf_tenants`, `lf_users`, `lf_plans`, `lf_subscriptions`):

```js
const sql = require("mssql");
require("dotenv").config();

const config = {
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  server: process.env.DB_SERVER,
  database: process.env.DB_NAME,
  options: { encrypt: true, trustServerCertificate: false },
};

const poolPromise = new sql.ConnectionPool(config)
  .connect()
  .then((pool) => { console.log("SQL connected"); return pool; })
  .catch((err) => { console.error("SQL connection failed:", err); process.exit(1); });

module.exports = { sql, poolPromise };
```

- [ ] Create `.env` file at repo root with: `DB_USER`, `DB_PASSWORD`, `DB_SERVER`, `DB_NAME`, `JWT_SECRET`, `PORT`, `FRONTEND_URL`, `AZURE_CLIENT_ID` (for MSAL)
- [ ] Add `.env` to `.gitignore`

---

### Task 1.5: `server/server.js` — Express Entry Point

**Files:**
- Create: `server/server.js`

- [ ] Create Express server with:
  - CORS restricted to `localhost:3000` + production Azure URL
  - `express.json()` middleware
  - Routes: `/api/auth` → `routes/auth.js`, `/api/subscriptions` → `routes/subscriptions.js`
  - Static serving of `dist/` in production
  - Named pipe support for IISNode (Azure)

---

### Task 1.6: Auth Middleware + Routes

**Files:**
- Create: `server/middleware/auth.js`
- Create: `server/routes/auth.js`
- Create: `server/routes/subscriptions.js`

- [ ] `auth.js` middleware: `verifyToken(req, res, next)` — validates `Authorization: Bearer <token>`, attaches `req.user = { sub, email, tenantId, role }`
- [ ] `signToken(payload)` helper — signs with `JWT_SECRET`, expires `8h`
- [ ] `routes/auth.js`: `POST /register`, `POST /login`, `GET /me` (protected)
- [ ] `routes/subscriptions.js`: `GET /plans`, `GET /current` (protected), `POST /upgrade` (protected)

---

### Task 1.7: `web.config` + `deploy-pkg/`

**Files:**
- Create: `web.config`

- [ ] Create `web.config` for IISNode (copy from AuditFlow, update handler path to `server/server.js`)
- [ ] Mirror `server/` into `deploy-pkg/server/` — this is the Azure deployment artifact

---

## Phase 2 — Crypto Module (Web Crypto API)

> All encryption runs in the browser. The server never sees plaintext or keys.

### Task 2.1: `src/crypto/crypto.js`

**Files:**
- Create: `src/crypto/crypto.js`

- [ ] Implement key derivation:

```js
// Derives an AES-256-GCM CryptoKey from a user password + salt using PBKDF2
export async function deriveKey(password, salt) {
  const enc = new TextEncoder();
  const keyMaterial = await crypto.subtle.importKey(
    "raw", enc.encode(password), "PBKDF2", false, ["deriveKey"]
  );
  return crypto.subtle.deriveKey(
    { name: "PBKDF2", salt: enc.encode(salt), iterations: 310_000, hash: "SHA-256" },
    keyMaterial,
    { name: "AES-GCM", length: 256 },
    false,
    ["encrypt", "decrypt"]
  );
}
```

- [ ] Implement `encrypt(key, plaintext)`:
  - Generate random 12-byte IV via `crypto.getRandomValues()`
  - Encrypt with `AES-GCM`
  - Return `{ iv: base64, ciphertext: base64 }`

- [ ] Implement `decrypt(key, iv, ciphertext)`:
  - Decode base64 IV + ciphertext
  - Decrypt with `AES-GCM`
  - Return plaintext string

- [ ] Implement `generateSalt()` — returns 16-byte random base64 string
- [ ] Key is held only in memory (`sessionStorage` max — never `localStorage`)

---

### Task 2.2: Key Session Management

**Files:**
- Modify: `src/crypto/crypto.js`

- [ ] Add `setSessionKey(key)` — stores `CryptoKey` object in a module-level variable (in-memory only, cleared on page unload)
- [ ] Add `getSessionKey()` — returns the key or throws if not set
- [ ] Add `clearSessionKey()` — wipes the in-memory key (called on logout)
- [ ] On first login: derive key from password, store encrypted salt in `localStorage` (`lf_salt`). On subsequent logins: read salt, re-derive key.

---

## Phase 3 — OneDrive Integration (MSAL + Graph API)

### Task 3.1: MSAL Configuration

**Files:**
- Create: `src/onedrive/onedrive.js`

- [ ] Configure MSAL `PublicClientApplication`:

```js
import * as msal from "@azure/msal-browser";

const msalConfig = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID,  // injected by webpack DefinePlugin
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin,
  },
  cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false },
};

const pca = new msal.PublicClientApplication(msalConfig);
```

- [ ] Implement `signInOneDrive()` — popup flow requesting `Files.ReadWrite` and `User.Read` scopes
- [ ] Implement `getGraphToken()` — silent token acquisition with popup fallback
- [ ] Implement `signOutOneDrive()` — clears MSAL cache

> **Azure App Registration required:** Register app in Azure Entra ID, add `Files.ReadWrite` delegated permission, add redirect URI for both localhost:3000 and production Azure URL.

---

### Task 3.2: Graph API File Operations

**Files:**
- Modify: `src/onedrive/onedrive.js`

- [ ] Implement `ensureFolder()` — creates `/LedgerFlow/` folder in user's OneDrive root if not exists:
  ```
  PUT https://graph.microsoft.com/v1.0/me/drive/root:/LedgerFlow:
  ```

- [ ] Implement `writeFile(filename, encryptedPayload)`:
  ```
  PUT https://graph.microsoft.com/v1.0/me/drive/root:/LedgerFlow/{filename}:/content
  Body: JSON.stringify(encryptedPayload)   // { iv, ciphertext } — unreadable blob
  ```

- [ ] Implement `readFile(filename)`:
  ```
  GET https://graph.microsoft.com/v1.0/me/drive/root:/LedgerFlow/{filename}:/content
  Returns: { iv, ciphertext }
  ```

- [ ] Implement `listFiles()` — lists all `.lf` files in `/LedgerFlow/` folder

---

### Task 3.3: High-Level Data Store API

**Files:**
- Create: `src/onedrive/store.js`

- [ ] Implement `saveData(collection, data)`:
  1. Serialize `data` to JSON string
  2. Encrypt with `getSessionKey()` via `crypto.encrypt()`
  3. Call `writeFile(`${collection}.lf`, encryptedPayload)`

- [ ] Implement `loadData(collection)`:
  1. Call `readFile(`${collection}.lf`)`
  2. Decrypt with `getSessionKey()` via `crypto.decrypt()`
  3. Return parsed JSON object
  4. On file-not-found (404): return empty default for that collection

- [ ] Implement `localFallback(collection, data)` — writes to `localStorage` as encrypted string when offline (same encrypt/decrypt, base64 stored)

---

## Phase 4 — API Service Layer

### Task 4.1: `src/api/ledgerflow.js`

**Files:**
- Create: `src/api/ledgerflow.js`

- [ ] Implement `getApiUrl()` — returns localhost in dev, Azure URL in prod
- [ ] Implement `request(method, path, body, authenticated)` — attaches Bearer token, handles 401 → `clearSession()`
- [ ] Implement auth functions: `login(email, password)`, `register(...)`, `getMe()`, `clearSession()`
- [ ] Implement data collection helpers (wrappers over `store.js`):
  - `getTransactions()` / `saveTransactions(data)`
  - `getVATRecords()` / `saveVATRecords(data)`
  - `getPayrollRuns()` / `savePayrollRuns(data)`
  - `getAssets()` / `saveAssets(data)`
  - `getBudgets()` / `saveBudgets(data)`
  - `getJournals()` / `saveJournals(data)`

---

## Phase 5 — TaskPane Shell

### Task 5.1: `src/taskpane/taskpane.html`

**Files:**
- Create: `src/taskpane/taskpane.html`

- [ ] HTML shell with:
  - Office.js CDN script tag
  - Navigation tabs: `reconcile`, `closing`, `vat`, `payroll`, `assets`, `reports`, `budget`, `settings`
  - Feature-gated tabs hidden behind `.platform-only` class
  - `#toast` div for `showToast()`
  - `#auth-panel` for login/register form
  - `#onedrive-panel` for OneDrive sign-in prompt

---

### Task 5.2: `src/taskpane/taskpane.js` — App Shell

**Files:**
- Create: `src/taskpane/taskpane.js`

- [ ] `Office.onReady()` callback calling `initApp()`
- [ ] `initApp()` calls all `setup*()` functions and restores session
- [ ] `switchTab(tabName)` — shows/hides tab panels, lazy-loads data
- [ ] `showToast(message, type)` — `success`, `error`, `info` with 3.5s auto-hide
- [ ] `showPlatformGate(tabName)` — prompt to sign in for cloud features
- [ ] Session state object:
  ```js
  const state = {
    user: null,           // decoded JWT payload
    planFeatures: [],     // from subscription
    oneDriveConnected: false,
    cryptoReady: false,   // key derived and in memory
  };
  ```

---

## Phase 6 — Feature Modules

### Task 6.1: `setupReconcile()` — Bank/Ledger Reconciliation

**Tab:** `reconcile`

**Excel columns expected:** A=Date, B=Description, C=Amount, D=Reference

- [ ] **Match Transactions button:**
  - `Excel.run()` reads two named ranges: `BankData` and `LedgerData`
  - Fuzzy-match on Amount + Date (±3 days tolerance)
  - Stamp matched rows green (`fill.color = "#C6EFCE"`), unmatched red (`"#FFC7CE"`)
  - Write match summary to a new sheet `Reconciliation_Summary`

- [ ] **Clear Matches button:** resets fill colors, removes summary sheet

- [ ] **Export Unmatched button:** copies unmatched rows to clipboard-ready range

---

### Task 6.2: `setupClosing()` — Month-End Close Suite

**Tab:** `closing`

- [ ] **Journal Validator:**
  - Read active sheet, sum Debit column vs Credit column
  - Show `✓ Balanced` or `✗ Out by {amount}` with exact difference
  - Highlight rows where debit/credit cells are blank

- [ ] **Accruals Scheduler:**
  - Input: description, amount, start month, end month, GL code
  - Generates accrual/reversal journal entries across months into a new sheet `Accruals_Schedule`

- [ ] **Prepayments Amortiser:**
  - Input: total prepaid amount, coverage months
  - Outputs monthly charge schedule with running balance

- [ ] **Close Checklist:**
  - Rendered checklist of standard close tasks (customisable)
  - Saved per period to OneDrive (`closing_YYYY_MM.lf`)
  - Checkbox state persisted

---

### Task 6.3: `setupVAT()` — VAT & Withholding Tax

**Tab:** `vat`

- [ ] **VAT Tagger:**
  - Dropdown: Standard Rate / Zero-rated / Exempt / Out-of-scope
  - Apply tag to selected rows, write to column E
  - Auto-calculate VAT amount in column F = Amount × selected rate

- [ ] **VAT Summary:**
  - Reads all tagged rows, sums Output VAT, Input VAT, Net Payable
  - Writes to `VAT_Return_Working` sheet

- [ ] **WHT Calculator:**
  - Input: gross payment, service type
  - Lookup table of WHT rates (configurable per country)
  - Returns: WHT amount, net payment, journal entry preview

---

### Task 6.4: `setupPayroll()` — Payroll Calculator

**Tab:** `payroll`

- [ ] **Employee Table Reader:**
  - Reads sheet with columns: Name, Gross Salary, Allowances, Deductions
  - Applies PAYE tax table (configurable bands), statutory deductions (SSNIT/NHIL or configurable)
  - Writes Net Pay column

- [ ] **Payslip Generator:**
  - For each employee row, generates a formatted payslip in a new sheet
  - Uses a built-in template (can be customised)

- [ ] **Tax Bands Editor:**
  - UI table in taskpane to define up to 8 income bands + rates
  - Saved to OneDrive (`payroll_config.lf`)

---

### Task 6.5: `setupAssets()` — Fixed Asset Register

**Tab:** `assets`

- [ ] **Depreciation Calculator:**
  - Input: asset name, cost, residual value, useful life (years), method (SL / Reducing Balance)
  - Outputs a full depreciation schedule to a new sheet `Depreciation_Schedule`

- [ ] **Asset Register:**
  - Maintains running register: add, retire, revalue assets
  - Calculates NBV at any period
  - Data saved to OneDrive (`assets.lf`)

- [ ] **Disposal Calculator:**
  - Input: NBV at disposal date, proceeds
  - Outputs gain/loss on disposal + journal entry

---

### Task 6.6: `setupReports()` — Financial Statements Generator

**Tab:** `reports`

- [ ] **TB Mapper:**
  - Read trial balance sheet (A=Code, B=Name, C=Debit, D=Credit)
  - UI table: map each GL code range to P&L or Balance Sheet line
  - Mapping saved to OneDrive (`tb_mapping.lf`)

- [ ] **Generate P&L:**
  - Apply mapping, aggregate by financial statement line
  - Write formatted P&L to new sheet `Income_Statement`

- [ ] **Generate Balance Sheet:**
  - Apply mapping, write formatted Balance Sheet to `Balance_Sheet` sheet
  - Validates Assets = Liabilities + Equity

- [ ] **Cash Flow (Indirect Method):**
  - Input: Net income (from P&L), working capital changes (from BS movement)
  - Generates Cash Flow Statement on new sheet `Cash_Flow`

---

### Task 6.7: `setupBudget()` — Variance Analysis

**Tab:** `budget`

- [ ] **Variance Calculator:**
  - Reads two sheets: `Actual` and `Budget` (same row structure)
  - Computes Variance (Amount) and Variance (%) per line
  - Highlights favourable variance green, adverse red
  - Threshold configurable in taskpane (default ±10%)

- [ ] **Rolling Forecast:**
  - Input: number of actuals months locked
  - Replaces future budget months with forecast (user-input or extrapolated trend)
  - Writes `Forecast` sheet

---

## Phase 7 — OneDrive Connect Flow

### Task 7.1: `setupSettings()` + OneDrive Sign-In

**Tab:** `settings`

- [ ] **Connect OneDrive button:**
  - Calls `signInOneDrive()`
  - On success: calls `ensureFolder()`, updates `state.oneDriveConnected = true`
  - Shows connected account email + `LedgerFlow/` folder path

- [ ] **Disconnect button:** calls `signOutOneDrive()`, clears `state.oneDriveConnected`

- [ ] **Encryption Key Setup:**
  - First-time: prompt user to set a data password (separate from login password)
  - Derive key via `deriveKey(dataPassword, salt)`, hold in memory
  - Store salt only (not key, not password) in `localStorage`
  - On subsequent sessions: prompt for data password to re-derive key

- [ ] **Change Data Password:**
  - Re-encrypt all OneDrive files with new key
  - Replace old salt with new salt in `localStorage`

---

## Phase 8 — Feature Gating + Subscription

### Task 8.1: Plan Features

- [ ] Plans stored in `lf_plans` table with comma-separated `features` column:
  - `free`: `"reconcile,closing,vat"` (core)
  - `pro`: `"reconcile,closing,vat,payroll,assets,reports,budget"`
  - `firm`: all + multi-user (future)

- [ ] `switchTab()` checks `state.planFeatures.includes(tabName)` before showing tab content
- [ ] Free plan still allows **Use Without Account** mode (localStorage only, no OneDrive sync)

---

## Phase 9 — Deployment

### Task 9.1: Azure App Service Setup

- [ ] Create Azure App Service (Free/B1 tier to start)
- [ ] Update `manifest.xml` production `SourceLocation` URL
- [ ] Mirror `server/` → `deploy-pkg/server/`
- [ ] Verify `web.config` IISNode handler points to `server/server.js`
- [ ] Set environment variables in Azure App Service Configuration:
  `DB_USER`, `DB_PASSWORD`, `DB_SERVER`, `DB_NAME`, `JWT_SECRET`, `AZURE_CLIENT_ID`, `FRONTEND_URL`

### Task 9.2: Azure App Registration (Entra ID)

- [ ] Register new app in Azure Entra ID portal
- [ ] Add delegated API permission: `Files.ReadWrite` (Microsoft Graph)
- [ ] Add redirect URIs: `https://localhost:3000`, production Azure URL
- [ ] Copy `Application (client) ID` → `AZURE_CLIENT_ID` env var + webpack `DefinePlugin`
- [ ] No client secret needed — this is a public client (SPA/add-in)

### Task 9.3: Production Build + First Deploy

- [ ] Run `npm run build` → verify `dist/` generated cleanly
- [ ] Deploy `deploy-pkg/` to Azure App Service via zip deploy or GitHub Actions
- [ ] Test `npm run validate` against production manifest
- [ ] Sideload in Excel desktop via `npm run sideload` and smoke-test all tabs

---

## Security Checklist

- [ ] CORS `allowedOrigins` limited to localhost:3000 + production Azure URL only
- [ ] All SQL queries use parameterized inputs (no string concatenation)
- [ ] JWT verified on every protected route via `verifyToken` middleware
- [ ] Encryption key never written to `localStorage`, `sessionStorage`, or server
- [ ] Salt stored in `localStorage` but salt alone is not useful without password
- [ ] PBKDF2 iterations set to 310,000 (OWASP 2024 recommendation for SHA-256)
- [ ] Graph API token stored in `sessionStorage` (MSAL default) — cleared on tab close
- [ ] OneDrive files contain only `{ iv, ciphertext }` — no plaintext field names visible
- [ ] `.env` in `.gitignore`
- [ ] No sensitive data in webpack bundle (use `DefinePlugin` for public config only)

---

## Milestone Summary

| Milestone | Phases | Deliverable |
|---|---|---|
| **M1 — Scaffold** | 1 | Running dev server + sideloadable add-in shell |
| **M2 — Crypto + OneDrive** | 2, 3, 4 | Encrypted read/write to user's OneDrive working |
| **M3 — Core Features** | 5, 6.1, 6.2 | Reconcile + Closing tabs functional |
| **M4 — Tax + Payroll** | 6.3, 6.4 | VAT + Payroll tabs functional |
| **M5 — Reporting** | 6.5, 6.6, 6.7 | Assets + Reports + Budget tabs functional |
| **M6 — Settings + Auth** | 7, 8 | OneDrive connect flow + subscription gating |
| **M7 — Production** | 9 | Deployed to Azure, validated manifest |
