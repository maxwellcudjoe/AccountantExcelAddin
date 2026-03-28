/* global Office, Excel */
/* eslint-disable no-unused-vars */

import * as API from "../api/ledgerflow.js";
import {
  deriveKey,
  generateSalt,
  setSessionKey,
  clearSessionKey,
  isKeyLoaded,
} from "../crypto/crypto.js";
import {
  signInOneDrive,
  signOutOneDrive,
  isOneDriveConnected,
} from "../onedrive/onedrive.js";

// ── App State ─────────────────────────────────────────────────────────────────
const state = {
  user: null,           // decoded from JWT / /me response
  planFeatures: [],     // e.g. ["reconcile","closing","vat"]
  oneDriveConnected: false,
  offlineMode: false,
};

const FREE_FEATURES = ["reconcile", "closing", "vat", "settings"];

// ── Bootstrap ─────────────────────────────────────────────────────────────────
Office.onReady(async () => {
  setupAuth();
  setupSettings();
  setupReconcile();
  setupClosing();
  setupVAT();
  setupPayroll();
  setupAssets();
  setupReports();
  setupBudget();

  // Wire nav tabs
  document.querySelectorAll(".nav-tab").forEach((tab) => {
    tab.addEventListener("click", () => switchTab(tab.dataset.tab));
  });

  // Restore session if token exists
  const token = API.getToken();
  if (token) {
    await restoreSession();
  }
});

// ── Session ───────────────────────────────────────────────────────────────────
async function restoreSession() {
  try {
    const me = await API.getMe();
    state.user = me;
    const sub = await API.getCurrentSubscription().catch(() => ({ features: FREE_FEATURES }));
    state.planFeatures = Array.isArray(sub.features)
      ? sub.features
      : (sub.features || "reconcile,closing,vat").split(",");

    showApp();
    updateUI();

    // Check OneDrive connection status
    state.oneDriveConnected = await isOneDriveConnected().catch(() => false);
    updateOneDriveUI();
  } catch {
    // Token invalid — show auth panel
    API.clearSession();
    showAuthPanel();
  }
}

function showApp() {
  document.getElementById("auth-panel").style.display = "none";
  document.getElementById("nav").style.display = "flex";
  document.getElementById("onedrive-panel").style.display = "block";
  applyFeatureGating();
}

function showAuthPanel() {
  document.getElementById("auth-panel").style.display = "block";
  document.getElementById("nav").style.display = "none";
  document.getElementById("onedrive-panel").style.display = "none";
  document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
}

function updateUI() {
  const label = document.getElementById("user-label");
  if (state.user) {
    label.textContent = state.user.full_name || state.user.email;
  } else if (state.offlineMode) {
    label.textContent = "Offline mode";
  }

  const info = document.getElementById("settings-user-info");
  if (state.user) {
    info.textContent = `${state.user.full_name} · ${state.user.email}`;
  }

  const planInfo = document.getElementById("settings-plan-info");
  if (state.planFeatures.length) {
    const isProOrFirm = state.planFeatures.includes("payroll");
    planInfo.textContent = isProOrFirm
      ? `Pro plan · Features: ${state.planFeatures.join(", ")}`
      : `Free plan · Features: ${state.planFeatures.join(", ")}`;
  }
}

function applyFeatureGating() {
  document.querySelectorAll(".nav-tab").forEach((tab) => {
    const tabName = tab.dataset.tab;
    const allowed = state.offlineMode
      ? FREE_FEATURES.includes(tabName)
      : state.planFeatures.includes(tabName) || FREE_FEATURES.includes(tabName);
    tab.classList.toggle("gated", !allowed);
  });

  // Show/hide gated content panels
  ["payroll", "assets", "reports", "budget"].forEach((name) => {
    const gate = document.getElementById(`${name}-gate`);
    const content = document.getElementById(`${name}-content`);
    const allowed = state.planFeatures.includes(name);
    if (gate) gate.style.display = allowed ? "none" : "block";
    if (content) content.style.display = allowed ? "block" : "none";
  });
}

// ── Tab switching ─────────────────────────────────────────────────────────────
function switchTab(tabName) {
  document.querySelectorAll(".nav-tab").forEach((t) =>
    t.classList.toggle("active", t.dataset.tab === tabName)
  );
  document.querySelectorAll(".tab-panel").forEach((p) =>
    p.classList.toggle("active", p.id === `tab-${tabName}`)
  );
}

// ── Toast ─────────────────────────────────────────────────────────────────────
function showToast(message, type = "info") {
  const toast = document.getElementById("toast");
  toast.textContent = message.substring(0, 100);
  toast.className = type;
  toast.style.display = "block";
  clearTimeout(showToast._timer);
  showToast._timer = setTimeout(() => (toast.style.display = "none"), 3500);
}

function showResult(id, message, type) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent = message;
  el.className = `lf-result ${type}`;
  el.style.display = "block";
}

// ── Auth setup ────────────────────────────────────────────────────────────────
function setupAuth() {
  let isRegister = false;

  const toggleLink = document.getElementById("auth-toggle-link");
  const title = document.getElementById("auth-title");
  const submitBtn = document.getElementById("auth-submit-btn");
  const fullNameField = document.getElementById("auth-fullname");
  const firmField = document.getElementById("auth-firm");

  toggleLink.addEventListener("click", () => {
    isRegister = !isRegister;
    title.textContent = isRegister ? "Create account" : "Sign in";
    submitBtn.textContent = isRegister ? "Register" : "Sign In";
    toggleLink.textContent = isRegister
      ? "Already have an account? Sign in"
      : "Don't have an account? Register";
    fullNameField.style.display = isRegister ? "block" : "none";
    firmField.style.display = isRegister ? "block" : "none";
  });

  submitBtn.addEventListener("click", async () => {
    const email = document.getElementById("auth-email").value.trim();
    const password = document.getElementById("auth-password").value;

    if (!email || !password) {
      showToast("Email and password are required", "error");
      return;
    }

    submitBtn.disabled = true;
    submitBtn.textContent = isRegister ? "Registering…" : "Signing in…";

    try {
      if (isRegister) {
        const fullName = fullNameField.value.trim();
        const firmName = firmField.value.trim();
        if (!fullName) { showToast("Full name is required", "error"); return; }
        await API.register(email, password, fullName, firmName);
      } else {
        await API.login(email, password);
      }
      await restoreSession();
      showToast("Welcome to LedgerFlow Pro!", "success");
    } catch (err) {
      showToast(err.message, "error");
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = isRegister ? "Register" : "Sign In";
    }
  });

  document.getElementById("offline-btn").addEventListener("click", () => {
    state.offlineMode = true;
    state.planFeatures = [...FREE_FEATURES];
    document.getElementById("auth-panel").style.display = "none";
    document.getElementById("nav").style.display = "flex";
    document.getElementById("onedrive-panel").style.display = "none";
    document.getElementById("user-label").textContent = "Offline mode";
    applyFeatureGating();
    switchTab("reconcile");
    showToast("Running in offline mode. Connect OneDrive in Settings to sync.", "info");
  });
}

// ── Settings setup ────────────────────────────────────────────────────────────
function setupSettings() {
  document.getElementById("settings-logout-btn").addEventListener("click", () => {
    API.clearSession();
    clearSessionKey();
    state.user = null;
    state.planFeatures = [];
    state.offlineMode = false;
    state.oneDriveConnected = false;
    showAuthPanel();
    showToast("Signed out", "info");
  });

  // OneDrive connect
  const oneDriveBtn = document.getElementById("settings-onedrive-btn");
  const disconnectBtn = document.getElementById("settings-onedrive-disconnect");

  async function connectOneDrive() {
    oneDriveBtn.disabled = true;
    oneDriveBtn.textContent = "Connecting…";
    try {
      const { email } = await signInOneDrive();
      state.oneDriveConnected = true;
      updateOneDriveUI(email);
      showToast("OneDrive connected!", "success");
    } catch (err) {
      showToast(`OneDrive error: ${err.message}`, "error");
    } finally {
      oneDriveBtn.disabled = false;
      oneDriveBtn.textContent = "Connect OneDrive";
    }
  }

  oneDriveBtn.addEventListener("click", connectOneDrive);
  document.getElementById("onedrive-connect-btn").addEventListener("click", connectOneDrive);

  disconnectBtn.addEventListener("click", async () => {
    await signOutOneDrive();
    state.oneDriveConnected = false;
    updateOneDriveUI();
    showToast("OneDrive disconnected", "info");
  });

  // Encryption key
  const dataPwdBtn = document.getElementById("settings-data-pwd-btn");
  const dataPwdInput = document.getElementById("settings-data-pwd");
  const cryptoStatus = document.getElementById("settings-crypto-status");

  dataPwdBtn.addEventListener("click", async () => {
    const pwd = dataPwdInput.value;
    if (!pwd || pwd.length < 8) {
      showToast("Data password must be at least 8 characters", "error");
      return;
    }
    dataPwdBtn.disabled = true;
    dataPwdBtn.textContent = "Deriving key…";
    try {
      let salt = localStorage.getItem("lf_salt");
      if (!salt) {
        salt = generateSalt();
        localStorage.setItem("lf_salt", salt);
      }
      const key = await deriveKey(pwd, salt);
      setSessionKey(key);
      dataPwdInput.value = "";
      cryptoStatus.textContent = "✓ Encryption key active for this session";
      showToast("Encryption key set", "success");
    } catch (err) {
      showToast(`Key derivation failed: ${err.message}`, "error");
    } finally {
      dataPwdBtn.disabled = false;
      dataPwdBtn.textContent = "Set Encryption Key";
    }
  });

  // Show key status if already loaded
  if (isKeyLoaded()) {
    document.getElementById("settings-crypto-status").textContent =
      "✓ Encryption key active for this session";
  }
}

function updateOneDriveUI(email) {
  const info = document.getElementById("settings-onedrive-info");
  const banner = document.getElementById("onedrive-status");
  const connectBtn = document.getElementById("settings-onedrive-btn");
  const disconnectBtn = document.getElementById("settings-onedrive-disconnect");

  if (state.oneDriveConnected) {
    if (info) info.textContent = `Connected${email ? `: ${email}` : ""} · /LedgerFlow/`;
    if (banner) banner.textContent = `Connected${email ? ` as ${email}` : ""}`;
    if (connectBtn) connectBtn.style.display = "none";
    if (disconnectBtn) disconnectBtn.style.display = "inline-block";
    // Hide the top banner once connected
    document.getElementById("onedrive-panel").style.display = "none";
  } else {
    if (info) info.textContent = "Not connected";
    if (banner) banner.textContent = "";
    if (connectBtn) connectBtn.style.display = "inline-block";
    if (disconnectBtn) disconnectBtn.style.display = "none";
  }
}

// ── Reconcile setup ───────────────────────────────────────────────────────────
function setupReconcile() {
  document.getElementById("recon-match-btn").addEventListener("click", async () => {
    const bankRange = document.getElementById("recon-bank-range").value.trim();
    const ledgerRange = document.getElementById("recon-ledger-range").value.trim();
    const tolerance = parseInt(document.getElementById("recon-tolerance").value || "3", 10);

    if (!bankRange || !ledgerRange) {
      showResult("recon-result", "Please enter both Bank and Ledger ranges.", "error");
      return;
    }

    try {
      await Excel.run(async (ctx) => {
        const bankSheet = ctx.workbook.worksheets.getActiveWorksheet();
        const bankData = bankSheet.getRange(bankRange);
        const ledgerData = bankSheet.getRange(ledgerRange);

        bankData.load("values");
        ledgerData.load("values");
        await ctx.sync();

        const bankRows = bankData.values.filter((r) => r[0]);  // Col A = Date
        const ledgerRows = ledgerData.values.filter((r) => r[0]);

        let matched = 0, unmatched = 0;
        const MS_PER_DAY = 86400000;
        const toleranceMs = tolerance * MS_PER_DAY;

        // Simple amount + date tolerance matching
        const ledgerUsed = new Set();

        bankRows.forEach((bRow, bi) => {
          const bDate = new Date(bRow[0]);
          const bAmt = parseFloat(bRow[2]) || 0;  // Col C = Amount

          let matchIdx = -1;
          ledgerRows.forEach((lRow, li) => {
            if (ledgerUsed.has(li)) return;
            const lDate = new Date(lRow[0]);
            const lAmt = parseFloat(lRow[2]) || 0;
            const dateDiff = Math.abs(bDate - lDate);
            if (Math.abs(bAmt - lAmt) < 0.01 && dateDiff <= toleranceMs) {
              matchIdx = li;
            }
          });

          const bankCell = bankSheet.getRange(bankRange).getCell(bi, 0).getEntireRow();
          if (matchIdx >= 0) {
            bankCell.format.fill.color = "#C6EFCE";
            ledgerUsed.add(matchIdx);
            matched++;
          } else {
            bankCell.format.fill.color = "#FFC7CE";
            unmatched++;
          }
        });

        await ctx.sync();
        showResult(
          "recon-result",
          `Matched: ${matched} · Unmatched: ${unmatched} · Total bank rows: ${bankRows.length}`,
          "success"
        );
      });
    } catch (err) {
      showResult("recon-result", err.message, "error");
    }
  });

  document.getElementById("recon-clear-btn").addEventListener("click", async () => {
    try {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const used = sheet.getUsedRange();
        used.format.fill.clear();
        await ctx.sync();
        showResult("recon-result", "Formatting cleared.", "info");
      });
    } catch (err) {
      showResult("recon-result", err.message, "error");
    }
  });
}

// ── Closing setup ─────────────────────────────────────────────────────────────
function setupClosing() {
  // Journal validator
  document.getElementById("close-validate-btn").addEventListener("click", async () => {
    const debitCol = document.getElementById("close-debit-col").value.trim().toUpperCase() || "C";
    const creditCol = document.getElementById("close-credit-col").value.trim().toUpperCase() || "D";

    try {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const used = sheet.getUsedRange();
        used.load("values, rowCount");
        await ctx.sync();

        const colIndex = (letter) => letter.charCodeAt(0) - 65;
        const di = colIndex(debitCol);
        const ci = colIndex(creditCol);

        let totalDebit = 0, totalCredit = 0;
        used.values.forEach((row) => {
          totalDebit += parseFloat(row[di]) || 0;
          totalCredit += parseFloat(row[ci]) || 0;
        });

        const diff = Math.abs(totalDebit - totalCredit);
        if (diff < 0.01) {
          showResult("close-result", `✓ Balanced — Debits: ${totalDebit.toFixed(2)} = Credits: ${totalCredit.toFixed(2)}`, "success");
        } else {
          showResult("close-result", `✗ Out of balance by ${diff.toFixed(2)} — Debits: ${totalDebit.toFixed(2)}, Credits: ${totalCredit.toFixed(2)}`, "error");
        }
        await ctx.sync();
      });
    } catch (err) {
      showResult("close-result", err.message, "error");
    }
  });

  // Accruals scheduler
  document.getElementById("accrual-generate-btn").addEventListener("click", async () => {
    const desc = document.getElementById("accrual-desc").value.trim();
    const amount = parseFloat(document.getElementById("accrual-amount").value) || 0;
    const start = document.getElementById("accrual-start").value;
    const end = document.getElementById("accrual-end").value;
    const glCode = document.getElementById("accrual-gl").value.trim();

    if (!desc || !amount || !start || !end) {
      showResult("accrual-result", "Please fill in all fields.", "error");
      return;
    }

    try {
      await Excel.run(async (ctx) => {
        // Build list of months between start and end
        const startDate = new Date(`${start}-01`);
        const endDate = new Date(`${end}-01`);
        const rows = [["Period", "GL Code", "Description", "Debit", "Credit", "Type"]];

        const d = new Date(startDate);
        while (d <= endDate) {
          const period = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
          rows.push([period, glCode, desc, amount, "", "Accrual"]);
          rows.push([period, glCode, `Reversal: ${desc}`, "", amount, "Reversal"]);
          d.setMonth(d.getMonth() + 1);
        }

        // Write to new sheet
        let sheet;
        try {
          sheet = ctx.workbook.worksheets.getItem("Accruals_Schedule");
          sheet.delete();
          await ctx.sync();
        } catch { /* sheet doesn't exist yet */ }

        sheet = ctx.workbook.worksheets.add("Accruals_Schedule");
        const range = sheet.getRange(`A1:F${rows.length}`);
        range.values = rows;
        range.getRow(0).format.font.bold = true;
        range.getRow(0).format.fill.color = "#1a6b3c";
        range.getRow(0).format.font.color = "#ffffff";
        sheet.getUsedRange().format.autofitColumns();
        sheet.activate();

        await ctx.sync();
        showResult("accrual-result", `Accruals schedule generated (${rows.length - 1} entries).`, "success");
      });
    } catch (err) {
      showResult("accrual-result", err.message, "error");
    }
  });
}

// ── VAT setup ─────────────────────────────────────────────────────────────────
function setupVAT() {
  document.getElementById("vat-tag-btn").addEventListener("click", async () => {
    const amtCol = document.getElementById("vat-amount-col").value.trim().toUpperCase() || "C";
    const tagCol = document.getElementById("vat-tag-col").value.trim().toUpperCase() || "E";
    const outCol = document.getElementById("vat-out-col").value.trim().toUpperCase() || "F";
    const vatType = document.getElementById("vat-type").value;
    const vatRate = parseFloat(document.getElementById("vat-rate").value || "15") / 100;

    const rateMap = { standard: vatRate, zero: 0, exempt: 0, oos: 0 };
    const rate = rateMap[vatType];

    try {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const selection = ctx.workbook.getSelectedRange();
        selection.load("rowIndex, rowCount");
        await ctx.sync();

        const startRow = selection.rowIndex;
        const rowCount = selection.rowCount;

        for (let i = 0; i < rowCount; i++) {
          const row = startRow + i + 1; // 1-based for getRange
          const amtCell = sheet.getRange(`${amtCol}${row}`);
          amtCell.load("values");
          await ctx.sync();

          const amt = parseFloat(amtCell.values[0][0]) || 0;
          sheet.getRange(`${tagCol}${row}`).values = [[vatType.toUpperCase()]];
          sheet.getRange(`${outCol}${row}`).values = [[amt * rate]];
        }

        await ctx.sync();
        showResult("vat-result", `Applied ${vatType} tag to ${rowCount} rows.`, "success");
      });
    } catch (err) {
      showResult("vat-result", err.message, "error");
    }
  });

  document.getElementById("vat-summary-btn").addEventListener("click", async () => {
    const tagCol = document.getElementById("vat-tag-col").value.trim().toUpperCase() || "E";
    const outCol = document.getElementById("vat-out-col").value.trim().toUpperCase() || "F";
    const amtCol = document.getElementById("vat-amount-col").value.trim().toUpperCase() || "C";

    try {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const used = sheet.getUsedRange();
        used.load("values, rowCount");
        await ctx.sync();

        const colIdx = (l) => l.charCodeAt(0) - 65;
        const ti = colIdx(tagCol), oi = colIdx(outCol), ai = colIdx(amtCol);

        let outputVAT = 0, inputVAT = 0, zeroRated = 0, exempt = 0;
        used.values.forEach((row) => {
          const tag = (row[ti] || "").toString().toUpperCase();
          const net = parseFloat(row[ai]) || 0;
          const vat = parseFloat(row[oi]) || 0;
          if (tag === "STANDARD") outputVAT += vat;
          if (tag === "ZERO") zeroRated += net;
          if (tag === "EXEMPT") exempt += net;
        });

        const summary = [
          ["VAT Return Working", "", ""],
          ["Output VAT (Standard-rated)", "", outputVAT],
          ["Less: Input VAT",             "", inputVAT],
          ["Net VAT Payable",             "", outputVAT - inputVAT],
          ["Zero-rated turnover",         "", zeroRated],
          ["Exempt supplies",             "", exempt],
        ];

        let wsheet;
        try {
          wsheet = ctx.workbook.worksheets.getItem("VAT_Return_Working");
          wsheet.delete();
          await ctx.sync();
        } catch { /* */ }
        wsheet = ctx.workbook.worksheets.add("VAT_Return_Working");
        const range = wsheet.getRange(`A1:C${summary.length}`);
        range.values = summary;
        range.getRow(0).format.font.bold = true;
        range.getRow(0).format.fill.color = "#1a6b3c";
        range.getRow(0).format.font.color = "#ffffff";
        wsheet.getUsedRange().format.autofitColumns();
        wsheet.activate();
        await ctx.sync();
        showResult("vat-result", "VAT Return Working sheet created.", "success");
      });
    } catch (err) {
      showResult("vat-result", err.message, "error");
    }
  });
}

// ── Payroll setup ─────────────────────────────────────────────────────────────

// Default PAYE bands (Ghana-style — user can override and save)
let payrollBands = [
  { from: 0,      to: 4380,   rate: 0   },
  { from: 4380,   to: 5100,   rate: 5   },
  { from: 5100,   to: 6240,   rate: 10  },
  { from: 6240,   to: 18600,  rate: 17.5},
  { from: 18600,  to: 50000,  rate: 25  },
  { from: 50000,  to: 0,      rate: 30  },
];

function renderPayBands() {
  const container = document.getElementById("pay-bands-rows");
  container.innerHTML = "";
  payrollBands.forEach((band, i) => {
    const row = document.createElement("div");
    row.style.cssText = "display:grid;grid-template-columns:1fr 1fr 1fr 24px;gap:4px;margin-bottom:3px";
    row.innerHTML = `
      <input class="lf-input" style="margin:0;font-size:11px" data-i="${i}" data-f="from" type="number" value="${band.from}"/>
      <input class="lf-input" style="margin:0;font-size:11px" data-i="${i}" data-f="to"   type="number" value="${band.to}"/>
      <input class="lf-input" style="margin:0;font-size:11px" data-i="${i}" data-f="rate" type="number" value="${band.rate}" step="0.5"/>
      <button data-i="${i}" class="pay-del-band" style="background:#dc3545;color:#fff;border:none;border-radius:3px;cursor:pointer;font-size:11px">✕</button>
    `;
    container.appendChild(row);
  });
  container.querySelectorAll(".pay-del-band").forEach((btn) =>
    btn.addEventListener("click", () => {
      payrollBands.splice(parseInt(btn.dataset.i), 1);
      renderPayBands();
    })
  );
}

function calcPAYE(annualIncome) {
  let tax = 0;
  let remaining = annualIncome;
  for (const band of payrollBands) {
    if (remaining <= 0) break;
    const bandWidth = band.to > 0 ? band.to - band.from : Infinity;
    const taxable = Math.min(remaining, bandWidth);
    tax += (taxable * band.rate) / 100;
    remaining -= taxable;
  }
  return tax / 12; // monthly tax
}

function setupPayroll() {
  renderPayBands();

  // Load saved bands from OneDrive/localStorage
  API.getPayrollConfig().then((cfg) => {
    if (cfg && cfg.bands) {
      payrollBands = cfg.bands;
      renderPayBands();
    }
  }).catch(() => {});

  document.getElementById("pay-add-band-btn").addEventListener("click", () => {
    payrollBands.push({ from: 0, to: 0, rate: 0 });
    renderPayBands();
  });

  document.getElementById("pay-save-bands-btn").addEventListener("click", async () => {
    // Read current values from inputs
    document.querySelectorAll("#pay-bands-rows [data-i]").forEach((inp) => {
      const i = parseInt(inp.dataset.i);
      payrollBands[i][inp.dataset.f] = parseFloat(inp.value) || 0;
    });
    try {
      await API.savePayrollConfig({ bands: payrollBands });
      showResult("pay-bands-result", "Tax bands saved.", "success");
    } catch (err) {
      showResult("pay-bands-result", err.message, "error");
    }
  });

  document.getElementById("pay-calc-btn").addEventListener("click", async () => {
    const nameCol   = document.getElementById("pay-name-col").value.trim().toUpperCase()  || "A";
    const grossCol  = document.getElementById("pay-gross-col").value.trim().toUpperCase() || "B";
    const allowCol  = document.getElementById("pay-allow-col").value.trim().toUpperCase() || "C";
    const netCol    = document.getElementById("pay-net-col").value.trim().toUpperCase()   || "F";
    const ssRate    = parseFloat(document.getElementById("pay-ss-rate").value  || "5.5")  / 100;

    try {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const used  = sheet.getUsedRange();
        used.load("values, rowCount");
        await ctx.sync();

        const colIdx = (l) => l.charCodeAt(0) - 65;
        const ni = colIdx(nameCol), gi = colIdx(grossCol), ai = colIdx(allowCol);
        let processed = 0;

        used.values.forEach((row, ri) => {
          const gross    = parseFloat(row[gi]) || 0;
          if (!gross || ri === 0) return; // skip header / empty
          const allowances = parseFloat(row[ai]) || 0;
          const taxable    = gross + allowances;
          const paye       = calcPAYE(taxable * 12);
          const ss         = gross * ssRate;
          const net        = gross + allowances - paye - ss;

          const outputRow  = ri + 1; // 1-based
          sheet.getRange(`D${outputRow}`).values = [[paye]];    // PAYE col D
          sheet.getRange(`E${outputRow}`).values = [[ss]];      // SS col E
          sheet.getRange(`${netCol}${outputRow}`).values = [[net]];
          processed++;
        });

        // Write column headers on row 1
        sheet.getRange("D1").values = [["PAYE"]];
        sheet.getRange("E1").values = [["SS Contribution"]];
        sheet.getRange(`${netCol}1`).values = [["Net Pay"]];

        await ctx.sync();
        showResult("pay-result", `Payroll calculated for ${processed} employee(s).`, "success");
      });
    } catch (err) {
      showResult("pay-result", err.message, "error");
    }
  });

  document.getElementById("pay-payslip-btn").addEventListener("click", async () => {
    const nameCol  = document.getElementById("pay-name-col").value.trim().toUpperCase()  || "A";
    const grossCol = document.getElementById("pay-gross-col").value.trim().toUpperCase() || "B";
    const allowCol = document.getElementById("pay-allow-col").value.trim().toUpperCase() || "C";
    const ssRate   = parseFloat(document.getElementById("pay-ss-rate").value  || "5.5") / 100;

    try {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const used  = sheet.getUsedRange();
        used.load("values, rowCount");
        await ctx.sync();

        const colIdx   = (l) => l.charCodeAt(0) - 65;
        const ni = colIdx(nameCol), gi = colIdx(grossCol), ai = colIdx(allowCol);
        const period   = new Date().toLocaleDateString("en-GB", { month: "long", year: "numeric" });

        let slipSheet;
        try {
          slipSheet = ctx.workbook.worksheets.getItem("Payslips");
          slipSheet.delete();
          await ctx.sync();
        } catch { /* */ }
        slipSheet = ctx.workbook.worksheets.add("Payslips");

        let startRow = 1;
        const employees = used.values.filter((r, i) => i > 0 && (parseFloat(r[gi]) || 0) > 0);

        employees.forEach((row) => {
          const name       = row[ni] || "Employee";
          const gross      = parseFloat(row[gi]) || 0;
          const allowances = parseFloat(row[ai]) || 0;
          const paye       = calcPAYE((gross + allowances) * 12);
          const ss         = gross * ssRate;
          const net        = gross + allowances - paye - ss;

          const slipData = [
            [`PAYSLIP — ${period}`, "", ""],
            ["Employee", name, ""],
            ["", "", ""],
            ["Gross Salary", "", gross],
            ["Allowances", "", allowances],
            ["", "", ""],
            ["PAYE Tax", "", -paye],
            ["Social Security", "", -ss],
            ["", "", ""],
            ["NET PAY", "", net],
            ["", "", ""],
            ["", "", ""],
          ];

          const r = slipSheet.getRange(`A${startRow}:C${startRow + slipData.length - 1}`);
          r.values = slipData;
          slipSheet.getRange(`A${startRow}`).format.font.bold = true;
          slipSheet.getRange(`A${startRow + 9}`).format.font.bold = true;
          slipSheet.getRange(`A${startRow}:C${startRow}`).format.fill.color = "#1a6b3c";
          slipSheet.getRange(`A${startRow}:C${startRow}`).format.font.color = "#ffffff";
          startRow += slipData.length;
        });

        slipSheet.getUsedRange().format.autofitColumns();
        slipSheet.activate();
        await ctx.sync();
        showResult("pay-result", `${employees.length} payslip(s) generated on sheet 'Payslips'.`, "success");
      });
    } catch (err) {
      showResult("pay-result", err.message, "error");
    }
  });
}

// ── Fixed Assets setup ────────────────────────────────────────────────────────

function setupAssets() {
  // Add asset to register
  document.getElementById("ast-add-btn").addEventListener("click", async () => {
    const name     = document.getElementById("ast-name").value.trim();
    const cost     = parseFloat(document.getElementById("ast-cost").value) || 0;
    const residual = parseFloat(document.getElementById("ast-residual").value) || 0;
    const life     = parseInt(document.getElementById("ast-life").value)   || 1;
    const date     = document.getElementById("ast-date").value;
    const method   = document.getElementById("ast-method").value;
    const gl       = document.getElementById("ast-gl").value.trim();

    if (!name || !cost || !date) {
      showResult("ast-result", "Name, cost and date are required.", "error");
      return;
    }
    try {
      const assets = (await API.getAssets()) || [];
      assets.push({ id: Date.now(), name, cost, residual, life, date, method, gl, active: true });
      await API.saveAssets(assets);
      showResult("ast-result", `"${name}" added to register.`, "success");
      ["ast-name","ast-cost","ast-residual","ast-life","ast-date","ast-gl"].forEach(
        (id) => (document.getElementById(id).value = "")
      );
    } catch (err) {
      showResult("ast-result", err.message, "error");
    }
  });

  // Generate depreciation schedule to Excel
  document.getElementById("ast-schedule-btn").addEventListener("click", async () => {
    const name     = document.getElementById("ast-name").value.trim();
    const cost     = parseFloat(document.getElementById("ast-cost").value) || 0;
    const residual = parseFloat(document.getElementById("ast-residual").value) || 0;
    const life     = parseInt(document.getElementById("ast-life").value) || 1;
    const method   = document.getElementById("ast-method").value;
    const assetLabel = name || "Asset";

    if (!cost) { showResult("ast-result", "Enter asset cost to generate schedule.", "error"); return; }

    try {
      await Excel.run(async (ctx) => {
        const rows = [["Year", "Opening NBV", "Depreciation", "Closing NBV", "Accumulated Dep"]];
        let nbv = cost;
        let accumulated = 0;
        const annualDep = method === "sl"
          ? (cost - residual) / life
          : null; // null = reducing balance % calculated per year

        for (let y = 1; y <= life; y++) {
          const dep = method === "sl"
            ? annualDep
            : nbv * (1 - Math.pow(residual / cost, 1 / life));
          const effectiveDep = Math.min(dep, nbv - residual);
          accumulated += effectiveDep;
          rows.push([y, nbv, effectiveDep, nbv - effectiveDep, accumulated]);
          nbv -= effectiveDep;
          if (nbv <= residual + 0.01) break;
        }

        let sh;
        try { sh = ctx.workbook.worksheets.getItem("Depreciation_Schedule"); sh.delete(); await ctx.sync(); } catch { /* */ }
        sh = ctx.workbook.worksheets.add("Depreciation_Schedule");
        sh.getRange("A1").values = [[`Depreciation Schedule: ${assetLabel} (${method === "sl" ? "Straight-Line" : "Reducing Balance"})`]];
        sh.getRange("A1").format.font.bold = true;
        const dataRange = sh.getRange(`A3:E${rows.length + 2}`);
        dataRange.values = rows;
        sh.getRange("A3:E3").format.font.bold = true;
        sh.getRange("A3:E3").format.fill.color = "#1a6b3c";
        sh.getRange("A3:E3").format.font.color = "#ffffff";
        sh.getUsedRange().format.autofitColumns();
        sh.activate();
        await ctx.sync();
        showResult("ast-result", `Depreciation schedule generated (${rows.length - 1} years).`, "success");
      });
    } catch (err) {
      showResult("ast-result", err.message, "error");
    }
  });

  // View register in Excel sheet
  document.getElementById("ast-view-btn").addEventListener("click", async () => {
    try {
      const assets = (await API.getAssets()) || [];
      if (!assets.length) { showResult("ast-result", "No assets in register.", "info"); return; }

      await Excel.run(async (ctx) => {
        let sh;
        try { sh = ctx.workbook.worksheets.getItem("Asset_Register"); sh.delete(); await ctx.sync(); } catch { /* */ }
        sh = ctx.workbook.worksheets.add("Asset_Register");

        const header = [["ID", "Name", "Cost", "Residual", "Life (yrs)", "Date", "Method", "GL Code", "Status"]];
        const rows   = assets.map((a) => [a.id, a.name, a.cost, a.residual, a.life, a.date, a.method, a.gl, a.active ? "Active" : "Disposed"]);
        const all    = [...header, ...rows];

        const range = sh.getRange(`A1:I${all.length}`);
        range.values = all;
        sh.getRange("A1:I1").format.font.bold  = true;
        sh.getRange("A1:I1").format.fill.color = "#1a6b3c";
        sh.getRange("A1:I1").format.font.color = "#ffffff";
        sh.getUsedRange().format.autofitColumns();
        sh.activate();
        await ctx.sync();
        showResult("ast-result", `${assets.length} asset(s) displayed on sheet 'Asset_Register'.`, "success");
      });
    } catch (err) {
      showResult("ast-result", err.message, "error");
    }
  });

  // Disposal calculator
  document.getElementById("disp-calc-btn").addEventListener("click", async () => {
    const assetName = document.getElementById("disp-name").value.trim();
    const nbv       = parseFloat(document.getElementById("disp-nbv").value)      || 0;
    const proceeds  = parseFloat(document.getElementById("disp-proceeds").value) || 0;

    if (!nbv) { showResult("disp-result", "Enter Net Book Value.", "error"); return; }

    const gainLoss = proceeds - nbv;
    const type     = gainLoss >= 0 ? "Gain on disposal" : "Loss on disposal";

    try {
      await Excel.run(async (ctx) => {
        const sh    = ctx.workbook.worksheets.getActiveWorksheet();
        const used  = sh.getUsedRange();
        used.load("rowCount");
        await ctx.sync();

        const nextRow = used.rowCount + 2;
        const entry   = [
          [`Disposal: ${assetName || "Asset"}`, "", "", ""],
          ["Account", "Description",        "Debit",  "Credit"],
          ["Cash/Bank",   "Disposal proceeds",  proceeds, ""],
          ["Accum. Dep.", "Remove accumulated dep.", nbv,  ""],
          ["Asset Cost",  "Remove asset cost",  "",       nbv + (proceeds > nbv ? gainLoss : 0)],
          [gainLoss >= 0 ? "Gain on Disposal" : "Loss on Disposal",
           type, gainLoss < 0 ? Math.abs(gainLoss) : "", gainLoss >= 0 ? gainLoss : ""],
        ];

        sh.getRange(`A${nextRow}:D${nextRow + entry.length - 1}`).values = entry;
        sh.getRange(`A${nextRow}`).format.font.bold = true;
        await ctx.sync();
        showResult("disp-result", `${type}: ${Math.abs(gainLoss).toFixed(2)}. Journal written to active sheet.`, gainLoss >= 0 ? "success" : "info");
      });
    } catch (err) {
      showResult("disp-result", err.message, "error");
    }
  });
}

// ── Financial Reports setup ───────────────────────────────────────────────────

let tbMapping = [];

function renderMappingRows() {
  const container = document.getElementById("rpt-mapping-rows");
  container.innerHTML = "";
  tbMapping.forEach((m, i) => {
    const row = document.createElement("div");
    row.style.cssText = "display:grid;grid-template-columns:60px 60px 1fr 24px;gap:4px;margin-bottom:3px";
    const fsLines = [
      "Revenue","Cost of Sales","Gross Profit",
      "Operating Expenses","Operating Profit","Finance Costs",
      "Profit Before Tax","Tax","Profit After Tax",
      "Non-current Assets","Current Assets","Total Assets",
      "Non-current Liabilities","Current Liabilities","Total Liabilities","Equity",
    ];
    row.innerHTML = `
      <input class="lf-input" style="margin:0;font-size:11px" data-i="${i}" data-f="from" type="number" value="${m.from}"/>
      <input class="lf-input" style="margin:0;font-size:11px" data-i="${i}" data-f="to"   type="number" value="${m.to}"/>
      <select class="lf-input" style="margin:0;font-size:11px" data-i="${i}" data-f="line">
        ${fsLines.map((l) => `<option${l === m.line ? " selected" : ""}>${l}</option>`).join("")}
      </select>
      <button data-i="${i}" class="rpt-del-row" style="background:#dc3545;color:#fff;border:none;border-radius:3px;cursor:pointer;font-size:11px">✕</button>
    `;
    container.appendChild(row);
  });
  container.querySelectorAll(".rpt-del-row").forEach((btn) =>
    btn.addEventListener("click", () => { tbMapping.splice(parseInt(btn.dataset.i), 1); renderMappingRows(); })
  );
}

function setupReports() {
  // Load saved mapping
  API.getTBMapping().then((m) => {
    if (m && m.mapping) { tbMapping = m.mapping; renderMappingRows(); }
    else {
      // Bootstrap with a sensible default
      tbMapping = [
        { from: 4000, to: 4999, line: "Revenue" },
        { from: 5000, to: 5999, line: "Cost of Sales" },
        { from: 6000, to: 6999, line: "Operating Expenses" },
        { from: 7000, to: 7999, line: "Finance Costs" },
        { from: 1000, to: 1499, line: "Non-current Assets" },
        { from: 1500, to: 1999, line: "Current Assets" },
        { from: 2000, to: 2499, line: "Non-current Liabilities" },
        { from: 2500, to: 2999, line: "Current Liabilities" },
        { from: 3000, to: 3999, line: "Equity" },
      ];
      renderMappingRows();
    }
  }).catch(() => {});

  document.getElementById("rpt-add-mapping-btn").addEventListener("click", () => {
    tbMapping.push({ from: 0, to: 0, line: "Revenue" });
    renderMappingRows();
  });

  document.getElementById("rpt-save-mapping-btn").addEventListener("click", async () => {
    document.querySelectorAll("#rpt-mapping-rows [data-i]").forEach((inp) => {
      const i = parseInt(inp.dataset.i);
      const f = inp.dataset.f;
      tbMapping[i][f] = (f === "line") ? inp.value : (parseFloat(inp.value) || 0);
    });
    try {
      await API.saveTBMapping({ mapping: tbMapping });
      showResult("rpt-mapping-result", "Mapping saved.", "success");
    } catch (err) {
      showResult("rpt-mapping-result", err.message, "error");
    }
  });

  // Helper: read TB and aggregate by FS line
  async function readAndAggregateTB(ctx) {
    const codeCol   = document.getElementById("tb-col-code").value.trim().toUpperCase()   || "A";
    const debitCol  = document.getElementById("tb-col-debit").value.trim().toUpperCase()  || "C";
    const creditCol = document.getElementById("tb-col-credit").value.trim().toUpperCase() || "D";

    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const used  = sheet.getUsedRange();
    used.load("values");
    await ctx.sync();

    const colIdx = (l) => l.charCodeAt(0) - 65;
    const ci = colIdx(codeCol), di = colIdx(debitCol), cri = colIdx(creditCol);

    const aggregated = {};
    used.values.forEach((row, ri) => {
      if (ri === 0) return; // skip header
      const code   = parseFloat(row[ci]);
      if (!code) return;
      const net    = (parseFloat(row[di]) || 0) - (parseFloat(row[cri]) || 0);
      const match  = tbMapping.find((m) => code >= m.from && (m.to === 0 || code <= m.to));
      const line   = match ? match.line : "Unclassified";
      aggregated[line] = (aggregated[line] || 0) + net;
    });
    return aggregated;
  }

  async function writeStyledSheet(ctx, sheetName, rows, titleColor = "#1a6b3c") {
    let sh;
    try { sh = ctx.workbook.worksheets.getItem(sheetName); sh.delete(); await ctx.sync(); } catch { /* */ }
    sh = ctx.workbook.worksheets.add(sheetName);
    const range = sh.getRange(`A1:C${rows.length}`);
    range.values = rows;
    rows.forEach((row, i) => {
      if (row[2] === "" && row[0] !== "") {
        // Section header row
        sh.getRange(`A${i + 1}:C${i + 1}`).format.font.bold = true;
        sh.getRange(`A${i + 1}:C${i + 1}`).format.fill.color = titleColor;
        sh.getRange(`A${i + 1}:C${i + 1}`).format.font.color = "#ffffff";
      }
    });
    sh.getUsedRange().format.autofitColumns();
    sh.activate();
    return sh;
  }

  // P&L
  document.getElementById("rpt-pl-btn").addEventListener("click", async () => {
    try {
      await Excel.run(async (ctx) => {
        const agg = await readAndAggregateTB(ctx);

        const revenue     = agg["Revenue"]              || 0;
        const cos         = agg["Cost of Sales"]        || 0;
        const opex        = agg["Operating Expenses"]   || 0;
        const finance     = agg["Finance Costs"]        || 0;
        const grossProfit = revenue - cos;
        const opProfit    = grossProfit - opex;
        const pbt         = opProfit - finance;

        const rows = [
          ["Income Statement", "", ""],
          ["Revenue", "", revenue],
          ["Cost of Sales", "", -Math.abs(cos)],
          ["GROSS PROFIT", "", grossProfit],
          ["", "", ""],
          ["Operating Expenses", "", -Math.abs(opex)],
          ["OPERATING PROFIT", "", opProfit],
          ["", "", ""],
          ["Finance Costs", "", -Math.abs(finance)],
          ["PROFIT BEFORE TAX", "", pbt],
        ];
        await writeStyledSheet(ctx, "Income_Statement", rows);
        await ctx.sync();
        showResult("rpt-result", "Income Statement generated.", "success");
      });
    } catch (err) { showResult("rpt-result", err.message, "error"); }
  });

  // Balance Sheet
  document.getElementById("rpt-bs-btn").addEventListener("click", async () => {
    try {
      await Excel.run(async (ctx) => {
        const agg = await readAndAggregateTB(ctx);

        const nca   = agg["Non-current Assets"]      || 0;
        const ca    = agg["Current Assets"]           || 0;
        const ncl   = agg["Non-current Liabilities"]  || 0;
        const cl    = agg["Current Liabilities"]      || 0;
        const eq    = agg["Equity"]                   || 0;
        const totalAssets = nca + ca;
        const totalLiab   = ncl + cl;
        const check       = totalAssets - (totalLiab + eq);

        const rows = [
          ["Balance Sheet", "", ""],
          ["NON-CURRENT ASSETS", "", ""],
          ["Non-current Assets", "", nca],
          ["", "", ""],
          ["CURRENT ASSETS", "", ""],
          ["Current Assets", "", ca],
          ["TOTAL ASSETS", "", totalAssets],
          ["", "", ""],
          ["NON-CURRENT LIABILITIES", "", ""],
          ["Non-current Liabilities", "", ncl],
          ["", "", ""],
          ["CURRENT LIABILITIES", "", ""],
          ["Current Liabilities", "", cl],
          ["TOTAL LIABILITIES", "", totalLiab],
          ["", "", ""],
          ["Equity", "", eq],
          ["TOTAL LIABILITIES + EQUITY", "", totalLiab + eq],
          ["", "", ""],
          ["Check (Assets - L&E)", "", check],
        ];
        await writeStyledSheet(ctx, "Balance_Sheet", rows);
        await ctx.sync();
        showResult("rpt-result",
          `Balance Sheet generated. ${Math.abs(check) < 1 ? "✓ Balanced." : `⚠ Out by ${check.toFixed(2)}.`}`,
          Math.abs(check) < 1 ? "success" : "error"
        );
      });
    } catch (err) { showResult("rpt-result", err.message, "error"); }
  });

  // Cash Flow (Indirect Method)
  document.getElementById("rpt-cf-btn").addEventListener("click", async () => {
    try {
      await Excel.run(async (ctx) => {
        const agg    = await readAndAggregateTB(ctx);
        const pbt    = (agg["Revenue"] || 0) - (agg["Cost of Sales"] || 0) - (agg["Operating Expenses"] || 0) - (agg["Finance Costs"] || 0);
        const nca    = agg["Non-current Assets"]   || 0;
        const ca     = agg["Current Assets"]        || 0;
        const cl     = agg["Current Liabilities"]   || 0;
        const wcChange = ca - cl; // simplified working capital movement

        const rows = [
          ["Cash Flow Statement (Indirect)", "", ""],
          ["OPERATING ACTIVITIES", "", ""],
          ["Profit Before Tax", "", pbt],
          ["Adjustments for non-cash items", "", ""],
          ["Depreciation (add back)", "", 0],
          ["Working Capital Changes", "", -wcChange],
          ["Cash from Operations", "", pbt - wcChange],
          ["", "", ""],
          ["INVESTING ACTIVITIES", "", ""],
          ["Purchase of Fixed Assets", "", -nca],
          ["Cash used in Investing", "", -nca],
          ["", "", ""],
          ["FINANCING ACTIVITIES", "", ""],
          ["Net Financing", "", 0],
          ["", "", ""],
          ["NET CHANGE IN CASH", "", pbt - wcChange - nca],
        ];
        await writeStyledSheet(ctx, "Cash_Flow", rows, "#0b3d91");
        await ctx.sync();
        showResult("rpt-result", "Cash Flow Statement generated.", "success");
      });
    } catch (err) { showResult("rpt-result", err.message, "error"); }
  });
}

// ── Budget Variance setup ─────────────────────────────────────────────────────

function setupBudget() {
  document.getElementById("bud-variance-btn").addEventListener("click", async () => {
    const actualSheet  = document.getElementById("bud-actual-sheet").value.trim()  || "Actual";
    const budgetSheet  = document.getElementById("bud-budget-sheet").value.trim()  || "Budget";
    const amtCol       = document.getElementById("bud-amt-col").value.trim().toUpperCase() || "B";
    const threshold    = parseFloat(document.getElementById("bud-threshold").value || "10") / 100;
    const outSheetName = document.getElementById("bud-out-sheet").value.trim()     || "Variance_Analysis";

    try {
      await Excel.run(async (ctx) => {
        const wb = ctx.workbook;
        const actual = wb.worksheets.getItem(actualSheet);
        const budget = wb.worksheets.getItem(budgetSheet);

        const actUsed = actual.getUsedRange();
        const budUsed = budget.getUsedRange();
        actUsed.load("values");
        budUsed.load("values");
        await ctx.sync();

        const colIdx = (l) => l.charCodeAt(0) - 65;
        const ai = colIdx(amtCol);

        const rows = [["Description", "Actual", "Budget", "Variance", "Var %", "Status"]];

        actUsed.values.forEach((row, ri) => {
          const desc   = row[0];
          const actual = parseFloat(row[ai]) || 0;
          const budget = parseFloat((budUsed.values[ri] || [])[ai]) || 0;
          if (!desc) return;
          if (ri === 0) return; // skip header row

          const variance  = actual - budget;
          const varPct    = budget !== 0 ? (variance / Math.abs(budget)) * 100 : 0;
          const status    = Math.abs(varPct) <= threshold * 100
            ? "On Track"
            : variance > 0 ? "Favourable" : "Adverse";
          rows.push([desc, actual, budget, variance, `${varPct.toFixed(1)}%`, status]);
        });

        let outSheet;
        try { outSheet = wb.worksheets.getItem(outSheetName); outSheet.delete(); await ctx.sync(); } catch { /* */ }
        outSheet = wb.worksheets.add(outSheetName);

        const range = outSheet.getRange(`A1:F${rows.length}`);
        range.values = rows;
        outSheet.getRange("A1:F1").format.font.bold  = true;
        outSheet.getRange("A1:F1").format.fill.color = "#1a6b3c";
        outSheet.getRange("A1:F1").format.font.color = "#ffffff";

        // Colour-code status column (F)
        rows.forEach((row, i) => {
          if (i === 0) return;
          const cell = outSheet.getRange(`F${i + 1}`);
          if (row[5] === "Favourable") cell.format.fill.color = "#C6EFCE";
          else if (row[5] === "Adverse") cell.format.fill.color = "#FFC7CE";
        });

        outSheet.getUsedRange().format.autofitColumns();
        outSheet.activate();
        await ctx.sync();
        showResult("bud-variance-result", `Variance analysis complete — ${rows.length - 1} line(s).`, "success");
      });
    } catch (err) {
      showResult("bud-variance-result", err.message, "error");
    }
  });

  document.getElementById("bud-forecast-btn").addEventListener("click", async () => {
    const budgetSheet  = document.getElementById("bud-budget-sheet").value.trim() || "Budget";
    const actualSheet  = document.getElementById("bud-actual-sheet").value.trim() || "Actual";
    const lockedMonths = parseInt(document.getElementById("bud-locked-months").value || "3");
    const method       = document.getElementById("bud-forecast-method").value;

    try {
      await Excel.run(async (ctx) => {
        const wb     = ctx.workbook;
        const budget = wb.worksheets.getItem(budgetSheet);
        const actual = wb.worksheets.getItem(actualSheet);

        const budUsed = budget.getUsedRange();
        const actUsed = actual.getUsedRange();
        budUsed.load("values, columnCount, rowCount");
        actUsed.load("values");
        await ctx.sync();

        const budVals = budUsed.values.map((r) => [...r]);
        const actVals = actUsed.values;

        // Replace forecast months with actuals (locked) or extrapolated trend
        budVals.forEach((row, ri) => {
          if (ri === 0) return; // header
          for (let ci = 1; ci <= budVals[ri].length; ci++) {
            if (ci <= lockedMonths) {
              // Use actual
              const actVal = (actVals[ri] || [])[ci];
              if (actVal !== undefined) budVals[ri][ci] = actVal;
            } else if (method === "trend" && lockedMonths >= 2) {
              // Simple linear extrapolation from last 2 actuals
              const prev1 = parseFloat((actVals[ri] || [])[lockedMonths])     || 0;
              const prev2 = parseFloat((actVals[ri] || [])[lockedMonths - 1]) || 0;
              const trend = prev1 - prev2;
              budVals[ri][ci] = prev1 + trend * (ci - lockedMonths);
            }
          }
        });

        let fcSheet;
        try { fcSheet = wb.worksheets.getItem("Forecast"); fcSheet.delete(); await ctx.sync(); } catch { /* */ }
        fcSheet = wb.worksheets.add("Forecast");

        const range = fcSheet.getRange(`A1:${String.fromCharCode(65 + budVals[0].length - 1)}${budVals.length}`);
        range.values = budVals;
        fcSheet.getRange("A1").format.font.bold = true;

        // Shade locked actual columns blue, forecast columns green
        for (let ci = 1; ci <= budVals[0].length - 1; ci++) {
          const col = String.fromCharCode(65 + ci);
          const colRange = fcSheet.getRange(`${col}1:${col}${budVals.length}`);
          colRange.format.fill.color = ci <= lockedMonths ? "#DDEBF7" : "#E2EFDA";
        }

        fcSheet.getUsedRange().format.autofitColumns();
        fcSheet.activate();
        await ctx.sync();
        showResult("bud-forecast-result",
          `Forecast generated. ${lockedMonths} month(s) actuals (blue) + ${budVals[0].length - 1 - lockedMonths} forecast months (green).`,
          "success"
        );
      });
    } catch (err) {
      showResult("bud-forecast-result", err.message, "error");
    }
  });
}
