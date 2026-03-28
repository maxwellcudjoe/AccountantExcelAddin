/* global Office, Excel */

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
