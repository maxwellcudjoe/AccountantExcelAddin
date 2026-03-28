/**
 * ledgerflow.js — API service layer
 * - Cloud calls (auth, subscriptions): use request() with JWT Bearer token
 * - Data collections: use store.js (encrypted OneDrive / localStorage)
 */

import { saveData, loadData } from "../onedrive/store.js";

const API_URL = process.env.API_URL || "https://localhost:7264/api";
const TOKEN_KEY = "lf_token";

// ── HTTP helper ───────────────────────────────────────────────────────────────

export function getApiUrl() {
  return API_URL;
}

export async function request(method, path, body = null, authenticated = false) {
  const headers = { "Content-Type": "application/json" };

  if (authenticated) {
    const token = localStorage.getItem(TOKEN_KEY);
    if (token) headers["Authorization"] = `Bearer ${token}`;
  }

  const options = { method, headers };
  if (body) options.body = JSON.stringify(body);

  const res = await fetch(`${API_URL}${path}`, options);

  if (res.status === 401) {
    clearSession();
    throw new Error("Session expired. Please sign in again.");
  }

  if (!res.ok) {
    const err = await res.json().catch(() => ({ error: res.statusText }));
    throw new Error(err.error || "Request failed");
  }

  return res.json();
}

// ── Auth ─────────────────────────────────────────────────────────────────────

export async function login(email, password) {
  const { token } = await request("POST", "/auth/login", { email, password });
  localStorage.setItem(TOKEN_KEY, token);
  return token;
}

export async function register(email, password, fullName, firmName) {
  const { token } = await request("POST", "/auth/register", {
    email,
    password,
    fullName,
    firmName,
  });
  localStorage.setItem(TOKEN_KEY, token);
  return token;
}

export async function getMe() {
  return request("GET", "/auth/me", null, true);
}

export function getToken() {
  return localStorage.getItem(TOKEN_KEY);
}

export function clearSession() {
  localStorage.removeItem(TOKEN_KEY);
  localStorage.removeItem("lf_salt");
}

// ── Subscription ──────────────────────────────────────────────────────────────

export async function getPlans() {
  return request("GET", "/subscriptions/plans");
}

export async function getCurrentSubscription() {
  return request("GET", "/subscriptions/current", null, true);
}

export async function upgradePlan(planName) {
  return request("POST", "/subscriptions/upgrade", { planName }, true);
}

// ── Data collections (encrypted, stored in OneDrive / localStorage) ──────────

export const getTransactions  = () => loadData("transactions");
export const saveTransactions = (d) => saveData("transactions", d);

export const getVATRecords    = () => loadData("vat_records");
export const saveVATRecords   = (d) => saveData("vat_records", d);

export const getPayrollRuns   = () => loadData("payroll_runs");
export const savePayrollRuns  = (d) => saveData("payroll_runs", d);

export const getAssets        = () => loadData("assets");
export const saveAssets       = (d) => saveData("assets", d);

export const getBudgets       = () => loadData("budgets");
export const saveBudgets      = (d) => saveData("budgets", d);

export const getJournals      = () => loadData("journals");
export const saveJournals     = (d) => saveData("journals", d);

export const getTBMapping     = () => loadData("tb_mapping");
export const saveTBMapping    = (d) => saveData("tb_mapping", d);

export const getPayrollConfig = () => loadData("payroll_config");
export const savePayrollConfig = (d) => saveData("payroll_config", d);
