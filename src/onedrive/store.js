/**
 * store.js — High-level data persistence
 * Transparently encrypts data before saving and decrypts on load.
 * Primary storage: user's OneDrive (/LedgerFlow/*.lf)
 * Fallback (offline): encrypted blobs in localStorage
 */

import { encrypt, decrypt, getSessionKey } from "../crypto/crypto.js";
import { writeFile, readFile, isOneDriveConnected } from "./onedrive.js";

const LOCAL_PREFIX = "lf_enc_";

/**
 * Saves a data object to OneDrive (or localStorage offline fallback).
 * Data is AES-256-GCM encrypted before leaving the browser.
 * @param {string} collection — e.g. "transactions"
 * @param {any} data
 */
export async function saveData(collection, data) {
  const key = getSessionKey();
  const plaintext = JSON.stringify(data);
  const encrypted = await encrypt(key, plaintext);

  const connected = await isOneDriveConnected().catch(() => false);

  if (connected) {
    await writeFile(`${collection}.lf`, encrypted);
  } else {
    // Offline fallback: store encrypted blob in localStorage
    localStorage.setItem(`${LOCAL_PREFIX}${collection}`, JSON.stringify(encrypted));
  }
}

/**
 * Loads and decrypts a data collection from OneDrive or localStorage.
 * Returns null if no data found.
 * @param {string} collection
 * @returns {Promise<any>}
 */
export async function loadData(collection) {
  const key = getSessionKey();
  let encrypted = null;

  const connected = await isOneDriveConnected().catch(() => false);

  if (connected) {
    encrypted = await readFile(`${collection}.lf`);
  }

  // Fall back to localStorage if OneDrive returned nothing
  if (!encrypted) {
    const local = localStorage.getItem(`${LOCAL_PREFIX}${collection}`);
    if (local) encrypted = JSON.parse(local);
  }

  if (!encrypted) return null;

  const plaintext = await decrypt(key, encrypted.iv, encrypted.ciphertext);
  return JSON.parse(plaintext);
}

/**
 * Saves data to localStorage only (no OneDrive), still encrypted.
 * Used in offline / Use-Without-Account mode.
 */
export async function saveLocalOnly(collection, data) {
  const key = getSessionKey();
  const plaintext = JSON.stringify(data);
  const encrypted = await encrypt(key, plaintext);
  localStorage.setItem(`${LOCAL_PREFIX}${collection}`, JSON.stringify(encrypted));
}
