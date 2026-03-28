/**
 * crypto.js — Web Crypto API wrapper
 * AES-256-GCM encryption with PBKDF2 key derivation.
 * The encryption key lives ONLY in memory — never written to disk or server.
 */

// In-memory key — cleared on logout/page unload
let _sessionKey = null;

/**
 * Derives an AES-256-GCM CryptoKey from a user password + salt using PBKDF2.
 * @param {string} password
 * @param {string} salt — base64-encoded 16-byte salt
 * @returns {Promise<CryptoKey>}
 */
export async function deriveKey(password, salt) {
  const enc = new TextEncoder();
  const keyMaterial = await crypto.subtle.importKey(
    "raw",
    enc.encode(password),
    "PBKDF2",
    false,
    ["deriveKey"]
  );
  return crypto.subtle.deriveKey(
    {
      name: "PBKDF2",
      salt: enc.encode(salt),
      iterations: 310_000, // OWASP 2024 recommendation for PBKDF2-SHA256
      hash: "SHA-256",
    },
    keyMaterial,
    { name: "AES-GCM", length: 256 },
    false,
    ["encrypt", "decrypt"]
  );
}

/**
 * Generates a random 16-byte salt, returned as base64.
 * @returns {string}
 */
export function generateSalt() {
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  return btoa(String.fromCharCode(...bytes));
}

/**
 * Encrypts a plaintext string with AES-256-GCM.
 * @param {CryptoKey} key
 * @param {string} plaintext
 * @returns {Promise<{iv: string, ciphertext: string}>} — both base64-encoded
 */
export async function encrypt(key, plaintext) {
  const enc = new TextEncoder();
  const iv = new Uint8Array(12);
  crypto.getRandomValues(iv);

  const ciphertextBuffer = await crypto.subtle.encrypt(
    { name: "AES-GCM", iv },
    key,
    enc.encode(plaintext)
  );

  return {
    iv: btoa(String.fromCharCode(...iv)),
    ciphertext: btoa(String.fromCharCode(...new Uint8Array(ciphertextBuffer))),
  };
}

/**
 * Decrypts an AES-256-GCM encrypted payload.
 * @param {CryptoKey} key
 * @param {string} iv — base64-encoded
 * @param {string} ciphertext — base64-encoded
 * @returns {Promise<string>} plaintext
 */
export async function decrypt(key, iv, ciphertext) {
  const ivBytes = Uint8Array.from(atob(iv), (c) => c.charCodeAt(0));
  const ciphertextBytes = Uint8Array.from(atob(ciphertext), (c) => c.charCodeAt(0));

  const plaintextBuffer = await crypto.subtle.decrypt(
    { name: "AES-GCM", iv: ivBytes },
    key,
    ciphertextBytes
  );

  return new TextDecoder().decode(plaintextBuffer);
}

// ── Session key management ────────────────────────────────────────────────────

/** Store derived key in memory for this session. */
export function setSessionKey(key) {
  _sessionKey = key;
}

/** Retrieve the in-memory session key.  Throws if not yet set. */
export function getSessionKey() {
  if (!_sessionKey) throw new Error("Encryption key not set. Please enter your data password.");
  return _sessionKey;
}

/** Wipe the key from memory (call on logout). */
export function clearSessionKey() {
  _sessionKey = null;
}

/** Returns true if a session key is currently loaded. */
export function isKeyLoaded() {
  return _sessionKey !== null;
}
