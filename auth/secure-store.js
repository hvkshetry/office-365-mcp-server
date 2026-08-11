/**
 * Secure token cache storage for Office MCP.
 *
 * On Windows, the token bundle (access_token, refresh_token, id metadata, etc.) is
 * encrypted at rest using the Windows Data Protection API (DPAPI, CurrentUser scope)
 * before being written to disk. Decryption is bound to the current Windows user
 * account on this machine, so the on-disk ciphertext is useless if copied to another
 * machine, another user profile, a backup, or a cloud-sync folder (OneDrive, etc.).
 *
 * DPAPI is invoked through Windows PowerShell
 * (System.Security.Cryptography.ProtectedData), so there is NO native build
 * dependency (no node-gyp / win-dpapi).
 *
 * On non-Windows platforms it falls back to the previous plaintext-at-0600 behavior.
 * The server is primarily run on Windows; macOS/Linux operators should prefer an
 * MSAL-based flow backed by Keychain / libsecret.
 *
 * THREAT MODEL: this protects tokens AT REST (file copy, cloud sync, backup, another
 * user or admin reading the file). It does NOT defend against malicious code already
 * running as the same Windows user, which can ask DPAPI to decrypt exactly as this
 * module does. Use Conditional Access / token protection for that layer.
 */
const fs = require('fs');
const { execFileSync } = require('child_process');
const config = require('../config');

const TOKEN_PATH = config.AUTH_CONFIG.tokenStorePath;
const FORMAT = 'dpapi-cu-v1';
// App-specific optional entropy mixed into DPAPI. NOT a secret; it only scopes the
// blob to this application so unrelated CurrentUser DPAPI blobs can't be cross-used.
const ENTROPY = 'office-mcp-token-store-v1';
const isWindows = process.platform === 'win32';

// PowerShell payload: reads base64 from stdin, runs Protect/Unprotect per
// $env:DPAPI_MODE under CurrentUser scope with fixed app entropy, writes base64
// result to stdout (no trailing newline via [Console]::Out.Write).
const PS_SCRIPT = `
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Security
$mode = $env:DPAPI_MODE
$entropy = [System.Text.Encoding]::UTF8.GetBytes('${ENTROPY}')
$scope = [System.Security.Cryptography.DataProtectionScope]::CurrentUser
$inB64 = [Console]::In.ReadToEnd().Trim()
$inBytes = [System.Convert]::FromBase64String($inB64)
if ($mode -eq 'protect') {
  $out = [System.Security.Cryptography.ProtectedData]::Protect($inBytes, $entropy, $scope)
} else {
  $out = [System.Security.Cryptography.ProtectedData]::Unprotect($inBytes, $entropy, $scope)
}
[Console]::Out.Write([System.Convert]::ToBase64String($out))
`;

function runDpapi(mode, inputBuffer) {
  // -EncodedCommand expects UTF-16LE base64; data flows via stdin (never the cmdline).
  const encoded = Buffer.from(PS_SCRIPT, 'utf16le').toString('base64');
  const stdout = execFileSync(
    'powershell.exe',
    ['-NoProfile', '-NonInteractive', '-EncodedCommand', encoded],
    {
      input: inputBuffer.toString('base64'),
      env: { ...process.env, DPAPI_MODE: mode },
      maxBuffer: 16 * 1024 * 1024,
      windowsHide: true,
    }
  );
  return Buffer.from(stdout.toString('utf8').trim(), 'base64');
}

function encryptToFileContents(tokensObj) {
  const plaintext = Buffer.from(JSON.stringify(tokensObj), 'utf8');
  const ciphertext = runDpapi('protect', plaintext);
  return JSON.stringify({ format: FORMAT, ciphertext: ciphertext.toString('base64') }, null, 2);
}

function decryptCiphertext(ciphertextB64) {
  const ciphertext = Buffer.from(ciphertextB64, 'base64');
  const plaintext = runDpapi('unprotect', ciphertext);
  return JSON.parse(plaintext.toString('utf8'));
}

function atomicWrite(contents) {
  const tempPath = TOKEN_PATH + '.tmp';
  fs.writeFileSync(tempPath, contents, { mode: 0o600 });
  fs.renameSync(tempPath, TOKEN_PATH);
  try { fs.chmodSync(TOKEN_PATH, 0o600); } catch (e) { /* Windows may not support chmod */ }
}

/**
 * Reads and returns the token object, transparently decrypting on Windows.
 * Transparently migrates a legacy plaintext file to encrypted-at-rest on first
 * read (Windows only).
 * @returns {object|null}
 */
function readTokens() {
  try {
    if (!fs.existsSync(TOKEN_PATH)) {
      return null;
    }
    const raw = fs.readFileSync(TOKEN_PATH, 'utf8');
    let parsed;
    try {
      parsed = JSON.parse(raw);
    } catch (e) {
      console.error('[SECURE-STORE] Token file is not valid JSON');
      return null;
    }

    // Encrypted format
    if (parsed && parsed.format === FORMAT && parsed.ciphertext) {
      if (!isWindows) {
        console.error('[SECURE-STORE] Encrypted token file found but not on Windows; cannot decrypt');
        return null;
      }
      try {
        return decryptCiphertext(parsed.ciphertext);
      } catch (e) {
        console.error('[SECURE-STORE] Failed to decrypt token cache (re-authentication required)');
        return null;
      }
    }

    // Legacy plaintext format (token fields at top level)
    if (parsed && parsed.access_token) {
      if (isWindows) {
        // One-time migration: re-write encrypted at rest.
        try {
          atomicWrite(encryptToFileContents(parsed));
          console.error('[SECURE-STORE] Migrated plaintext token cache to DPAPI-encrypted at rest');
        } catch (e) {
          console.error('[SECURE-STORE] Plaintext->encrypted migration failed (continuing):', e.message);
        }
      }
      return parsed;
    }

    return null;
  } catch (error) {
    console.error('[SECURE-STORE] Error reading tokens:', error.message);
    return null;
  }
}

/**
 * Writes the token object, encrypting at rest on Windows (plaintext 0600 elsewhere).
 * @param {object} tokensObj
 * @returns {boolean}
 */
function writeTokens(tokensObj) {
  try {
    const contents = isWindows
      ? encryptToFileContents(tokensObj)
      : JSON.stringify(tokensObj, null, 2);
    atomicWrite(contents);
    return true;
  } catch (error) {
    console.error('[SECURE-STORE] Error writing tokens:', error.message);
    return false;
  }
}

module.exports = {
  readTokens,
  writeTokens,
  TOKEN_PATH,
  encryptionAvailable: isWindows,
};
