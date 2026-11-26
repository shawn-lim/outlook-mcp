/**
 * Authentication module for Outlook MCP server
 */
const fs = require('fs');
const config = require('../config');
const tokenManager = require('./token-manager');
const TokenStorage = require('./token-storage');
const { authTools } = require('./tools');

let storage = null;
let loadedMtime = 0;

/**
 * Shared TokenStorage for this process. If the token file was rewritten on disk
 * (e.g. the auth server just completed a fresh sign-in), drop the in-memory copy
 * so the new tokens are picked up without restarting Claude.
 */
function getTokenStorage() {
  if (!storage) storage = new TokenStorage();
  try {
    const mtime = fs.statSync(storage.config.tokenStorePath).mtimeMs;
    if (mtime !== loadedMtime) {
      storage.tokens = null;
      loadedMtime = mtime;
    }
  } catch (_) {
    // No token file yet; getValidAccessToken() will report "not authenticated".
  }
  return storage;
}

/**
 * Ensures the user is authenticated and returns an access token, refreshing it
 * with the stored refresh token when it is expired or about to expire.
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated(forceNew = false) {
  if (forceNew) {
    throw new Error('Authentication required');
  }

  // Test mode keeps using the simple token-manager cache with mock tokens.
  if (config.USE_TEST_MODE) {
    const testToken = tokenManager.getAccessToken();
    if (!testToken) throw new Error('Authentication required');
    return testToken;
  }

  const ts = getTokenStorage();
  const accessToken = await ts.getValidAccessToken();
  if (!accessToken) {
    throw new Error('Authentication required');
  }
  // Our own refresh just rewrote the file; remember its mtime so we don't reload needlessly.
  try { loadedMtime = fs.statSync(ts.config.tokenStorePath).mtimeMs; } catch (_) {}
  return accessToken;
}

module.exports = {
  tokenManager,
  authTools,
  ensureAuthenticated,
  getTokenStorage
};
