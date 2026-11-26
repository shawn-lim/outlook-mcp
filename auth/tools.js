/**
 * Authentication-related tools for the Outlook MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');

/**
 * About tool handler
 * @returns {object} - MCP response
 */
async function handleAbout() {
  return {
    content: [{
      type: "text",
      text: `M365 Assistant MCP Server v${config.SERVER_VERSION}\n\nProvides access to Microsoft 365 services through Microsoft Graph API:\n- Outlook (email, calendar, folders, rules)\n- OneDrive (files, folders, sharing)\n- Power Automate (flows, environments, runs)\n\nModular architecture for improved maintainability.`
    }]
  };
}

/**
 * Authentication tool handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAuthenticate(args) {
  const force = args && args.force === true;
  
  // For test mode, create a test token
  if (config.USE_TEST_MODE) {
    // Create a test token with a 1-hour expiry
    tokenManager.createTestTokens();
    
    return {
      content: [{
        type: "text",
        text: 'Successfully authenticated with Microsoft Graph API (test mode)'
      }]
    };
  }
  
  // For real authentication, generate an auth URL and instruct the user to visit it
  const { account, authServerUrl } = config.AUTH_CONFIG;
  const authUrl = `${authServerUrl}/auth`;

  return {
    content: [{
      type: "text",
      text: `To sign in account "${account}":\n1. In the outlook-mcp folder run: node outlook-auth-server.js --account ${account}\n2. Open ${authUrl} and complete the Microsoft sign-in.\nTokens then refresh automatically; no restart needed.`
    }]
  };
}

/**
 * Check authentication status tool handler
 * @returns {object} - MCP response
 */
async function handleCheckAuthStatus() {
  const { account, tokenStorePath } = config.AUTH_CONFIG;
  if (config.USE_TEST_MODE) {
    const tokens = tokenManager.loadTokenCache();
    return { content: [{ type: "text", text: tokens ? "Authenticated (test mode)" : "Not authenticated (test mode)" }] };
  }
  // Lazy require avoids a circular import (auth/index.js requires this file).
  const { getTokenStorage } = require('./index');
  try {
    const ts = getTokenStorage();
    const token = await ts.getValidAccessToken(); // refreshes if expired
    if (!token) {
      return { content: [{ type: "text", text: `Not authenticated (account: ${account}). Run: node outlook-auth-server.js --account ${account}  then open http://localhost:3333/auth` }] };
    }
    const mins = Math.round((ts.getExpiryTime() - Date.now()) / 60000);
    return { content: [{ type: "text", text: `Authenticated and ready (account: ${account}, access token valid ~${mins} min, auto-refreshes). Token file: ${tokenStorePath}` }] };
  } catch (error) {
    return { content: [{ type: "text", text: `Auth check failed (account: ${account}): ${error.message}` }] };
  }
}

// Tool definitions
const authTools = [
  {
    name: "about",
    description: "Returns information about this M365 Assistant server",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleAbout
  },
  {
    name: "authenticate",
    description: "Authenticate with Microsoft Graph API to access Outlook data",
    inputSchema: {
      type: "object",
      properties: {
        force: {
          type: "boolean",
          description: "Force re-authentication even if already authenticated"
        }
      },
      required: []
    },
    handler: handleAuthenticate
  },
  {
    name: "check-auth-status",
    description: "Check the current authentication status with Microsoft Graph API",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleCheckAuthStatus
  }
];

module.exports = {
  authTools,
  handleAbout,
  handleAuthenticate,
  handleCheckAuthStatus
};
