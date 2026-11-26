/**
 * Configuration for Outlook MCP Server
 */
const path = require('path');
const os = require('os');

// Load .env from the repo directory, not the process CWD (Claude Desktop launches
// MCP servers from an unrelated working directory). Existing env vars win.
require('dotenv').config({ path: process.env.OUTLOOK_ENV_FILE || path.join(__dirname, '.env'), quiet: true });

// Ensure we have a home directory path even if process.env.HOME is undefined
const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp';

const DEFAULT_SCOPES =
  'offline_access User.Read Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Calendars.ReadWrite Contacts.Read';

/**
 * Resolve where this instance keeps its OAuth tokens. One file per account lets
 * several instances (one per mailbox) run side by side.
 *   1. OUTLOOK_TOKEN_PATH  - explicit file path (~ expanded)
 *   2. OUTLOOK_ACCOUNT     - ~/.outlook-mcp/<account>.json
 *   3. legacy default      - ~/.outlook-mcp-tokens.json
 */
function resolveTokenPath(env = process.env, home = homeDir) {
  const expand = (p) => (p.startsWith('~') ? path.join(home, p.slice(1)) : p);
  if (env.OUTLOOK_TOKEN_PATH) return path.resolve(expand(env.OUTLOOK_TOKEN_PATH));
  if (env.OUTLOOK_ACCOUNT) {
    const safe = String(env.OUTLOOK_ACCOUNT).trim().replace(/[^A-Za-z0-9._-]/g, '_');
    return path.join(home, '.outlook-mcp', `${safe}.json`);
  }
  return path.join(home, '.outlook-mcp-tokens.json');
}

const tenantId = process.env.MS_TENANT_ID || process.env.OUTLOOK_TENANT_ID || 'common';
const authorityHost = (process.env.MS_AUTHORITY_HOST || 'https://login.microsoftonline.com').replace(/\/+$/, '');

module.exports = {
  resolveTokenPath,
  DEBUG: process.env.OUTLOOK_DEBUG === 'true',
  // Server information
  SERVER_NAME: "m365-assistant",
  SERVER_VERSION: "2.0.0",
  
  // Test mode setting
  USE_TEST_MODE: process.env.USE_TEST_MODE === 'true',
  
  // Authentication configuration
  AUTH_CONFIG: {
    // Accept both naming schemes (README/.env.example used OUTLOOK_*, the auth server used MS_*)
    clientId: process.env.OUTLOOK_CLIENT_ID || process.env.MS_CLIENT_ID || '',
    clientSecret: process.env.OUTLOOK_CLIENT_SECRET || process.env.MS_CLIENT_SECRET || '',
    tenantId,
    authorityHost,
    tokenEndpoint: process.env.MS_TOKEN_ENDPOINT || `${authorityHost}/${tenantId}/oauth2/v2.0/token`,
    redirectUri: process.env.MS_REDIRECT_URI || 'http://localhost:3333/auth/callback',
    scopes: (process.env.MS_SCOPES || DEFAULT_SCOPES).split(/\s+/).filter(Boolean),
    account: process.env.OUTLOOK_ACCOUNT || 'default',
    tokenStorePath: resolveTokenPath(),
    authServerUrl: 'http://localhost:3333'
  },
  
  // Microsoft Graph API
  GRAPH_API_ENDPOINT: 'https://graph.microsoft.com/v1.0/',
  
  // Calendar constants
  CALENDAR_SELECT_FIELDS: 'id,subject,start,end,location,bodyPreview,isAllDay,recurrence,attendees',

  // Email constants
  EMAIL_SELECT_FIELDS: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead',
  EMAIL_DETAIL_FIELDS: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead,internetMessageHeaders',
  
  // Calendar constants
  CALENDAR_SELECT_FIELDS: 'id,subject,bodyPreview,start,end,location,organizer,attendees,isAllDay,isCancelled',
  
  // Pagination
  DEFAULT_PAGE_SIZE: 25,
  MAX_RESULT_COUNT: 50,

  // Timezone
  DEFAULT_TIMEZONE: "Central European Standard Time",

  // OneDrive constants
  ONEDRIVE_SELECT_FIELDS: 'id,name,size,lastModifiedDateTime,webUrl,folder,file,parentReference',
  ONEDRIVE_UPLOAD_THRESHOLD: 4 * 1024 * 1024, // 4MB - files larger than this need chunked upload

  // Power Automate / Flow constants
  FLOW_API_ENDPOINT: 'https://api.flow.microsoft.com',
  FLOW_SCOPE: 'https://service.flow.microsoft.com/.default',
};
