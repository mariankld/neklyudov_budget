const axios = require('axios');

const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;

let cachedToken = null;
let cachedTokenExpiry = 0;

/**
 * Acquires an app-only (client credentials) access token for Microsoft Graph.
 * Requires the app registration to have ADMIN-CONSENTED application permissions
 * (Files.ReadWrite.All and/or Sites.ReadWrite.All) — see README "Migrating to
 * Excel via Microsoft Graph" section. Until consent is granted this will fail
 * with a 401/403 from Graph, not from this function.
 */
async function getAccessToken() {
  if (cachedToken && Date.now() < cachedTokenExpiry - 60_000) {
    return cachedToken;
  }
  if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
    throw new Error('Missing AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET in .env');
  }
  const url = `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: AZURE_CLIENT_ID,
    client_secret: AZURE_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });
  const { data } = await axios.post(url, body.toString(), {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
  });
  cachedToken = data.access_token;
  cachedTokenExpiry = Date.now() + data.expires_in * 1000;
  return cachedToken;
}

/** Base64url-encodes a sharing URL into the Graph "shares/{shareId}" format. */
function encodeShareUrl(shareUrl) {
  const base64 = Buffer.from(shareUrl, 'utf8').toString('base64');
  const base64url = base64.replace(/=/g, '').replace(/\//g, '_').replace(/\+/g, '-');
  return `u!${base64url}`;
}

/**
 * Turns a SharePoint/OneDrive share link into the {driveId, itemId} pair Graph
 * needs for every other call. Run this once via `npm run resolve-excel-share`
 * and paste the results into .env as EXCEL_DRIVE_ID / EXCEL_ITEM_ID so we don't
 * re-resolve the link on every request.
 */
async function resolveShareToDriveItem(shareUrl) {
  const token = await getAccessToken();
  const shareId = encodeShareUrl(shareUrl);
  const { data } = await axios.get(`https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  return { driveId: data.parentReference.driveId, itemId: data.id, name: data.name };
}

/** Lists worksheets + Excel Tables in the workbook — useful for sanity-checking structure. */
async function listWorkbookTables(driveId, itemId) {
  const token = await getAccessToken();
  const { data } = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables?$expand=worksheet`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  return data.value.map((t) => ({ name: t.name, worksheet: t.worksheet && t.worksheet.name }));
}

/**
 * Appends one row to a named Excel Table via the Graph Tables API — the Graph
 * equivalent of the googleapis `spreadsheets.values.append` call the Google
 * Sheets version uses. `values` must be in the same column order as the table.
 * Every sheet in Family_Budget_v2.xlsx is already a proper Excel Table
 * (confirmed: RAW, Shopping, Income, TransportTable, etc.), so this endpoint
 * works everywhere without needing to track "last used row" ourselves.
 */
async function appendTableRow(driveId, itemId, tableName, values) {
  const token = await getAccessToken();
  try {
    const { data } = await axios.post(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables('${encodeURIComponent(
        tableName
      )}')/rows`,
      { values: [values] },
      { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
    );
    return data;
  } catch (err) {
    // Graph's real error (bad column count, invalid value for a calculated
    // column, etc.) is in the response body — axios's err.message is just
    // "Request failed with status code 400" and hides it. Surface it so it
    // actually reaches the logs / Telegram error message.
    if (err.response && err.response.data) {
      const graphError = err.response.data.error || err.response.data;
      err.graphDetail = `[${tableName}] ${graphError.code || ''} ${graphError.message || JSON.stringify(graphError)}`.trim();
      err.message = err.graphDetail;
    }
    throw err;
  }
}

module.exports = {
  getAccessToken,
  resolveShareToDriveItem,
  listWorkbookTables,
  appendTableRow,
};
