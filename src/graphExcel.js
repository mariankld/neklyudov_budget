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

/**
 * Reads every data row of a named Excel Table (no header row) — used by the sync job and
 * FX-history hydration.
 *
 * Bug fix (2026-08-21): Graph's Tables API has a known quirk where this endpoint returns
 * 404 "ItemNotFound" for a table that currently has zero data rows (header row only),
 * instead of the `{ value: [] }` every other empty collection returns — this is what broke
 * /sync's CurrencyRates refresh ("Could not read "CurrencyRates" rows — Request failed with
 * status code 404") even though the table genuinely exists (its headers read back fine
 * moments earlier in refreshCurrencyRatesTable). Since every caller of this function already
 * only reaches it after establishing the table exists, treat a 404 here as "no data rows
 * yet" and return [] instead of throwing — otherwise any table that's temporarily empty
 * (CurrencyRates cleared, or a category table with no transactions yet) permanently breaks
 * the daily refresh/sync job instead of just reporting nothing to do.
 */
async function getTableRows(driveId, itemId, tableName) {
  const token = await getAccessToken();
  try {
    const { data } = await axios.get(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables('${encodeURIComponent(
        tableName
      )}')/rows`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    return (data.value || []).map((r) => r.values[0]);
  } catch (err) {
    if (err.response && err.response.status === 404) {
      return [];
    }
    if (err.response && err.response.data) {
      const graphError = err.response.data.error || err.response.data;
      err.graphDetail = `[${tableName} rows] ${graphError.code || ""} ${
        graphError.message || JSON.stringify(graphError)
      }`.trim();
      err.message = err.graphDetail;
    }
    throw err;
  }
}

/**
 * Reads a table's data-body range address (e.g. "Staff Spending!A3:N93") — used to translate
 * between "the row I see on screen in Excel" (an absolute worksheet row number) and "the
 * zero-based data-row index" every other helper here uses (getTableRows/updateTableRowByIndex/
 * deleteTableRowByIndex). Needed because a table's data doesn't necessarily start at worksheet
 * row 2 — there can be title rows, frozen panes, etc. above it.
 */
async function getTableDataBodyRangeAddress(driveId, itemId, tableName) {
  const token = await getAccessToken();
  try {
    const { data } = await axios.get(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables('${encodeURIComponent(
        tableName
      )}')/dataBodyRange?$select=address`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    return data.address;
  } catch (err) {
    if (err.response && err.response.data) {
      const graphError = err.response.data.error || err.response.data;
      err.graphDetail = `[${tableName} dataBodyRange] ${graphError.code || ''} ${
        graphError.message || JSON.stringify(graphError)
      }`.trim();
      err.message = err.graphDetail;
    }
    throw err;
  }
}

/** Reads a table's header row (column names in order) — used to locate columns by name instead of a hardcoded index. */
async function getTableHeaders(driveId, itemId, tableName) {
  const token = await getAccessToken();
  try {
    const { data } = await axios.get(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables('${encodeURIComponent(
        tableName
      )}')/headerRowRange`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    return data.values[0];
  } catch (err) {
    if (err.response && err.response.data) {
      const graphError = err.response.data.error || err.response.data;
      err.graphDetail = `[${tableName} headers] ${graphError.code || ""} ${
        graphError.message || JSON.stringify(graphError)
      }`.trim();
      err.message = err.graphDetail;
    }
    throw err;
  }
}

/**
 * Overwrites one existing Excel Table row in place by its zero-based data-row index
 * (`rows/itemAt(index=N)`), e.g. to refresh the CurrencyRates lookup table daily, or to
 * apply a sync-job reconciliation update. `values` must be the full row in column order.
 */
async function updateTableRowByIndex(driveId, itemId, tableName, rowIndex, values) {
  const token = await getAccessToken();
  try {
    const { data } = await axios.patch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables('${encodeURIComponent(
        tableName
      )}')/rows/itemAt(index=${rowIndex})`,
      { values: [values] },
      { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
    );
    return data;
  } catch (err) {
    if (err.response && err.response.data) {
      const graphError = err.response.data.error || err.response.data;
      err.graphDetail = `[${tableName} row ${rowIndex}] ${graphError.code || ''} ${
        graphError.message || JSON.stringify(graphError)
      }`.trim();
      err.message = err.graphDetail;
    }
    throw err;
  }
}

/**
 * Deletes one existing Excel Table row in place by its zero-based data-row index
 * (`rows/itemAt(index=N)`) — used by the sync job to mirror a deletion made on one side
 * (RAW or a category table) onto the other side.
 */
async function deleteTableRowByIndex(driveId, itemId, tableName, rowIndex) {
  const token = await getAccessToken();
  try {
    await axios.delete(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables('${encodeURIComponent(
        tableName
      )}')/rows/itemAt(index=${rowIndex})`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
  } catch (err) {
    if (err.response && err.response.data) {
      const graphError = err.response.data.error || err.response.data;
      err.graphDetail = `[${tableName} delete row ${rowIndex}] ${graphError.code || ''} ${
        graphError.message || JSON.stringify(graphError)
      }`.trim();
      err.message = err.graphDetail;
    }
    throw err;
  }
}

/** Adds a new column to an existing Excel Table (used by the one-time migration script for RowID/SyncHash). */
async function addTableColumn(driveId, itemId, tableName, columnName) {
  const token = await getAccessToken();
  try {
    const { data } = await axios.post(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables('${encodeURIComponent(
        tableName
      )}')/columns`,
      { name: columnName },
      { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
    );
    return data;
  } catch (err) {
    if (err.response && err.response.data) {
      const graphError = err.response.data.error || err.response.data;
      err.graphDetail = `[${tableName} add column ${columnName}] ${graphError.code || ''} ${
        graphError.message || JSON.stringify(graphError)
      }`.trim();
      err.message = err.graphDetail;
    }
    throw err;
  }
}

/**
 * Reads a table's full range address (header row + data rows together), e.g.
 * "Personal Spending!A4:L25" — used before deleting a table so the leftover cells it occupied
 * can be cleared afterward (deleting an Excel Table via Graph un-tableifies the range but leaves
 * the cell values sitting there as plain cells).
 */
async function getTableRangeAddress(driveId, itemId, tableName) {
  const token = await getAccessToken();
  const { data } = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables('${encodeURIComponent(
      tableName
    )}')/range?$select=address`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  return data.address;
}

/** Deletes an Excel Table object (un-tableifies its range; does NOT clear the underlying cell values — call clearRange for that). */
async function deleteTable(driveId, itemId, tableName) {
  const token = await getAccessToken();
  await axios.delete(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables('${encodeURIComponent(
      tableName
    )}')`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
}

/** Clears the contents (not formatting) of an explicit worksheet range — used to wipe leftover cell values after a table has been moved elsewhere and deleted. */
async function clearRangeContents(driveId, itemId, worksheetName, address) {
  const token = await getAccessToken();
  await axios.post(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(
      worksheetName
    )}')/range(address='${address}')/clear`,
    { applyTo: 'Contents' },
    { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
  );
}

/** Lists every worksheet/tab name in the workbook — used to check whether one already exists before creating it (idempotency after a partial/failed run). */
async function listWorksheets(driveId, itemId) {
  const token = await getAccessToken();
  const { data } = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets?$select=name`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  return data.value.map((w) => w.name);
}

/** Lists every worksheet/tab with its current left-to-right tab position (0-based) — used to reorder tabs. */
async function listWorksheetsWithPosition(driveId, itemId) {
  const token = await getAccessToken();
  const { data } = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets?$select=name,position`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  return data.value.map((w) => ({ name: w.name, position: w.position })).sort((a, b) => a.position - b.position);
}

/** Moves a worksheet/tab to a specific 0-based position, reordering the tab strip (other tabs shift to make room). */
async function setWorksheetPosition(driveId, itemId, worksheetName, position) {
  const token = await getAccessToken();
  const { data } = await axios.patch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(
      worksheetName
    )}')`,
    { position },
    { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
  );
  return data;
}

/** Adds a brand-new worksheet/tab to the workbook (e.g. for splitting a combined tab into two). */
async function addWorksheet(driveId, itemId, name) {
  const token = await getAccessToken();
  const { data } = await axios.post(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets/add`,
    { name },
    { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
  );
  return data;
}

/** Renames a worksheet/tab (distinct from renameTable, which renames the Excel Table object, not the tab itself). */
async function renameWorksheet(driveId, itemId, currentName, newName) {
  const token = await getAccessToken();
  const { data } = await axios.patch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(
      currentName
    )}')`,
    { name: newName },
    { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
  );
  return data;
}

/** Creates a brand-new Excel Table on an existing worksheet from a starting range (e.g. "A1:E1" with headers already in row 1). */
async function createTable(driveId, itemId, worksheetName, address, hasHeaders = true) {
  const token = await getAccessToken();
  const { data } = await axios.post(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(
      worksheetName
    )}')/tables/add`,
    // Relative address only ("A1:L1") — the endpoint is already scoped to worksheetName via the
    // URL path above. Prefixing it here too (e.g. "Staff Advances!A1:L1") broke for any sheet
    // name containing a space (Graph 400 InvalidArgument, reproduced 2026-09-13 on "Staff
    // Advances").
    { address, hasHeaders },
    { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
  );
  return data;
}

/** Renames a table (Graph creates tables as "Table1" etc. by default). */
async function renameTable(driveId, itemId, currentName, newName) {
  const token = await getAccessToken();
  const { data } = await axios.patch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables('${encodeURIComponent(
      currentName
    )}')`,
    { name: newName },
    { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
  );
  return data;
}

/** Writes literal values (not formulas) into an explicit range — used to set header text for a new table. */
async function setRangeValues(driveId, itemId, worksheetName, address, values) {
  const token = await getAccessToken();
  const { data } = await axios.patch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(
      worksheetName
    )}')/range(address='${address}')`,
    { values },
    { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
  );
  return data;
}

/** Reads a worksheet range's formulas (2D array), e.g. to audit/patch Summary tab SUMIFS chains. */
async function getRangeFormulas(driveId, itemId, worksheetName, address) {
  const token = await getAccessToken();
  const { data } = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(
      worksheetName
    )}')/range(address='${address}')?$select=formulas,address`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  return data;
}

/** Writes formulas (2D array, cells starting with "=") into an explicit range. */
async function setRangeFormulas(driveId, itemId, worksheetName, address, formulas) {
  const token = await getAccessToken();
  const { data } = await axios.patch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(
      worksheetName
    )}')/range(address='${address}')`,
    { formulas },
    { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
  );
  return data;
}

/** Reads a worksheet range's numberFormat (2D array) — used to diagnose date/currency display bugs. */
async function getRangeNumberFormat(driveId, itemId, worksheetName, address) {
  const token = await getAccessToken();
  const { data } = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(
      worksheetName
    )}')/range(address='${address}')?$select=numberFormat,values,address`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  return data;
}

/** Writes a numberFormat (2D array, e.g. [["dd/mm/yyyy"]]) onto an explicit range — cell values are untouched. */
async function setRangeNumberFormat(driveId, itemId, worksheetName, address, numberFormat) {
  const token = await getAccessToken();
  const { data } = await axios.patch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(
      worksheetName
    )}')/range(address='${address}')`,
    { numberFormat },
    { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
  );
  return data;
}

/** Gets a worksheet's used range (values + formulas + address) — the entry point for the Summary formula audit. */
async function getUsedRange(driveId, itemId, worksheetName) {
  const token = await getAccessToken();
  const { data } = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(
      worksheetName
    )}')/usedRange(valuesOnly=false)?$select=formulas,values,address`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  return data;
}

module.exports = {
  getAccessToken,
  resolveShareToDriveItem,
  listWorkbookTables,
  appendTableRow,
  getTableRows,
  getTableDataBodyRangeAddress,
  getTableHeaders,
  updateTableRowByIndex,
  deleteTableRowByIndex,
  addTableColumn,
  getTableRangeAddress,
  deleteTable,
  clearRangeContents,
  listWorksheets,
  listWorksheetsWithPosition,
  setWorksheetPosition,
  addWorksheet,
  renameWorksheet,
  createTable,
  renameTable,
  setRangeValues,
  getRangeFormulas,
  setRangeFormulas,
  getRangeNumberFormat,
  setRangeNumberFormat,
  getUsedRange,
};
