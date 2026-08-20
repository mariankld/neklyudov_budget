const { getTableRows, getTableHeaders, updateTableRowByIndex } = require("./graphExcel");
const fxRates = require("./fxRates");

/**
 * Daily maintenance job: refreshes the workbook's "today's rate" lookup table
 * (CurrencyRates by default) so anyone manually typing a row into Excel still has an
 * up-to-date VLOOKUP source to reference. This is deliberately separate from the
 * bot's own writes, which never read this table — every bot-logged expense freezes its
 * own historical rate via fxRates.getRateToHkd at write time (see index.js /
 * appendTransactionToExcel) and never depends on "today's" rate.
 *
 * Column layout is NOT hardcoded: the table's header row is read first and columns are
 * located by name (case-insensitive substring match), so this keeps working even if the
 * table gains/loses/reorders columns for anything other than the code/rate/timestamp
 * fields it manages. Any other column (e.g. a currency display name) is left untouched.
 */

function formatDdMmYyyy(date) {
  const d = date.getDate().toString().padStart(2, "0");
  const m = (date.getMonth() + 1).toString().padStart(2, "0");
  const y = date.getFullYear();
  return `${d}/${m}/${y}`;
}

/** Finds the index of the first header whose name contains any of `needles` (case-insensitive). */
function findColumnIndex(headers, needles) {
  const lower = headers.map((h) => String(h || "").trim().toLowerCase());
  for (let i = 0; i < lower.length; i++) {
    if (needles.some((n) => lower[i].includes(n))) return i;
  }
  return -1;
}

/**
 * Overwrites every row's rate + last-updated columns in the CurrencyRates table with
 * today's rate for that row's currency. Rows for currencies not in fxRates.TRACKED_CURRENCIES
 * (or that fail to resolve a currency code) are left untouched and reported as skipped.
 */
async function refreshCurrencyRatesTable({ driveId, itemId, tableName }) {
  const report = { updated: 0, skipped: 0, errors: [] };

  let headers;
  try {
    headers = await getTableHeaders(driveId, itemId, tableName);
  } catch (err) {
    report.errors.push(`Could not read "${tableName}" headers — ${err.message}`);
    return report;
  }

  const codeCol = findColumnIndex(headers, ["код", "code", "currency", "валюта"]);
  const rateCol = findColumnIndex(headers, ["курс", "rate"]);
  const updatedCol = findColumnIndex(headers, ["обновлен", "updated", "update"]);

  if (codeCol === -1 || rateCol === -1) {
    report.errors.push(
      `"${tableName}" headers ${JSON.stringify(headers)} — could not find both a currency-code column and a rate column by name. Refusing to guess; fix headers or pass an explicit column mapping.`
    );
    return report;
  }

  let rows;
  try {
    rows = await getTableRows(driveId, itemId, tableName);
  } catch (err) {
    report.errors.push(`Could not read "${tableName}" rows — ${err.message}`);
    return report;
  }

  const today = formatDdMmYyyy(new Date());
  const nowIso = new Date().toISOString();

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const currency = String(row[codeCol] || "").trim().toUpperCase();

    if (!currency || !fxRates.TRACKED_CURRENCIES.includes(currency)) {
      report.skipped++;
      continue;
    }

    try {
      const rate = currency === "HKD" ? 1 : await fxRates.getRateToHkd(currency, today);
      const updated = [...row];
      updated[rateCol] = rate;
      if (updatedCol !== -1) updated[updatedCol] = nowIso;
      await updateTableRowByIndex(driveId, itemId, tableName, i, updated);
      report.updated++;
    } catch (err) {
      report.errors.push(`${tableName} row ${i + 1} (${currency}): ${err.message}`);
    }
  }

  return report;
}

function formatDailyJobsReport({ currencyRates }) {
  const lines = [
    `CurrencyRates: updated ${currencyRates.updated}, skipped ${currencyRates.skipped}.`,
  ];
  if (currencyRates.errors.length) {
    lines.push(`⚠️ ${currencyRates.errors.length} CurrencyRates error(s):`);
    currencyRates.errors.slice(0, 10).forEach((e) => lines.push(`- ${e}`));
    if (currencyRates.errors.length > 10) lines.push(`...and ${currencyRates.errors.length - 10} more.`);
  }
  return lines.join("\n");
}

module.exports = {
  refreshCurrencyRatesTable,
  formatDailyJobsReport,
};
