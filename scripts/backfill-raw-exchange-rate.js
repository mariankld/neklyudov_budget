require("dotenv").config();
const { getTableRows, getTableHeaders, updateTableRowByIndex } = require("../src/graphExcel");
const fxRates = require("../src/fxRates");
const { RAW_COLS, catColsFor } = require("../src/syncJob");
const CATEGORY_TABLE_MAP = require("./lib/categoryTableMap");

// One-off backfill: fills the new "Exchange Rate" column (added to RAW 2026-08-21, right after
// Currency) on every existing RAW row logged before the column existed. Idempotent — never
// overwrites a cell that already has a value, so it's safe to re-run or resume after a partial
// run / transient error.
//
// Source of the rate, per row:
//   - HKD rows: always 1, no lookup needed.
//   - Expense rows with a matching category-table row (via RowID): copies that row's own Курс
//     cell — the exact rate already frozen for this transaction when it was first logged, no
//     network call, guaranteed to agree with the category tab.
//   - Everything else (Income rows, which have no category counterpart, or expense rows whose
//     category counterpart is missing): re-fetches the historical rate for that row's own
//     Date+Currency via fxRates.getRateToHkd — the same historical-rate source the bot itself
//     uses, so it reproduces the number the bot would have written at the time.
//
// Run this AFTER the "Exchange Rate" column has actually been added to the live RAW table
// (Mariya is adding it by hand in Excel) — the script looks the column up by header name and
// refuses to guess a position if it can't find one.
//
// Run:
//   node scripts/backfill-raw-exchange-rate.js            (dry run, prints planned writes)
//   node scripts/backfill-raw-exchange-rate.js --apply     (writes them)

const RAW_SHEET_NAME = (process.env.RAW_SHEET_NAME || "RAW").trim();
const APPLY = process.argv.includes("--apply");

function findColumnIndex(headers, needles) {
  const lower = headers.map((h) => String(h || "").trim().toLowerCase());
  for (let i = 0; i < lower.length; i++) {
    if (needles.some((n) => lower[i].includes(n))) return i;
  }
  return -1;
}

// Graph returns a table cell's raw value, not its display text. Most RAW rows store the Date
// column as the "dd/mm/yyyy" string the bot writes, but some rows (older/manually-touched ones)
// have Excel's own date type underneath, which Graph hands back as a bare serial number (days
// since 1899-12-30, the classic Excel/Lotus epoch) — e.g. 46098, not "17/03/2026". Passing that
// straight to fxRates.getRateToHkd blew up ("Invalid date for FX lookup") or silently asked CBR
// for a nonsense date. Convert serials to dd/mm/yyyy before doing anything else with them.
function normalizeDate(rawDate) {
  const asString = String(rawDate == null ? "" : rawDate).trim();
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(asString)) return asString;

  const serial = Number(asString);
  if (asString && Number.isFinite(serial) && serial > 0) {
    const ms = Date.UTC(1899, 11, 30) + serial * 86400000;
    const d = new Date(ms);
    const dd = String(d.getUTCDate()).padStart(2, "0");
    const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
    const yyyy = d.getUTCFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }
  return asString;
}

(async () => {
  const { EXCEL_DRIVE_ID, EXCEL_ITEM_ID } = process.env;
  if (!EXCEL_DRIVE_ID || !EXCEL_ITEM_ID) {
    console.error("Set EXCEL_DRIVE_ID and EXCEL_ITEM_ID in .env first.");
    process.exit(1);
  }

  console.log(`\n=== RAW "Exchange Rate" backfill — ${APPLY ? "APPLY" : "DRY RUN"} ===\n`);

  let headers;
  try {
    headers = await getTableHeaders(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, RAW_SHEET_NAME);
  } catch (err) {
    console.error(`Could not read "${RAW_SHEET_NAME}" headers — ${err.message}`);
    process.exit(1);
  }

  const rateCol = findColumnIndex(headers, ["exchange rate"]);
  if (rateCol === -1) {
    console.error(
      `Could not find an "Exchange Rate" column in "${RAW_SHEET_NAME}" (headers: ${JSON.stringify(headers)}). Add the column first (right after Currency), then re-run.`
    );
    process.exit(1);
  }
  console.log(`Found "Exchange Rate" at column index ${rateCol} (0-based).\n`);

  let rawRows;
  try {
    rawRows = await getTableRows(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, RAW_SHEET_NAME);
  } catch (err) {
    console.error(`Could not read "${RAW_SHEET_NAME}" rows — ${err.message}`);
    process.exit(1);
  }

  // Pre-load every category table once (rather than per-row) and index by RowID, so expense
  // rows can pull their already-frozen Курс without a network call.
  const categoryRateByRowId = new Map(); // RowID -> rate
  const tableNames = [...new Set(Object.values(CATEGORY_TABLE_MAP).filter((v) => typeof v === "string"))];
  for (const tableName of tableNames) {
    const catCols = catColsFor(tableName);
    let catRows;
    try {
      catRows = await getTableRows(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, tableName);
    } catch (err) {
      console.error(`Could not read "${tableName}" — ${err.message}. Rows depending on it will fall back to a fresh FX lookup.`);
      continue;
    }
    for (const row of catRows) {
      const rowId = row[catCols.rowId];
      const rate = row[catCols.rate];
      if (rowId && rate !== "" && rate != null && Number.isFinite(Number(rate))) {
        categoryRateByRowId.set(String(rowId), Number(rate));
      }
    }
  }

  let filled = 0;
  let skippedAlreadySet = 0;
  let fromCategory = 0;
  let fromFreshLookup = 0;
  const errors = [];

  for (let i = 0; i < rawRows.length; i++) {
    const row = rawRows[i];
    const existing = row[rateCol];
    if (existing !== "" && existing != null) {
      skippedAlreadySet++;
      continue;
    }

    const date = normalizeDate(row[RAW_COLS.date]);
    const currency = String(row[RAW_COLS.currency] || "").trim().toUpperCase();
    const rowId = row[RAW_COLS.rowId];

    if (!currency) {
      errors.push(`Row ${i + 1}: no Currency value, skipped.`);
      continue;
    }

    let rate;
    let source;
    try {
      if (currency === "HKD") {
        rate = 1;
        source = "HKD";
      } else if (rowId && categoryRateByRowId.has(String(rowId))) {
        rate = categoryRateByRowId.get(String(rowId));
        source = "category tab";
        fromCategory++;
      } else {
        rate = await fxRates.getRateToHkd(currency, date);
        source = "fresh FX lookup";
        fromFreshLookup++;
      }
    } catch (err) {
      errors.push(`Row ${i + 1} (${currency} ${date}): FX lookup failed — ${err.message}`);
      continue;
    }

    if (APPLY) {
      const updated = [...row];
      updated[rateCol] = rate;
      try {
        await updateTableRowByIndex(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, RAW_SHEET_NAME, i, updated);
      } catch (err) {
        errors.push(`Row ${i + 1}: write failed — ${err.message}`);
        continue;
      }
    }
    filled++;
    console.log(`Row ${i + 1}: ${currency} ${date} → ${rate} (${source})`);
  }

  console.log(
    `\n${APPLY ? "Wrote" : "Would write"} ${filled} row(s) — ${fromCategory} from a matching category tab, ${fromFreshLookup} via fresh historical FX lookup, ${filled - fromCategory - fromFreshLookup} HKD (rate 1). Skipped ${skippedAlreadySet} row(s) that already had a value.`
  );
  if (errors.length) {
    console.log(`\n⚠️ ${errors.length} error(s):`);
    errors.forEach((e) => console.log(`- ${e}`));
  }
  if (!APPLY) {
    console.log("\nDry run only — nothing was written. Re-run with --apply once this looks right.");
  }
})();
