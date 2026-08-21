require("dotenv").config();
const { getTableHeaders, addTableColumn } = require("../src/graphExcel");
const CATEGORY_TABLE_MAP = require("./lib/categoryTableMap");

// One-time migration: adds the "RowID" and "SyncHash" technical columns to RAW and every
// expense category Excel Table the bot actually writes to. index.js's
// buildRawRowValues/buildCategoryRowValues have already been writing rows with RowID/SyncHash
// appended at the end (see src/syncJob.js's RAW_COLS/STANDARD_CAT_COLS/INSURANCE_CAT_COLS) since
// the RAW<->category sync job shipped — but the live workbook's tables were never actually
// widened to match, because this script didn't exist yet. Any write to a table still stuck at
// its old width fails with Graph's "InvalidArgument: number of rows or columns in the input
// array doesn't match the size or dimensions of the range" (reproduced 2026-08-20 logging a
// real expense against RAW).
//
// Per Mariya (2026-08-20):
//   - Credit Cards is manual-only — the bot never logs to it, so it's excluded from this
//     script entirely (and from CATEGORY_TABLE_MAP's effective set — see MANUAL_ONLY_CATEGORIES
//     in src/index.js). Left completely untouched.
//   - Health and MedInsurance (Insurance) keep their 5 insurance columns, so their
//     pre-migration width is 17, not 12 — RowID/SyncHash get appended on top, landing at
//     19 columns total. Matches src/syncJob.js's INSURANCE_CAT_COLS and
//     src/index.js's INSURANCE_CATEGORY_TABLES branch in buildCategoryRowValues.
//   - Every other category table (including Personal Spending / StaffExpenses) is expected at
//     the standard pre-migration width of 12 (5 insurance columns already removed, Пользователь
//     moved to the end — the cleanup Mariya did by hand to 12 tables on 2026-08-19). If
//     StaffExpenses still shows up at 17 columns, it hasn't had that cleanup done yet — this
//     script deliberately will NOT touch it (no safe automated column-delete/reorder exists in
//     graphExcel.js), and instead calls it out by name so it isn't silently skipped/confused
//     with a genuinely-unexpected table.
//
// Safety: for each table, only proceeds if the CURRENT header count matches that table's
// expected pre-migration width exactly. If a table already has RowID/SyncHash at the end, it's
// skipped (already migrated — safe to re-run this script). Any other column count is left
// untouched and reported, rather than guessed at — Graph's addTableColumn always appends at
// the end, so this only produces the right result when the starting width is exactly right.
//
// Run:
//   node scripts/migrate-workbook.js            (dry run — reports what would change)
//   node scripts/migrate-workbook.js --apply     (adds the columns)

const RAW_SHEET = (process.env.RAW_SHEET_NAME || "RAW").trim();
const APPLY = process.argv.includes("--apply");
const NEW_COLUMNS = ["RowID", "SyncHash"];
const STANDARD_PRE_MIGRATION_WIDTH = 12;
const INSURANCE_PRE_MIGRATION_WIDTH = 17;

// Tables the bot never writes to — excluded from this script (and from the sync job) entirely.
// Keep in sync with MANUAL_ONLY_CATEGORIES in src/index.js.
//
// CreditCards was removed from this list on 2026-08-20 (Mariya wants Credit Cards loggable via
// Telegram again). MANUAL_ONLY_CATEGORIES in src/index.js has been flipped to [] — Credit Cards
// is a normal pickable/loggable category like any other.
const MANUAL_ONLY_TABLES = [];

// Tables that keep their 5 insurance columns instead of getting the 2026-08-19 cleanup.
// Keep in sync with INSURANCE_CATEGORY_TABLES in src/index.js and src/syncJob.js.
//
// CreditCards was briefly here too (2026-08-20, when it was column-for-column identical to
// Health/MedInsurance's 17-column layout and migrated 17 -> 19 the same way). It already has
// RowID/SyncHash, so this one-time migration script has nothing left to do for it either way —
// but it no longer belongs in this list: on 2026-08-21 Mariya deleted 4 of its 5 insurance
// columns by hand, keeping only Карта. Its current 15-column shape lives in index.js's
// CREDITCARDS_TABLE branch and syncJob.js's CREDITCARDS_CAT_COLS, not here.
const INSURANCE_TABLES = ["Health", "MedInsurance"];

// Every table the bot writes RowID/SyncHash into: RAW plus every expense category table except
// the manual-only ones. CATEGORY_TABLE_MAP also carries a non-table property,
// REQUIRED_IN_MONTHLY_TOTALS (an array, used by the Summary-formula scripts) — must filter to
// string values only, or Object.values() picks that array up as a bogus "table name" too
// (caught by a dry-run sanity check before this script ever touched the live workbook).
const ALL_TABLES = [
  RAW_SHEET,
  ...new Set(
    Object.values(CATEGORY_TABLE_MAP)
      .filter((table) => typeof table === "string")
      .filter((table) => !MANUAL_ONLY_TABLES.includes(table))
  ),
];

function expectedPreMigrationWidth(tableName) {
  if (tableName === RAW_SHEET) return STANDARD_PRE_MIGRATION_WIDTH;
  return INSURANCE_TABLES.includes(tableName)
    ? INSURANCE_PRE_MIGRATION_WIDTH
    : STANDARD_PRE_MIGRATION_WIDTH;
}

(async () => {
  const { EXCEL_DRIVE_ID, EXCEL_ITEM_ID } = process.env;
  if (!EXCEL_DRIVE_ID || !EXCEL_ITEM_ID) {
    console.error("Set EXCEL_DRIVE_ID and EXCEL_ITEM_ID in .env first.");
    process.exit(1);
  }

  console.log(`\n=== Workbook migration: add RowID/SyncHash columns — ${APPLY ? "APPLY" : "DRY RUN"} ===\n`);
  console.log(`Skipping manual-only table(s), never touched by this script: ${MANUAL_ONLY_TABLES.join(", ")}\n`);

  const toMigrate = [];
  const alreadyDone = [];
  const unexpected = [];
  const needsManualCleanupFirst = [];

  for (const tableName of ALL_TABLES) {
    try {
      const headers = await getTableHeaders(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, tableName);
      const hasBoth = NEW_COLUMNS.every((c) => headers.includes(c));
      const expectedWidth = expectedPreMigrationWidth(tableName);

      if (hasBoth) {
        alreadyDone.push({ tableName, headers });
      } else if (headers.length === expectedWidth) {
        toMigrate.push({ tableName, headers, expectedWidth });
      } else if (
        !INSURANCE_TABLES.includes(tableName) &&
        headers.length === INSURANCE_PRE_MIGRATION_WIDTH
      ) {
        // Standard table still at the old 17-column insurance layout — the 2026-08-19
        // cleanup (remove 5 insurance columns, move Пользователь to the end) hasn't been done
        // for it yet. Flag by name rather than lumping in with genuinely unexpected widths.
        needsManualCleanupFirst.push({ tableName, headers });
      } else {
        unexpected.push({ tableName, headers, expectedWidth });
      }
    } catch (err) {
      console.error(`Could not read headers for "${tableName}": ${err.message}`);
      process.exit(1);
    }
  }

  if (alreadyDone.length) {
    console.log(`Already migrated (${alreadyDone.length}) — skipped:`);
    for (const { tableName } of alreadyDone) console.log(`  ✓ ${tableName}`);
    console.log("");
  }

  if (needsManualCleanupFirst.length) {
    console.log(
      `🛠  ${needsManualCleanupFirst.length} table(s) still have the old 17-column insurance layout and need the manual cleanup (remove Страховка/Статус выплаты/Страховая компания/Период покрытия/Карта, move Пользователь to the end) — same as the other 12 tables got on 2026-08-19 — before this script can add RowID/SyncHash:`
    );
    for (const { tableName, headers } of needsManualCleanupFirst) {
      console.log(`  ! ${tableName} — ${headers.length} columns: ${headers.join(", ")}`);
    }
    console.log("");
  }

  if (unexpected.length) {
    console.log(
      `⚠️ ${unexpected.length} table(s) have neither their expected pre-migration width nor RowID/SyncHash already — left untouched, check by hand:`
    );
    for (const { tableName, headers, expectedWidth } of unexpected) {
      console.log(`  ? ${tableName} — expected ${expectedWidth}, has ${headers.length} columns: ${headers.join(", ")}`);
    }
    console.log("");
  }

  if (!toMigrate.length) {
    console.log("Nothing to migrate.");
    return;
  }

  console.log(`To migrate (${toMigrate.length}):`);
  for (const { tableName, expectedWidth } of toMigrate) {
    console.log(`  + ${tableName} (${expectedWidth} → ${expectedWidth + NEW_COLUMNS.length}) — will append columns: ${NEW_COLUMNS.join(", ")}`);
  }
  console.log("");

  if (!APPLY) {
    console.log("Dry run only — nothing was written. Re-run with --apply once this looks right.");
    return;
  }

  for (const { tableName } of toMigrate) {
    for (const columnName of NEW_COLUMNS) {
      await addTableColumn(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, tableName, columnName);
      console.log(`  ✅ ${tableName} — added "${columnName}"`);
    }
  }

  console.log("\nDone.");
})().catch((err) => {
  if (err.response) {
    console.error(`Graph error ${err.response.status}:`, JSON.stringify(err.response.data, null, 2));
  } else {
    console.error(err.message);
  }
  process.exit(1);
});
