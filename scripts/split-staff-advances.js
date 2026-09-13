require("dotenv").config();
const {
  getTableHeaders,
  getTableRows,
  getTableRangeAddress,
  appendTableRow,
  deleteTable,
  clearRangeContents,
  listWorksheets,
  addWorksheet,
  renameWorksheet,
  setRangeValues,
  createTable,
  renameTable,
  listWorkbookTables,
} = require("../src/graphExcel");

// One-time migration (Mariya, 2026-09-13): the "Personal Spending" tab currently hosts TWO
// separate Excel Tables stacked on top of each other — StaffAdvances (A5:L25, 21 rows, 12
// columns, no RowID/SyncHash yet) and StaffExpenses (A34:N91, 58 rows, 14 columns) — confirmed
// live via `node scripts/inspect-staff-tables.js`. They're already correctly split data-wise;
// they just need to live on separate tabs instead of one confusing shared tab.
//
// This script:
//   1. Creates a brand-new "Staff Advances" worksheet.
//   2. Re-creates the StaffAdvances table there with its existing 12 columns and copies all its
//      existing rows over, in order.
//   3. Deletes the old StaffAdvances table object from "Personal Spending" and clears the now-
//      orphaned cells it used to occupy (deleting a Graph table un-tableifies the range but
//      leaves the raw cell values sitting there otherwise).
//   4. Renames the "Personal Spending" tab to "Staff Spending" (it now only holds StaffExpenses).
//
// RowID/SyncHash are deliberately NOT added here — run `npm run migrate-workbook -- --apply`
// afterward, which already knows how to add those two columns to any category table sitting at
// the standard pre-migration width (12), and will pick up the new StaffAdvances table
// automatically since it's now listed in CATEGORY_TABLE_MAP.
//
// Run:
//   node scripts/split-staff-advances.js            (dry run — reports what would happen)
//   node scripts/split-staff-advances.js --apply     (actually moves things)

const OLD_TABLE = "StaffAdvances";
const OLD_WORKSHEET = "Personal Spending";
const NEW_WORKSHEET = "Staff Advances";
const RENAMED_OLD_WORKSHEET = "Staff Spending";
const TEMP_TABLE_NAME_PREFIX = "StaffAdvancesTmp";

const APPLY = process.argv.includes("--apply");

(async () => {
  const { EXCEL_DRIVE_ID, EXCEL_ITEM_ID } = process.env;
  if (!EXCEL_DRIVE_ID || !EXCEL_ITEM_ID) {
    console.error("Set EXCEL_DRIVE_ID and EXCEL_ITEM_ID in .env first.");
    process.exit(1);
  }

  console.log(`\n=== Move ${OLD_TABLE} onto its own tab — ${APPLY ? "APPLY" : "DRY RUN"} ===\n`);

  const headers = await getTableHeaders(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, OLD_TABLE);
  const rows = await getTableRows(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, OLD_TABLE);
  const fullRangeAddress = await getTableRangeAddress(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, OLD_TABLE);

  console.log(`Found "${OLD_TABLE}" on "${OLD_WORKSHEET}" — range ${fullRangeAddress}.`);
  console.log(`Columns (${headers.length}): ${headers.join(", ")}`);
  console.log(`Data rows: ${rows.length}`);
  if (rows.length) {
    console.log("First row:", rows[0]);
    console.log("Last row:", rows[rows.length - 1]);
  }
  console.log("");
  console.log(`Plan:`);
  console.log(`  1. Create worksheet "${NEW_WORKSHEET}"`);
  console.log(`  2. Re-create "${OLD_TABLE}" there with the same ${headers.length} columns + ${rows.length} rows`);
  console.log(`  3. Delete the old table + clear ${fullRangeAddress} on "${OLD_WORKSHEET}"`);
  console.log(`  4. Rename "${OLD_WORKSHEET}" -> "${RENAMED_OLD_WORKSHEET}"`);
  console.log(`  (RowID/SyncHash columns NOT added here — run "npm run migrate-workbook -- --apply" after this)\n`);

  if (!APPLY) {
    console.log("Dry run only — nothing was written. Re-run with --apply once this looks right.");
    return;
  }

  // Idempotency: an earlier run may have already created the worksheet and/or a temp table
  // before failing partway through (e.g. the createTable address bug fixed 2026-09-13). Check
  // what's already there instead of blindly re-creating and duplicating worksheets/tables.
  const existingWorksheets = await listWorksheets(EXCEL_DRIVE_ID, EXCEL_ITEM_ID);
  if (existingWorksheets.includes(NEW_WORKSHEET)) {
    console.log(`Worksheet "${NEW_WORKSHEET}" already exists — reusing it.`);
  } else {
    console.log(`Creating worksheet "${NEW_WORKSHEET}"...`);
    await addWorksheet(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, NEW_WORKSHEET);
  }

  const existingTables = await listWorkbookTables(EXCEL_DRIVE_ID, EXCEL_ITEM_ID);
  const existingTemp = existingTables.find(
    (t) => t.worksheet === NEW_WORKSHEET && t.name.startsWith(TEMP_TABLE_NAME_PREFIX)
  );

  let tempName;
  if (existingTemp) {
    tempName = existingTemp.name;
    console.log(`Temp table "${tempName}" already exists on "${NEW_WORKSHEET}" from a previous run — reusing it, skipping row copy to avoid duplicates.`);
  } else {
    const lastColLetter = String.fromCharCode("A".charCodeAt(0) + headers.length - 1);
    const headerAddress = `A1:${lastColLetter}1`;
    console.log(`Writing header row to ${NEW_WORKSHEET}!${headerAddress}...`);
    await setRangeValues(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, NEW_WORKSHEET, headerAddress, [headers]);

    console.log(`Creating table from ${headerAddress}...`);
    const createdTable = await createTable(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, NEW_WORKSHEET, headerAddress, true);
    tempName = `${TEMP_TABLE_NAME_PREFIX}_${Date.now()}`;
    console.log(`Renaming table "${createdTable.name}" -> "${tempName}" (temporary, old "${OLD_TABLE}" still exists)...`);
    await renameTable(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, createdTable.name, tempName);

    console.log(`\nCopying ${rows.length} row(s) into the new table...`);
    for (let i = 0; i < rows.length; i++) {
      await appendTableRow(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, tempName, rows[i]);
      console.log(`  ✅ copied row ${i + 1}/${rows.length}`);
    }
  }

  console.log(`\nDeleting old table "${OLD_TABLE}" on "${OLD_WORKSHEET}"...`);
  await deleteTable(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, OLD_TABLE);

  console.log(`Clearing leftover cells ${fullRangeAddress} on "${OLD_WORKSHEET}"...`);
  const rangeOnly = fullRangeAddress.includes("!") ? fullRangeAddress.split("!")[1] : fullRangeAddress;
  await clearRangeContents(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, OLD_WORKSHEET, rangeOnly);

  console.log(`\nRenaming table "${tempName}" -> "${OLD_TABLE}"...`);
  await renameTable(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, tempName, OLD_TABLE);

  console.log(`\nRenaming worksheet "${OLD_WORKSHEET}" -> "${RENAMED_OLD_WORKSHEET}"...`);
  await renameWorksheet(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, OLD_WORKSHEET, RENAMED_OLD_WORKSHEET);

  console.log(
    `\nDone. "${OLD_TABLE}" now lives alone on "${NEW_WORKSHEET}"; "${RENAMED_OLD_WORKSHEET}" now holds only StaffExpenses.`
  );
  console.log(`Next: run "npm run migrate-workbook" (dry run), then "npm run migrate-workbook -- --apply" to add RowID/SyncHash to ${OLD_TABLE}.`);
})().catch((err) => {
  if (err.response) {
    console.error(`Graph error ${err.response.status}:`, JSON.stringify(err.response.data, null, 2));
  } else {
    console.error(err.message);
  }
  process.exit(1);
});
