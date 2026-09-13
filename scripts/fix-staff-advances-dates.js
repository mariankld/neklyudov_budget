require("dotenv").config();
const {
  getTableDataBodyRangeAddress,
  getRangeNumberFormat,
  setRangeNumberFormat,
} = require("../src/graphExcel");

// One-off fix (Mariya, 2026-09-13): after scripts/split-staff-advances.js moved the StaffAdvances
// table onto its own "Staff Advances" worksheet, the Дата (date) column started showing raw Excel
// serial numbers (46122, 46132, ...) instead of formatted dates. Root cause: appendTableRow only
// writes raw cell values — it never copied over the Дата column's date numberFormat, so the new
// worksheet's column A has no date format applied and Excel just prints the underlying serial
// number.
//
// This script reads the Дата column's numberFormat from the still-correctly-formatted
// "Staff Spending" tab (StaffExpenses table) as the reference, then applies that same format to
// the Дата column of the "Staff Advances" tab (StaffAdvances table).
//
// Run:
//   node scripts/fix-staff-advances-dates.js            (dry run — shows current vs reference format)
//   node scripts/fix-staff-advances-dates.js --apply     (actually applies the fix)

const REFERENCE_TABLE = "StaffExpenses"; // lives on "Staff Spending", still displays dates correctly
const REFERENCE_WORKSHEET = "Staff Spending";
const BROKEN_TABLE = "StaffAdvances";
const BROKEN_WORKSHEET = "Staff Advances";
const DATE_COLUMN_LETTER = "A"; // Дата is the first column on both tables

const APPLY = process.argv.includes("--apply");

function firstColumnRange(dataBodyRangeAddress, columnLetter) {
  // dataBodyRangeAddress looks like "Staff Advances!A6:L26" — we only want the Дата column's rows.
  const rangePart = dataBodyRangeAddress.includes("!")
    ? dataBodyRangeAddress.split("!")[1]
    : dataBodyRangeAddress;
  const [start, end] = rangePart.split(":");
  const startRow = start.match(/\d+/)[0];
  const endRow = end.match(/\d+/)[0];
  return `${columnLetter}${startRow}:${columnLetter}${endRow}`;
}

(async () => {
  const { EXCEL_DRIVE_ID, EXCEL_ITEM_ID } = process.env;
  if (!EXCEL_DRIVE_ID || !EXCEL_ITEM_ID) {
    console.error("Set EXCEL_DRIVE_ID and EXCEL_ITEM_ID in .env first.");
    process.exit(1);
  }

  console.log(`\n=== Fix Дата column formatting on "${BROKEN_WORKSHEET}" — ${APPLY ? "APPLY" : "DRY RUN"} ===\n`);

  const refBodyAddress = await getTableDataBodyRangeAddress(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, REFERENCE_TABLE);
  const refDateRange = firstColumnRange(refBodyAddress, DATE_COLUMN_LETTER);
  console.log(`Reference: "${REFERENCE_TABLE}" on "${REFERENCE_WORKSHEET}", Дата column range ${refDateRange}`);
  const refFormat = await getRangeNumberFormat(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, REFERENCE_WORKSHEET, refDateRange);
  console.log("Reference numberFormat (first few rows):", JSON.stringify(refFormat.numberFormat.slice(0, 3)));
  console.log("Reference values (first few rows):", JSON.stringify(refFormat.values.slice(0, 3)));

  const brokenBodyAddress = await getTableDataBodyRangeAddress(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, BROKEN_TABLE);
  const brokenDateRange = firstColumnRange(brokenBodyAddress, DATE_COLUMN_LETTER);
  console.log(`\nBroken: "${BROKEN_TABLE}" on "${BROKEN_WORKSHEET}", Дата column range ${brokenDateRange}`);
  const brokenFormat = await getRangeNumberFormat(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, BROKEN_WORKSHEET, brokenDateRange);
  console.log("Current numberFormat (first few rows):", JSON.stringify(brokenFormat.numberFormat.slice(0, 3)));
  console.log("Current values (first few rows):", JSON.stringify(brokenFormat.values.slice(0, 3)));

  const rowCount = brokenFormat.numberFormat.length;
  // Reference format should be uniform down the column — grab row 0's format and replicate it for
  // however many rows the broken table has (row counts differ between the two tables).
  const targetFormatCode = refFormat.numberFormat[0][0];
  const newNumberFormat = Array.from({ length: rowCount }, () => [targetFormatCode]);

  console.log(`\nWill apply numberFormat "${targetFormatCode}" to all ${rowCount} row(s) of ${brokenDateRange} on "${BROKEN_WORKSHEET}".`);

  if (!APPLY) {
    console.log("\nDry run only — nothing was changed. Re-run with --apply once this looks right.");
    return;
  }

  await setRangeNumberFormat(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, BROKEN_WORKSHEET, brokenDateRange, newNumberFormat);
  console.log(`\n✅ Applied date format to ${brokenDateRange} on "${BROKEN_WORKSHEET}".`);

  const after = await getRangeNumberFormat(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, BROKEN_WORKSHEET, brokenDateRange);
  console.log("New numberFormat (first few rows):", JSON.stringify(after.numberFormat.slice(0, 3)));
})().catch((err) => {
  if (err.response) {
    console.error(`Graph error ${err.response.status}:`, JSON.stringify(err.response.data, null, 2));
  } else {
    console.error(err.message);
  }
  process.exit(1);
});
