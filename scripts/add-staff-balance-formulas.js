require("dotenv").config();
const {
  getTableHeaders,
  setRangeValues,
  setRangeFormulas,
  getRangeFormulas,
  getRangeNumberFormat,
  setRangeNumberFormat,
} = require("../src/graphExcel");

// One-off setup (Mariya, 2026-09-15): adds a live, self-updating summary block to the right of the
// StaffAdvances table on the "Staff Advances" worksheet, so the remaining balance for each helper
// (Marietta, Yaya) is visible directly in Excel without needing the bot's /staff command.
//
// Columns Q:V, row 1 = headers, rows 2-3 = Marietta / Yaya:
//   Q Staff | R Total Advances (HKD) | S Total Spent (HKD) | T Remaining (HKD)
//   | U Last Advance Amount (HKD) | V Last Advance Date
//
// R/S/T are plain SUMIF nets — correct no matter how many advance or spending rows exist per
// person (Mariya: "there will be not just one receipt per advance - but many").
//
// U/V implement "treat each dated entry as its own advance" (Mariya, 2026-09-15): V finds the
// latest date among that person's StaffAdvances rows (MAXIFS — works directly on the Дата column
// since it stores real Excel date serials, confirmed via scripts/fix-staff-advances-dates.js), and
// U sums every StaffAdvances row for that person dated exactly that day (so several receipts for
// the same advance, logged the same day, are combined into one "advance" amount — but two
// genuinely different dates are two separate advances). This exactly mirrors the logic in
// getStaffBalancesReport() in src/index.js (used by the bot's /staff command), so the bot and this
// sheet always agree.
//
// Run:
//   node scripts/add-staff-balance-formulas.js            (dry run — shows the formulas that would be written)
//   node scripts/add-staff-balance-formulas.js --apply     (actually writes them)

const ADVANCES_TABLE = "StaffAdvances";
const EXPENSES_TABLE = "StaffExpenses";
const WORKSHEET = "Staff Advances";
const STAFF_MEMBERS = ["Marietta", "Yaya"];

const HEADERS = [
  "Staff",
  "Total Advances (HKD)",
  "Total Spent (HKD)",
  "Remaining (HKD)",
  "Last Advance Amount (HKD)",
  "Last Advance Date",
];

const APPLY = process.argv.includes("--apply");

function staffRowFormulas(staffCell) {
  return [
    `=SUMIF(${ADVANCES_TABLE}[Получатель/Сотрудник],${staffCell},${ADVANCES_TABLE}[Сумма (HKD)])`, // R: Total Advances
    `=SUMIF(${EXPENSES_TABLE}[Получатель/Сотрудник],${staffCell},${EXPENSES_TABLE}[Сумма (HKD)])`, // S: Total Spent
    `=R${staffCell.slice(1)}-S${staffCell.slice(1)}`, // T: Remaining = Total Advances - Total Spent
    `=SUMIFS(${ADVANCES_TABLE}[Сумма (HKD)],${ADVANCES_TABLE}[Получатель/Сотрудник],${staffCell},${ADVANCES_TABLE}[Дата],V${staffCell.slice(
      1
    )})`, // U: Last Advance Amount — every row dated exactly the max date found in V
    `=MAXIFS(${ADVANCES_TABLE}[Дата],${ADVANCES_TABLE}[Получатель/Сотрудник],${staffCell})`, // V: Last Advance Date
  ];
}

(async () => {
  const { EXCEL_DRIVE_ID, EXCEL_ITEM_ID } = process.env;
  if (!EXCEL_DRIVE_ID || !EXCEL_ITEM_ID) {
    console.error("Set EXCEL_DRIVE_ID and EXCEL_ITEM_ID in .env first.");
    process.exit(1);
  }

  console.log(`\n=== Add Q:V staff balance formulas to "${WORKSHEET}" — ${APPLY ? "APPLY" : "DRY RUN"} ===\n`);

  // Sanity-check the column headers this script hardcodes actually exist on both tables, so a
  // renamed column fails loudly here instead of silently producing #NAME?/#VALUE! formulas.
  const advancesHeaders = await getTableHeaders(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, ADVANCES_TABLE);
  const expensesHeaders = await getTableHeaders(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, EXPENSES_TABLE);
  for (const col of ["Получатель/Сотрудник", "Сумма (HKD)", "Дата"]) {
    if (!advancesHeaders.includes(col)) {
      console.error(`Column "${col}" not found on ${ADVANCES_TABLE}. Found: ${advancesHeaders.join(", ")}`);
      process.exit(1);
    }
  }
  if (!expensesHeaders.includes("Получатель/Сотрудник") || !expensesHeaders.includes("Сумма (HKD)")) {
    console.error(`Expected columns not found on ${EXPENSES_TABLE}. Found: ${expensesHeaders.join(", ")}`);
    process.exit(1);
  }
  console.log("Column headers verified on both tables.\n");

  // MAXIFS on the Дата column returns a bare Excel date serial (e.g. 46122), not a formatted date —
  // same underlying issue scripts/fix-staff-advances-dates.js fixed for column A itself. Reuse
  // whatever numberFormat is already applied to StaffAdvances' own Дата column (A2, a known-good
  // reference since that script already fixed it) so V2:V3 display as dates instead of serials.
  const dateFormatRef = await getRangeNumberFormat(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, WORKSHEET, "A2:A2");
  const dateNumberFormat = dateFormatRef.numberFormat[0][0];
  console.log(`Date numberFormat to apply to V2:V3 (from ${WORKSHEET}!A2):`, JSON.stringify(dateNumberFormat));

  const headerRow = [HEADERS];
  console.log(`Q1:V1 headers:`, JSON.stringify(HEADERS));

  const staffValues = STAFF_MEMBERS.map((name) => [name]); // Q2, Q3 — plain values, not formulas
  const rowFormulas = STAFF_MEMBERS.map((name, i) => staffRowFormulas(`Q${i + 2}`));

  STAFF_MEMBERS.forEach((name, i) => {
    console.log(`Row ${i + 2} (${name}): R:V =`, JSON.stringify(rowFormulas[i]));
  });

  if (!APPLY) {
    console.log("\nDry run only — nothing was changed. Re-run with --apply once this looks right.");
    return;
  }

  await setRangeValues(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, WORKSHEET, "Q1:V1", headerRow);
  await setRangeValues(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, WORKSHEET, "Q2:Q3", staffValues);
  await setRangeFormulas(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, WORKSHEET, "R2:V2", [rowFormulas[0]]);
  await setRangeFormulas(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, WORKSHEET, "R3:V3", [rowFormulas[1]]);
  await setRangeNumberFormat(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, WORKSHEET, "V2:V3", [
    [dateNumberFormat],
    [dateNumberFormat],
  ]);

  console.log(`\n✅ Wrote headers + formulas to Q1:V3 on "${WORKSHEET}" (V2:V3 formatted as dates).`);

  const after = await getRangeFormulas(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, WORKSHEET, "Q1:V3");
  console.log("Resulting Q1:V3 formulas/values:", JSON.stringify(after.formulas));
})().catch((err) => {
  if (err.response) {
    console.error(`Graph error ${err.response.status}:`, JSON.stringify(err.response.data, null, 2));
  } else {
    console.error(err.message);
  }
  process.exit(1);
});
