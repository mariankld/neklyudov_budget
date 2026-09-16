require("dotenv").config();
const {
  getTableHeaders,
  setRangeValues,
  setRangeFormulas,
  getRangeFormulas,
  getRangeNumberFormat,
  setRangeNumberFormat,
} = require("../src/graphExcel");

// One-off setup (Mariya, 2026-09-15; revised 2026-09-16), adds a live, self-updating summary block
// to the right of the StaffAdvances table on the "Staff Advances" worksheet, so each helper's
// balance is visible directly in Excel without needing the bot's /staff command.
//
// Columns Q:Y, row 1 = headers, rows 2-3 = Marietta / Yaya:
//   Q Staff | R Currency
//   | S Total Advances (Lifetime) | T Total Spent (Lifetime) | U Remaining (Lifetime)
//   | V Last Advance Amount | W Last Advance Date
//   | X Spent Since Last Advance | Y Remaining (Since Last Advance)
//
// Mariya wants BOTH numbers visible side by side (2026-09-16): S/T/U are the old-style lifetime
// running totals ("every advance ever minus every expense ever"), while V/W/X/Y are the
// period balance scoped to just the most recent advance — which is what actually answers "what's
// left from what I just gave them." X/Y are new in this revision; S/T/U/V/W existed before but are
// now computed per-currency (see below) instead of always in HKD.
//
// "Most recent advance" (Mariya, 2026-09-15, "treat each dated entry as its own advance"): every
// distinct date on which a person has StaffAdvances rows is one advance event; several receipts
// for the same advance logged the same day are summed into one event. The event with the latest
// date is "the most recent advance." W finds that date (MAXIFS on the Дата column, which stores
// real Excel date serials — confirmed via scripts/fix-staff-advances-dates.js) and V sums every
// row dated exactly that day. This exactly mirrors getStaffBalancesReport() in src/index.js (the
// /staff command), so the bot and this sheet always agree.
//
// X (Spent Since Last Advance) sums every StaffExpenses row for that person dated on/after W. If
// there's no advance on record yet, W/V/X/Y all show blank rather than a bogus number — an
// unguarded SUMIFS with "Дата >= 0" would otherwise match every historical row.
//
// Currency (Mariya, 2026-09-16): Marietta is paid and tracked in HKD, so her row (2) uses
// Сумма (HKD) throughout, no filtering. Yaya is always paid and spends in THB, so her row (3) uses
// the raw Сумма column restricted to rows tagged Валюта=THB — no FX conversion, and any of her rows
// not tagged THB are simply excluded from her totals (they won't match the Валюта=THB criteria).
//
// Run:
//   node scripts/add-staff-balance-formulas.js            (dry run — shows the formulas that would be written)
//   node scripts/add-staff-balance-formulas.js --apply     (actually writes them)

const ADVANCES_TABLE = "StaffAdvances";
const EXPENSES_TABLE = "StaffExpenses";
const WORKSHEET = "Staff Advances";

// Per-person currency: Marietta stays in HKD (Сумма (HKD), no filtering); Yaya is THB-only (raw
// Сумма, filtered to Валюта=THB) — see the file header comment for why.
const STAFF_MEMBERS = [
  { name: "Marietta", currency: "HKD" },
  { name: "Yaya", currency: "THB" },
];

const HEADERS = [
  "Staff",
  "Currency",
  "Total Advances (Lifetime)",
  "Total Spent (Lifetime)",
  "Remaining (Lifetime)",
  "Last Advance Amount",
  "Last Advance Date",
  "Spent Since Last Advance",
  "Remaining (Since Last Advance)",
];

const APPLY = process.argv.includes("--apply");

/**
 * Builds the S:Y formulas for one staff row. `staffCell` is the Q-column cell holding this
 * person's name (e.g. "Q2"); `currency` is "HKD" or "THB".
 *
 * For HKD, every SUM*IFS reads the frozen Сумма (HKD) column with no currency filter. For THB, the
 * same formulas read the raw Сумма column and add a Валюта="THB" criteria to every SUM*IFS/COUNTIFS
 * call, so non-THB rows for that person are simply excluded (mirrors the bot's skip-and-flag logic).
 */
function staffRowFormulas(staffCell, currency) {
  const row = staffCell.slice(1);
  const isThb = currency === "THB";

  const advAmountCol = isThb ? `${ADVANCES_TABLE}[Сумма]` : `${ADVANCES_TABLE}[Сумма (HKD)]`;
  const expAmountCol = isThb ? `${EXPENSES_TABLE}[Сумма]` : `${EXPENSES_TABLE}[Сумма (HKD)]`;
  const advCurrencyCriteria = isThb ? `,${ADVANCES_TABLE}[Валюта],"THB"` : "";
  const expCurrencyCriteria = isThb ? `,${EXPENSES_TABLE}[Валюта],"THB"` : "";

  const totalAdvances = isThb
    ? `=SUMIFS(${advAmountCol},${ADVANCES_TABLE}[Получатель/Сотрудник],${staffCell}${advCurrencyCriteria})`
    : `=SUMIF(${ADVANCES_TABLE}[Получатель/Сотрудник],${staffCell},${advAmountCol})`;

  const totalSpent = isThb
    ? `=SUMIFS(${expAmountCol},${EXPENSES_TABLE}[Получатель/Сотрудник],${staffCell}${expCurrencyCriteria})`
    : `=SUMIF(${EXPENSES_TABLE}[Получатель/Сотрудник],${staffCell},${expAmountCol})`;

  const remainingLifetime = `=S${row}-T${row}`; // U: lifetime Remaining = Total Advances - Total Spent

  // W: Last Advance Date — blank if this person has no (currency-matching) advance rows at all.
  const lastAdvanceDate = `=IF(COUNTIFS(${ADVANCES_TABLE}[Получатель/Сотрудник],${staffCell}${advCurrencyCriteria})=0,"",MAXIFS(${ADVANCES_TABLE}[Дата],${ADVANCES_TABLE}[Получатель/Сотрудник],${staffCell}${advCurrencyCriteria}))`;

  // V: Last Advance Amount — every row dated exactly the max date found in W; blank if W is blank.
  const lastAdvanceAmount = `=IF(W${row}="","",SUMIFS(${advAmountCol},${ADVANCES_TABLE}[Получатель/Сотрудник],${staffCell}${advCurrencyCriteria},${ADVANCES_TABLE}[Дата],W${row}))`;

  // X: Spent Since Last Advance — every StaffExpenses row dated on/after W; blank if W is blank.
  const spentSince = `=IF(W${row}="","",SUMIFS(${expAmountCol},${EXPENSES_TABLE}[Получатель/Сотрудник],${staffCell}${expCurrencyCriteria},${EXPENSES_TABLE}[Дата],">="&W${row}))`;

  // Y: period Remaining = Last Advance Amount - Spent Since Last Advance; blank if either is blank.
  const remainingSince = `=IF(OR(V${row}="",X${row}=""),"",V${row}-X${row})`;

  return [totalAdvances, totalSpent, remainingLifetime, lastAdvanceAmount, lastAdvanceDate, spentSince, remainingSince];
}

(async () => {
  const { EXCEL_DRIVE_ID, EXCEL_ITEM_ID } = process.env;
  if (!EXCEL_DRIVE_ID || !EXCEL_ITEM_ID) {
    console.error("Set EXCEL_DRIVE_ID and EXCEL_ITEM_ID in .env first.");
    process.exit(1);
  }

  console.log(`\n=== Add Q:Y staff balance formulas to "${WORKSHEET}" — ${APPLY ? "APPLY" : "DRY RUN"} ===\n`);

  // Sanity-check the column headers this script hardcodes actually exist on both tables, so a
  // renamed column fails loudly here instead of silently producing #NAME?/#VALUE! formulas.
  const advancesHeaders = await getTableHeaders(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, ADVANCES_TABLE);
  const expensesHeaders = await getTableHeaders(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, EXPENSES_TABLE);
  for (const col of ["Получатель/Сотрудник", "Сумма", "Сумма (HKD)", "Валюта", "Дата"]) {
    if (!advancesHeaders.includes(col)) {
      console.error(`Column "${col}" not found on ${ADVANCES_TABLE}. Found: ${advancesHeaders.join(", ")}`);
      process.exit(1);
    }
  }
  for (const col of ["Получатель/Сотрудник", "Сумма", "Сумма (HKD)", "Валюта"]) {
    if (!expensesHeaders.includes(col)) {
      console.error(`Column "${col}" not found on ${EXPENSES_TABLE}. Found: ${expensesHeaders.join(", ")}`);
      process.exit(1);
    }
  }
  console.log("Column headers verified on both tables.\n");

  // MAXIFS on the Дата column returns a bare Excel date serial (e.g. 46122), not a formatted date —
  // same underlying issue scripts/fix-staff-advances-dates.js fixed for column A itself. Reuse
  // whatever numberFormat is already applied to StaffAdvances' own Дата column (A2, a known-good
  // reference since that script already fixed it) so W2:W3 display as dates instead of serials.
  const dateFormatRef = await getRangeNumberFormat(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, WORKSHEET, "A2:A2");
  const dateNumberFormat = dateFormatRef.numberFormat[0][0];
  console.log(`Date numberFormat to apply to W2:W3 (from ${WORKSHEET}!A2):`, JSON.stringify(dateNumberFormat));

  const headerRow = [HEADERS];
  console.log(`Q1:Y1 headers:`, JSON.stringify(HEADERS));

  const staffValues = STAFF_MEMBERS.map((s) => [s.name, s.currency]); // Q2:R2, Q3:R3 — plain values
  const rowFormulas = STAFF_MEMBERS.map((s, i) => staffRowFormulas(`Q${i + 2}`, s.currency));

  STAFF_MEMBERS.forEach((s, i) => {
    console.log(`Row ${i + 2} (${s.name}, ${s.currency}): S:Y =`, JSON.stringify(rowFormulas[i]));
  });

  if (!APPLY) {
    console.log("\nDry run only — nothing was changed. Re-run with --apply once this looks right.");
    return;
  }

  await setRangeValues(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, WORKSHEET, "Q1:Y1", headerRow);
  await setRangeValues(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, WORKSHEET, "Q2:R3", staffValues);
  await setRangeFormulas(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, WORKSHEET, "S2:Y2", [rowFormulas[0]]);
  await setRangeFormulas(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, WORKSHEET, "S3:Y3", [rowFormulas[1]]);
  await setRangeNumberFormat(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, WORKSHEET, "W2:W3", [
    [dateNumberFormat],
    [dateNumberFormat],
  ]);

  console.log(`\n✅ Wrote headers + formulas to Q1:Y3 on "${WORKSHEET}" (W2:W3 formatted as dates).`);

  const after = await getRangeFormulas(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, WORKSHEET, "Q1:Y3");
  console.log("Resulting Q1:Y3 formulas/values:", JSON.stringify(after.formulas));
})().catch((err) => {
  if (err.response) {
    console.error(`Graph error ${err.response.status}:`, JSON.stringify(err.response.data, null, 2));
  } else {
    console.error(err.message);
  }
  process.exit(1);
});
