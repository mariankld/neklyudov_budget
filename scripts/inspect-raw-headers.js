require("dotenv").config();
const { getTableHeaders, getRangeFormulas } = require("../src/graphExcel");

// Read-only diagnostic — no writes. Two things needed before H20/H21 can be safely hand-edited
// to reference RAW instead of iterating every category table:
//   1. RAW's real header names (fix-summary-formulas.js's printed suggestion for H20/H21 mixes
//      English placeholder names — "Payment Method", "Type", "Sum (HKD)" — with one Cyrillic name
//      it happened to get right — "Дата" — so it can't be pasted in as-is; every column reference
//      needs to match RAW's actual headers exactly, or the formula will just #NAME? error).
//   2. The live formula (or hardcoded value) behind 'Credit Cards'!V24, since H20 is
//      'Credit Cards'!V24 + <sum of card-tagged purchases in every other category table> — if V24
//      is a hand-typed number rather than a formula that reads the CreditCards table, it won't
//      pick up whatever the bot now logs directly into CreditCards (fees, interest, etc. — not
//      repayments, those are H21's job) and will quietly drift once real logging starts.
//
// Run:
//   node scripts/inspect-raw-headers.js

const RAW_SHEET = (process.env.RAW_SHEET_NAME || "RAW").trim();
const CREDIT_CARDS_SHEET = "Credit Cards";

(async () => {
  const { EXCEL_DRIVE_ID, EXCEL_ITEM_ID } = process.env;
  if (!EXCEL_DRIVE_ID || !EXCEL_ITEM_ID) {
    console.error("Set EXCEL_DRIVE_ID and EXCEL_ITEM_ID in .env first.");
    process.exit(1);
  }

  try {
    const rawHeaders = await getTableHeaders(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, RAW_SHEET);
    console.log(`\n=== "${RAW_SHEET}" table headers (in column order) ===\n`);
    rawHeaders.forEach((h, i) => console.log(`  [${i}] ${h}`));

    console.log(`\n=== "${CREDIT_CARDS_SHEET}"!V24 ===\n`);
    try {
      const v24 = await getRangeFormulas(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, CREDIT_CARDS_SHEET, "V24");
      const cell = v24 && v24.formulas ? v24.formulas[0][0] : v24;
      console.log(`  ${JSON.stringify(cell)}`);
      if (typeof cell === "string" && cell.startsWith("=")) {
        console.log(`  -> formula. Check whether it references the CreditCards table.`);
      } else {
        console.log(`  -> hardcoded value, not a formula. Won't auto-update from new CreditCards rows.`);
      }
    } catch (err) {
      console.error(`  Could not read V24: ${err.message}`);
    }

    console.log("");
  } catch (err) {
    if (err.response) {
      console.error(`Graph error ${err.response.status}:`, JSON.stringify(err.response.data, null, 2));
    } else {
      console.error(err.message);
    }
    process.exit(1);
  }
})();
