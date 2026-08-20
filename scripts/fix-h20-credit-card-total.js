require("dotenv").config();
const { getRangeFormulas, setRangeFormulas } = require("../src/graphExcel");
const { CREDIT_CARD_PAYMENT_METHODS } = require("./lib/paymentMethods");
const CATEGORY_TABLE_MAP = require("./lib/categoryTableMap");

// One-off, hand-verified patch for Summary!H20 — the lifetime "spent on credit cards" total.
// Not handled by scripts/fix-summary-formulas.js on purpose: that script only auto-patches
// formulas that already use a per-category SUMIFS template it can clone. H20 instead matches
// specific credit-card names (Master SC, Visa BOC, Master BEA, Master Citic) via SUMIF on the
// "Метод оплаты" (Payment Method) column — Mariya confirmed (2026-08-19) that this is exactly
// right: a transaction counts toward the credit-card total when its Payment Method is one of
// these 4 known credit cards, nothing else needs to change about *how* it matches.
//
// v2 (2026-08-20) — full rewrite, not an append. Two problems with the original hand-written
// formula this replaces:
//   1. Coverage: it only summed 7 of 15 expense category tables, so a doctor's visit, rent
//      payment, tuition, utility bill, etc. paid on a credit card never reached the total. Fixed
//      by summing every category table in scripts/lib/categoryTableMap.js except "Credit Cards"
//      itself (already counted separately via 'Credit Cards'!T24).
//   2. Even within those 7 covered tables, most only checked 1–2 of the 4 cards by hand (e.g.
//      TransportTable only ever checked Visa BOC) — so a Transportation expense paid on Master SC
//      would have been silently dropped too. Fixed by checking all 4 cards on every table,
//      uniformly. NOTE: this can change the computed total vs. the old formula for exactly this
//      reason — it's now catching real spend the old formula missed, not a display-only cleanup.
//   3. Length/readability: SUMIF(table[col],"*card*",table[sum]) once per table×card produced a
//      ~40-term, ~2900-character formula. Collapsed here to one SUMPRODUCT(SUMIF(...,{array},...))
//      term per table — SUMIF accepts an array-constant criteria and returns one result per
//      criterion; SUMPRODUCT adds them up. Also switched from wildcard ("*Master SC*") to exact
//      match against the array, since Payment Method is now enum-constrained (see
//      scripts/lib/paymentMethods.js / src/index.js's getPaymentMethods()) — wildcards were only
//      ever needed to tolerate free-text variants, which can no longer occur.
//
// CREDIT_CARD_PAYMENT_METHODS is sourced from scripts/lib/paymentMethods.js, the same fixed base
// enum fed to OpenAI and used to constrain every Payment Method value written to the workbook —
// so this formula and the values it matches against can never drift out of sync again. Learned
// custom payment methods (src/index.js's customPaymentMethods) are never credit cards and are
// intentionally never added here.
//
// Safety: re-fetches H20 right before writing and aborts if it no longer matches the formula
// captured on 2026-08-19/20 (i.e. someone already changed it) — never overwrites blind.
//
// Run:
//   node scripts/fix-h20-credit-card-total.js            (dry run, prints old vs new)
//   node scripts/fix-h20-credit-card-total.js --apply     (writes it)

const SUMMARY_SHEET_NAME = (process.env.SUMMARY_SHEET_NAME || "Summary").trim();
const APPLY = process.argv.includes("--apply");
const ADDRESS = "H20";

const EXPECTED_CURRENT_FORMULA =
  `='Credit Cards'!T24+SUMIF(Shopping[Метод оплаты],"*Master SC*",Shopping[Сумма (HKD)])+SUMIF(Shopping[Метод оплаты],"*Visa BOC*",Shopping[Сумма (HKD)])+SUMIF(Shopping[Метод оплаты],"*Master BEA*",Shopping[Сумма (HKD)])+SUMIF(Shopping[Метод оплаты],"*Master Citic*",Shopping[Сумма (HKD)])+SUMIF(TransportTable[Метод оплаты],"*Visa BOC*",TransportTable[Сумма (HKD)])+SUMIF(Entertainment[Метод оплаты],"*Visa BOC*",Entertainment[Сумма (HKD)])+SUMIF(Entertainment[Метод оплаты],"*Master BEA*",Entertainment[Сумма (HKD)])+SUMIF(Restaurants[Метод оплаты],"*Visa BOC*",Restaurants[Сумма (HKD)])+SUMIF(TelecomSubscriptions[Метод оплаты],"*Visa BOC*",TelecomSubscriptions[Сумма (HKD)])+SUMIF(Other[Метод оплаты],"*Visa BOC*",Other[Сумма (HKD)])+SUMIF(FamilyStaff[Метод оплаты],"*Master BEA*",FamilyStaff[Сумма (HKD)])`;

// Every expense category table except "Credit Cards" (that one's already counted via the direct
// 'Credit Cards'!T24 reference — everything logged straight into that category is credit-card
// spend by definition). Derived from the same map src/index.js uses, so a newly added category
// table is automatically covered here too without hand-editing this list.
const ALL_EXPENSE_TABLES = Object.entries(CATEGORY_TABLE_MAP)
  .filter(([category, table]) => typeof table === "string" && category !== "Credit Cards")
  .map(([, table]) => table);

/** Excel array-constant literal, e.g. {"Master SC","Visa BOC","Master BEA","Master Citic"}. */
function buildCardArrayLiteral() {
  return `{${CREDIT_CARD_PAYMENT_METHODS.map((card) => `"${card}"`).join(",")}}`;
}

function buildNewFormula() {
  const cardArray = buildCardArrayLiteral();
  const terms = ALL_EXPENSE_TABLES.map(
    (table) => `SUMPRODUCT(SUMIF(${table}[Метод оплаты],${cardArray},${table}[Сумма (HKD)]))`
  );
  return `='Credit Cards'!T24+${terms.join("+")}`;
}

(async () => {
  const { EXCEL_DRIVE_ID, EXCEL_ITEM_ID } = process.env;
  if (!EXCEL_DRIVE_ID || !EXCEL_ITEM_ID) {
    console.error("Set EXCEL_DRIVE_ID and EXCEL_ITEM_ID in .env first.");
    process.exit(1);
  }

  try {
    const current = await getRangeFormulas(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, SUMMARY_SHEET_NAME, ADDRESS);
    const currentFormula = current.formulas && current.formulas[0] && current.formulas[0][0];

    console.log(`\n=== ${SUMMARY_SHEET_NAME}!${ADDRESS} credit-card total fix (v2) — ${APPLY ? "APPLY" : "DRY RUN"} ===\n`);
    console.log(`Current formula:\n  ${currentFormula}\n`);

    if (currentFormula !== EXPECTED_CURRENT_FORMULA) {
      console.error(
        `⚠️ ${ADDRESS} no longer matches the formula captured during the 2026-08-19 audit — someone may have already edited it (maybe this script's v1 already ran?).\nRefusing to overwrite blind. Re-run scripts/audit-summary-formulas.js and update EXPECTED_CURRENT_FORMULA in this script if the new formula still needs the same fix.`
      );
      process.exit(1);
    }

    const newFormula = buildNewFormula();
    console.log(
      `Proposed formula — covers all ${ALL_EXPENSE_TABLES.length} expense tables (${ALL_EXPENSE_TABLES.join(", ")}), all 4 cards each, exact match:\n  ${newFormula}\n`
    );
    console.log(
      `⚠️ Behavior change, not just formatting: the old formula only checked 1–4 of the 4 cards on 6 of its 7 tables (e.g. TransportTable only ever checked Visa BOC). This version checks all 4 cards on all ${ALL_EXPENSE_TABLES.length} tables, so the total may increase — it will now include credit-card spend the old formula was silently missing.\n`
    );

    if (APPLY) {
      await setRangeFormulas(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, SUMMARY_SHEET_NAME, ADDRESS, [[newFormula]]);
      console.log(`✅ Wrote ${ADDRESS}.`);
    } else {
      console.log("Dry run only — nothing was written. Re-run with --apply once this looks right.");
    }
  } catch (err) {
    if (err.response) {
      console.error(`Graph error ${err.response.status}:`, JSON.stringify(err.response.data, null, 2));
    } else {
      console.error(err.message);
    }
    process.exit(1);
  }
})();
