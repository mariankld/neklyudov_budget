require("dotenv").config();
const { getUsedRange } = require("../src/graphExcel");
const CATEGORY_TABLE_MAP = require("./lib/categoryTableMap");

// Read-only audit of the Summary tab's formulas. Never writes anything — safe to run anytime.
//
// Since this script has no prior knowledge of exactly which cells hold "monthly total"
// formulas, it works it out heuristically:
//   1. Reads every formula in the Summary sheet's used range.
//   2. For each formula cell, finds every category Excel Table name (from CATEGORY_TABLE_MAP)
//      referenced inside it (e.g. "Health[", "StaffExpenses[", a SUMIFS argument, etc).
//   3. Any formula that references 2+ distinct category tables is treated as a candidate
//      "monthly total" (or similar aggregate) formula — a single-category SUMIFS wouldn't
//      need more than one.
//   4. For each candidate cell, reports which of the 7 owner-confirmed categories
//      (Health, Education, Rent, Travel, StaffExpenses, MedInsurance, CAPEX) are referenced
//      and which are missing.
//   5. Separately flags any formula anywhere on the sheet that looks like it's computing
//      credit-card debt by hardcoding a table/card name (CreditCards, or common card/bank
//      keywords) instead of filtering by Payment Method — a candidate for the Payment-Method
//      based rewrite Mariya asked for.
//
// Run: npm run audit-summary  (or: node scripts/audit-summary-formulas.js)

const SUMMARY_SHEET_NAME = (process.env.SUMMARY_SHEET_NAME || "Summary").trim();
const CATEGORY_TABLE_NAMES = Object.values(CATEGORY_TABLE_MAP).filter((v) => typeof v === "string");
const REQUIRED = CATEGORY_TABLE_MAP.REQUIRED_IN_MONTHLY_TOTALS;

const CREDIT_CARD_KEYWORDS = [
  "CreditCards",
  "VISA",
  "MASTERCARD",
  "AMEX",
  "BOC",
  "HSBC",
  "CITI",
  "CARD",
];

/** Column letter for a 0-based column index (A, B, ..., Z, AA, ...). */
function colLetter(i) {
  let n = i + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Parses a used-range address like "Summary!A1:M40" into its top-left row/col so per-cell addresses can be reconstructed. */
function parseTopLeft(address) {
  const m = String(address || "").match(/![A-Z]+(\d+)/);
  const colMatch = String(address || "").match(/!([A-Z]+)\d+/);
  return {
    row0: m ? Number(m[1]) - 1 : 0,
    col0: colMatch ? colLetterToIndex(colMatch[1]) : 0,
  };
}

function colLetterToIndex(letters) {
  let n = 0;
  for (const ch of letters) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

function findReferencedTables(formula) {
  if (typeof formula !== "string" || !formula.startsWith("=")) return [];
  return CATEGORY_TABLE_NAMES.filter((table) => formula.includes(`${table}[`) || new RegExp(`\\b${table}\\b`).test(formula));
}

function looksLikeCreditCardFormula(formula) {
  if (typeof formula !== "string" || !formula.startsWith("=")) return false;
  const upper = formula.toUpperCase();
  return CREDIT_CARD_KEYWORDS.some((kw) => upper.includes(kw.toUpperCase()));
}

function isCreditCardPaymentMethodBased(formula) {
  const upper = formula.toUpperCase();
  return upper.includes("PAYMENT") || (upper.includes("SEARCH") && upper.includes("CREDIT"));
}

(async () => {
  const { EXCEL_DRIVE_ID, EXCEL_ITEM_ID } = process.env;
  if (!EXCEL_DRIVE_ID || !EXCEL_ITEM_ID) {
    console.error("Set EXCEL_DRIVE_ID and EXCEL_ITEM_ID in .env first.");
    process.exit(1);
  }

  try {
    const used = await getUsedRange(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, SUMMARY_SHEET_NAME);
    const { row0, col0 } = parseTopLeft(used.address);
    const formulas = used.formulas || [];

    const candidates = [];
    const creditCardFormulas = [];
    const categoryOccursAnywhere = new Set();

    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        const formula = formulas[r][c];
        const referenced = findReferencedTables(formula);
        referenced.forEach((t) => categoryOccursAnywhere.add(t));

        if (referenced.length >= 2) {
          const address = `${colLetter(col0 + c)}${row0 + r + 1}`;
          const missing = REQUIRED.filter((t) => !referenced.includes(t));
          candidates.push({ address, formula, referenced, missing });
        }

        if (looksLikeCreditCardFormula(formula) && !isCreditCardPaymentMethodBased(formula)) {
          const address = `${colLetter(col0 + c)}${row0 + r + 1}`;
          creditCardFormulas.push({ address, formula });
        }
      }
    }

    console.log(`\n=== Summary formula audit (${SUMMARY_SHEET_NAME}) ===\n`);
    console.log(`Used range: ${used.address}`);
    console.log(`Candidate aggregate formulas found (reference 2+ category tables): ${candidates.length}\n`);

    const neverReferenced = REQUIRED.filter((t) => !categoryOccursAnywhere.has(t));
    if (neverReferenced.length) {
      console.log(
        `⚠️ These owner-confirmed categories are NOT referenced anywhere on the ${SUMMARY_SHEET_NAME} tab at all: ${neverReferenced.join(", ")}`
      );
    } else {
      console.log(`✅ All 7 owner-confirmed categories (${REQUIRED.join(", ")}) are referenced somewhere on the sheet.`);
    }
    console.log("");

    const withMissing = candidates.filter((c) => c.missing.length);
    if (withMissing.length) {
      console.log(`⚠️ ${withMissing.length} aggregate formula(s) are missing one or more of the required categories:\n`);
      withMissing.forEach((c) => {
        console.log(`  ${c.address}: missing [${c.missing.join(", ")}]`);
        console.log(`    formula: ${c.formula}`);
      });
    } else if (candidates.length) {
      console.log("✅ Every candidate aggregate formula already references all 7 required categories.");
    } else {
      console.log(
        "No aggregate formulas (referencing 2+ category tables) were detected — either the Summary tab uses a different pattern (e.g. SUMIFS against RAW instead of per-category tables), or the sheet name / used range didn't capture them. Review manually."
      );
    }

    console.log(`\n--- Credit-card formula candidates (hardcoded table/card-name based) ---\n`);
    if (creditCardFormulas.length) {
      creditCardFormulas.forEach((c) => {
        console.log(`  ${c.address}: ${c.formula}`);
      });
      console.log(
        `\n${creditCardFormulas.length} formula(s) look like they hardcode a table or card/bank name. Mariya asked for these to be driven by "Payment Method contains credit" instead — see scripts/fix-summary-formulas.js for a suggested rewrite, but confirm the exact cell before applying.`
      );
    } else {
      console.log("None found by keyword match. If the credit-card debt total lives elsewhere, note its cell address manually.");
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
