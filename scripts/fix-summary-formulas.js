require("dotenv").config();
const { getUsedRange, setRangeFormulas } = require("../src/graphExcel");
const CATEGORY_TABLE_MAP = require("./lib/categoryTableMap");

// Semi-automated patcher for the Summary tab's monthly-total SUMIFS chains, driven by the same
// heuristic scan as scripts/audit-summary-formulas.js. DRY RUN BY DEFAULT — it only prints the
// proposed new formula for each affected cell. Pass --apply to actually write the changes.
//
// How the patch is generated (template substitution, no guessing at unrelated formula parts):
//   1. Find candidate aggregate cells (formula references 2+ category tables) missing one or
//      more of the 7 owner-confirmed categories, same as the audit script.
//   2. Split the formula into its top-level "+"-joined terms (respecting parentheses).
//   3. For each term that references exactly one already-present category table, treat it as a
//      reusable template for that category's SUMIFS/criteria shape.
//   4. For each missing category, clone the best available template term and replace every
//      occurrence of the template's table name with the missing category's table name (e.g.
//      "Health[" -> "CAPEX["), preserving every other argument (date criteria, column names,
//      etc.) untouched.
//   5. Append "+<new term>" for each missing category and write the formula back.
// If a candidate cell has no usable template term (e.g. every present category is referenced
// in some more complex/shared way this script doesn't recognize), it is skipped and reported
// so it can be fixed by hand instead of risking a wrong auto-patch.
//
// The credit-card "hardcoded table/card name" formulas found by the audit are NOT auto-patched
// here — replacing a debt-tracking formula with a Payment-Method-based filter changes what it
// actually computes (not just which categories it includes), so it needs a human to confirm the
// exact cell and intended date range first. This script just re-prints them with a suggested
// replacement pattern for review.
//
// Run:
//   node scripts/fix-summary-formulas.js            (dry run, prints proposed changes)
//   node scripts/fix-summary-formulas.js --apply     (writes the changes)

const SUMMARY_SHEET_NAME = (process.env.SUMMARY_SHEET_NAME || "Summary").trim();
const CATEGORY_TABLE_NAMES = Object.values(CATEGORY_TABLE_MAP).filter((v) => typeof v === "string");
const REQUIRED = CATEGORY_TABLE_MAP.REQUIRED_IN_MONTHLY_TOTALS;
const APPLY = process.argv.includes("--apply");

const CREDIT_CARD_KEYWORDS = ["CreditCards", "VISA", "MASTERCARD", "AMEX", "BOC", "HSBC", "CITI", "CARD"];

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

function colLetterToIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function parseTopLeft(address) {
  const rowMatch = String(address || "").match(/![A-Z]+(\d+)/);
  const colMatch = String(address || "").match(/!([A-Z]+)\d+/);
  return {
    row0: rowMatch ? Number(rowMatch[1]) - 1 : 0,
    col0: colMatch ? colLetterToIndex(colMatch[1]) : 0,
  };
}

function findReferencedTables(text) {
  if (typeof text !== "string") return [];
  return CATEGORY_TABLE_NAMES.filter((table) => text.includes(`${table}[`) || new RegExp(`\\b${table}\\b`).test(text));
}

/** Splits a formula body (without the leading "=") into top-level "+"-joined terms, ignoring "+" nested inside parentheses or string literals. */
function splitTopLevelPlus(body) {
  const terms = [];
  let depth = 0;
  let inString = false;
  let current = "";
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === '"') inString = !inString;
    if (!inString) {
      if (ch === "(") depth++;
      if (ch === ")") depth--;
      if (ch === "+" && depth === 0) {
        terms.push(current);
        current = "";
        continue;
      }
    }
    current += ch;
  }
  if (current.length) terms.push(current);
  return terms;
}

function buildMissingTerm(templateTerm, templateTable, missingTable) {
  const escaped = templateTable.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return templateTerm
    .replace(new RegExp(`${escaped}\\[`, "g"), `${missingTable}[`)
    .replace(new RegExp(`\\b${escaped}\\b(?!\\[)`, "g"), missingTable);
}

function looksLikeCreditCardFormula(formula) {
  if (typeof formula !== "string" || !formula.startsWith("=")) return false;
  const upper = formula.toUpperCase();
  return CREDIT_CARD_KEYWORDS.some((kw) => upper.includes(kw.toUpperCase())) && !upper.includes("PAYMENT");
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

    const patches = [];
    const skipped = [];
    const creditCardFormulas = [];

    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        const formula = formulas[r][c];
        if (typeof formula !== "string" || !formula.startsWith("=")) continue;

        const referenced = findReferencedTables(formula);
        const address = `${colLetter(col0 + c)}${row0 + r + 1}`;

        if (looksLikeCreditCardFormula(formula)) {
          creditCardFormulas.push({ address, formula });
        }

        if (referenced.length < 2) continue;
        const missing = REQUIRED.filter((t) => !referenced.includes(t));
        if (!missing.length) continue;

        const body = formula.slice(1);
        const terms = splitTopLevelPlus(body);
        const termByTable = new Map();
        for (const term of terms) {
          const refs = findReferencedTables(term);
          if (refs.length === 1 && !termByTable.has(refs[0])) {
            termByTable.set(refs[0], term);
          }
        }

        const newTerms = [];
        const unresolvable = [];
        for (const missingTable of missing) {
          // Prefer a template from another already-included required category (most likely to
          // share the exact same criteria shape), else fall back to any present category term.
          const templateTable =
            [...termByTable.keys()].find((t) => REQUIRED.includes(t)) || [...termByTable.keys()][0];
          if (!templateTable) {
            unresolvable.push(missingTable);
            continue;
          }
          const templateTerm = termByTable.get(templateTable);
          newTerms.push(buildMissingTerm(templateTerm, templateTable, missingTable));
        }

        if (unresolvable.length) {
          skipped.push({ address, formula, unresolvable, reason: "no usable template term found in this formula" });
          continue;
        }

        const newFormula = `=${body}+${newTerms.join("+")}`;
        patches.push({ address, oldFormula: formula, newFormula, addedFor: missing });
      }
    }

    console.log(`\n=== Summary formula fix (${SUMMARY_SHEET_NAME}) — ${APPLY ? "APPLY" : "DRY RUN"} ===\n`);

    if (!patches.length) {
      console.log("No auto-fixable cells found (either everything already includes the required categories, or nothing matched the template heuristic — run audit-summary-formulas.js for details).");
    }

    for (const p of patches) {
      console.log(`${p.address} — adding: ${p.addedFor.join(", ")}`);
      console.log(`  old: ${p.oldFormula}`);
      console.log(`  new: ${p.newFormula}\n`);
    }

    if (skipped.length) {
      console.log(`⚠️ ${skipped.length} cell(s) need manual fixing (no reusable template term found):`);
      skipped.forEach((s) => console.log(`  ${s.address}: missing [${s.unresolvable.join(", ")}] — ${s.formula}`));
      console.log("");
    }

    if (APPLY && patches.length) {
      for (const p of patches) {
        await setRangeFormulas(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, SUMMARY_SHEET_NAME, p.address, [[p.newFormula]]);
        console.log(`✅ Wrote ${p.address}`);
      }
    } else if (patches.length) {
      console.log(`Dry run only — nothing was written. Re-run with --apply once these look right.`);
    }

    console.log(`\n--- Credit-card formulas (NOT auto-patched — review manually) ---\n`);
    if (creditCardFormulas.length) {
      creditCardFormulas.forEach((c) => {
        console.log(`  ${c.address}: ${c.formula}`);
      });
      console.log(
        `\nSuggested replacement pattern (adjust the date-range criteria to match what "${creditCardFormulas[0].address}" currently uses):\n` +
          `  =SUMPRODUCT((ISNUMBER(SEARCH("credit",RAW[Payment Method])))*(RAW[Type]="Expense")*(RAW[Дата]>=<start>)*(RAW[Дата]<=<end>)*(RAW[Sum (HKD)]))\n` +
          `Replace <start>/<end> with whatever date-window reference the original formula used (e.g. a month-start/month-end cell), then paste this into the cell manually and confirm the total matches expectations before relying on it.`
      );
    } else {
      console.log("None found by keyword match.");
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
