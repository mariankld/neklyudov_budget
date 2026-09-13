// Kept in sync by hand with CATEGORY_TABLE_MAP in src/index.js. Duplicated here (rather than
// required from src/index.js) so the audit/fix scripts don't have to boot the Express app /
// Telegram webhook wiring just to read a lookup table.
module.exports = {
  "Credit Cards": "CreditCards",
  "Shopping": "Shopping",
  "Transportation": "TransportTable",
  "Utilities": "Utilities",
  "Entertainment": "Entertainment",
  "Restaurants": "Restaurants",
  "Family and Staff": "FamilyStaff",
  "Staff Spending": "StaffExpenses",
  "Staff Advances": "StaffAdvances",
  "Other": "Other",
  "Subscriptions": "TelecomSubscriptions",
  "Travel": "Travel",
  "Health": "Health",
  "Education": "Education",
  "Rent": "Rent",
  "Insurance": "MedInsurance",
  "CAPEX": "CAPEX",
};

// The 7 categories the workbook owner explicitly confirmed (2026-08-19) must be included in
// every monthly Summary total — these were the ones suspected/confirmed missing from the
// SUMIFS chains. Referenced by table name, since that's what a SUMIFS formula would cite.
//
// StaffAdvances (added 2026-09-13, split out of StaffExpenses) is deliberately NOT listed here
// yet — an advance isn't necessarily a final expense, so whether it should count toward the
// monthly Summary total is a call for Mariya to make. Add "StaffAdvances" to this array (and
// update the Summary!H21-style SUMIFS formula the same way StaffExpenses was) if advances should
// count too.
module.exports.REQUIRED_IN_MONTHLY_TOTALS = ["Health", "Education", "Rent", "Travel", "StaffExpenses", "MedInsurance", "CAPEX"];
