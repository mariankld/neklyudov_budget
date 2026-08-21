"use strict";

/**
 * Lightweight, framework-free test suite for src/syncJob.js.
 *
 * No test runner (jest/mocha) is installed in this repo, so this is a plain Node script:
 * fakes graphExcel (in-memory "workbook") and fxRates (in-memory rate table) via require.cache
 * injection, then drives runSync() through each scenario and asserts on the resulting writes.
 *
 * Run:  node tests/syncJob.test.js
 */

const path = require("path");
const assert = require("assert");

// ---------------------------------------------------------------------------
// Fakes, installed into require.cache BEFORE syncJob.js (and its own requires) load, so
// syncJob's `const { getTableRows, updateTableRowByIndex } = require("./graphExcel")` and
// `const fxRates = require("./fxRates")` pick these up instead of hitting real Graph/HTTP calls.
// ---------------------------------------------------------------------------

const graphExcelPath = require.resolve("../src/graphExcel");
const fxRatesPath = require.resolve("../src/fxRates");

let workbook; // tableName -> array of row-arrays
let writes; // log of every updateTableRowByIndex call, in order
let deletes; // log of every deleteTableRowByIndex call, in order

function resetWorkbook(tables) {
  workbook = {};
  for (const [name, rows] of Object.entries(tables)) {
    workbook[name] = rows.map((r) => [...r]);
  }
  writes = [];
  deletes = [];
}

const fakeGraphExcel = {
  async getTableRows(driveId, itemId, tableName) {
    if (!(tableName in workbook)) throw new Error(`fakeGraphExcel: no such table "${tableName}"`);
    return workbook[tableName].map((r) => [...r]);
  },
  async updateTableRowByIndex(driveId, itemId, tableName, rowIndex, values) {
    if (!(tableName in workbook)) throw new Error(`fakeGraphExcel: no such table "${tableName}"`);
    workbook[tableName][rowIndex] = [...values];
    writes.push({ tableName, rowIndex, values: [...values] });
  },
  // Mirrors the real Graph endpoint's behavior: removes the row and shifts every subsequent
  // row's index down by one — exactly the hazard runSync's descending-order deletion loop is
  // designed to avoid corrupting.
  async deleteTableRowByIndex(driveId, itemId, tableName, rowIndex) {
    if (!(tableName in workbook)) throw new Error(`fakeGraphExcel: no such table "${tableName}"`);
    workbook[tableName].splice(rowIndex, 1);
    deletes.push({ tableName, rowIndex });
  },
  // Unused by syncJob.js but harmless to include in case of future use.
  async appendTableRow() {
    throw new Error("fakeGraphExcel.appendTableRow not implemented for tests");
  },
};

let rateMap; // "CUR|dd/mm/yyyy" -> rate
let rateCalls; // [{currency, date}]

function resetRates(entries) {
  rateMap = new Map(Object.entries(entries));
  rateCalls = [];
}

const fakeFxRates = {
  TRACKED_CURRENCIES: ["HKD", "USD", "EUR", "GBP", "THB", "CNY", "RUB"],
  async getRateToHkd(currency, ddmmyyyy) {
    rateCalls.push({ currency, date: ddmmyyyy });
    const cur = String(currency || "").trim().toUpperCase();
    if (cur === "HKD") return 1;
    const key = `${cur}|${ddmmyyyy}`;
    if (!rateMap.has(key)) throw new Error(`fakeFxRates: no rate seeded for ${key}`);
    return rateMap.get(key);
  },
  primeCache() {},
  clampToToday(d) {
    return d;
  },
};

require.cache[graphExcelPath] = {
  id: graphExcelPath,
  filename: graphExcelPath,
  loaded: true,
  exports: fakeGraphExcel,
};
require.cache[fxRatesPath] = {
  id: fxRatesPath,
  filename: fxRatesPath,
  loaded: true,
  exports: fakeFxRates,
};

const { runSync, computeSyncHash, RAW_COLS, STANDARD_CAT_COLS } = require("../src/syncJob");

// ---------------------------------------------------------------------------
// Row-building helpers (mirror src/index.js's buildRawRowValues / buildCategoryRowValues for
// the standard 14-column category-table shape used by e.g. Restaurants).
// ---------------------------------------------------------------------------

function mergeableFields({ date, description, location, amount, currency, notes, paymentMethod, subcategory }) {
  return { date, description, location, amount: Number(amount), currency, notes, paymentMethod, subcategory };
}

function rawRow({ date, subcategory, description, location, amount, currency, sumHkd, notes = "", sender = "Mariya", paymentMethod = "PayMe", rowId = "rid1", syncHash }) {
  const row = [];
  row[RAW_COLS.date] = date;
  row[RAW_COLS.type] = "Expense";
  row[RAW_COLS.category] = "Restaurants";
  row[RAW_COLS.subcategory] = subcategory;
  row[RAW_COLS.description] = description;
  row[RAW_COLS.location] = location;
  row[RAW_COLS.amount] = amount;
  row[RAW_COLS.currency] = currency;
  row[RAW_COLS.sumHkd] = sumHkd;
  row[RAW_COLS.notes] = notes;
  row[RAW_COLS.sender] = sender;
  row[RAW_COLS.paymentMethod] = paymentMethod;
  row[RAW_COLS.rowId] = rowId;
  row[RAW_COLS.syncHash] = syncHash;
  return row;
}

function catRow({ date, subcategory, description, location, amount, currency, rate, sumHkd, notes = "", paymentMethod = "PayMe", recipient = "", sender = "Mariya", rowId = "rid1", syncHash }) {
  const row = [];
  row[STANDARD_CAT_COLS.date] = date;
  row[STANDARD_CAT_COLS.subcategory] = subcategory;
  row[STANDARD_CAT_COLS.description] = description;
  row[STANDARD_CAT_COLS.location] = location;
  row[STANDARD_CAT_COLS.amount] = amount;
  row[STANDARD_CAT_COLS.currency] = currency;
  row[STANDARD_CAT_COLS.rate] = rate;
  row[STANDARD_CAT_COLS.sumHkd] = sumHkd;
  row[STANDARD_CAT_COLS.notes] = notes;
  row[STANDARD_CAT_COLS.paymentMethod] = paymentMethod;
  row[STANDARD_CAT_COLS.recipient] = recipient;
  row[STANDARD_CAT_COLS.sender] = sender;
  row[STANDARD_CAT_COLS.rowId] = rowId;
  row[STANDARD_CAT_COLS.syncHash] = syncHash;
  return row;
}

const CATEGORY_TABLE_MAP = { Restaurants: "Restaurants" };

// ---------------------------------------------------------------------------
// Tiny test runner
// ---------------------------------------------------------------------------

const tests = [];
function test(name, fn) {
  tests.push({ name, fn });
}

// ---------------------------------------------------------------------------
// TEST 1 (regression test for the priority bug): RAW-side edit (amount 100 -> 200, same
// day/currency, so fxRates returns its already-cached rate) must propagate into the category
// row AND rewrite RAW's own Sum (HKD) — not just heal RAW's SyncHash.
// ---------------------------------------------------------------------------
test("rawChanged: propagates to category row AND rewrites RAW's own Sum (HKD)", async () => {
  const inSyncFields = mergeableFields({
    date: "10/08/2026",
    description: "Dinner",
    location: "Hong Kong",
    amount: 100,
    currency: "USD",
    notes: "",
    paymentMethod: "PayMe",
    subcategory: "Dinner",
  });
  const hash0 = computeSyncHash(inSyncFields);

  resetWorkbook({
    RAW: [
      rawRow({ date: "10/08/2026", subcategory: "Dinner", description: "Dinner", location: "Hong Kong", amount: 200 /* edited on RAW */, currency: "USD", sumHkd: 780 /* stale — 100*7.8 */, syncHash: hash0 }),
    ],
    Restaurants: [
      catRow({ date: "10/08/2026", subcategory: "Dinner", description: "Dinner", location: "Hong Kong", amount: 100 /* unchanged */, currency: "USD", rate: 7.8, sumHkd: 780, syncHash: hash0 }),
    ],
  });
  resetRates({ "USD|10/08/2026": 7.8 });

  const report = await runSync({ driveId: "d", itemId: "i", rawSheetName: "RAW", categoryTableMap: CATEGORY_TABLE_MAP });

  assert.strictEqual(report.errors.length, 0, `expected no errors, got: ${JSON.stringify(report.errors)}`);
  assert.strictEqual(report.conflicts.length, 0, `expected no conflicts, got: ${JSON.stringify(report.conflicts)}`);
  assert.strictEqual(report.propagatedToCategory, 1, "expected exactly one RAW->category propagation");

  const newRaw = workbook.RAW[0];
  const newCat = workbook.Restaurants[0];

  assert.strictEqual(newCat[STANDARD_CAT_COLS.amount], 200, "category row's amount should mirror RAW's new amount");
  assert.strictEqual(newCat[STANDARD_CAT_COLS.sumHkd], 1560, "category row's Sum (HKD) should be recomputed (200 * 7.8)");

  // The actual regression: RAW's OWN Sum (HKD) must also be rewritten to 1560, not left at the
  // stale 780 with only SyncHash healed.
  assert.strictEqual(newRaw[RAW_COLS.sumHkd], 1560, "RAW's own Sum (HKD) must be recomputed, not left stale");
  assert.strictEqual(newRaw[RAW_COLS.syncHash], newCat[RAW_COLS.syncHash] === undefined ? newRaw[RAW_COLS.syncHash] : newRaw[RAW_COLS.syncHash], "sanity");
  assert.notStrictEqual(newRaw[RAW_COLS.syncHash], hash0, "RAW's SyncHash must be healed to the new hash");
});

// ---------------------------------------------------------------------------
// TEST 2: category-side edit propagates into RAW (including RAW's Sum (HKD), via the
// targetIsRaw:true path — this already worked before the fix, kept as a regression guard).
// ---------------------------------------------------------------------------
test("catChanged: propagates to RAW, including RAW's Sum (HKD)", async () => {
  const inSyncFields = mergeableFields({
    date: "05/08/2026",
    description: "Groceries",
    location: "Hong Kong",
    amount: 50,
    currency: "HKD",
    notes: "",
    paymentMethod: "Octopus",
    subcategory: "Groceries",
  });
  const hash0 = computeSyncHash(inSyncFields);

  resetWorkbook({
    RAW: [
      rawRow({ date: "05/08/2026", subcategory: "Groceries", description: "Groceries", location: "Hong Kong", amount: 50, currency: "HKD", sumHkd: 50, paymentMethod: "Octopus", syncHash: hash0 }),
    ],
    Restaurants: [
      catRow({ date: "05/08/2026", subcategory: "Groceries", description: "Groceries", location: "Hong Kong", amount: 80 /* edited on category side */, currency: "HKD", rate: 1, sumHkd: 80, paymentMethod: "Octopus", syncHash: hash0 }),
    ],
  });
  resetRates({ "HKD|05/08/2026": 1 });

  const report = await runSync({ driveId: "d", itemId: "i", rawSheetName: "RAW", categoryTableMap: CATEGORY_TABLE_MAP });

  assert.strictEqual(report.propagatedToRaw, 1, "expected exactly one category->RAW propagation");
  const newRaw = workbook.RAW[0];
  assert.strictEqual(newRaw[RAW_COLS.amount], 80, "RAW's amount should mirror the category row's new amount");
  assert.strictEqual(newRaw[RAW_COLS.sumHkd], 80, "RAW's Sum (HKD) should be recomputed from the category row's new amount");
});

// ---------------------------------------------------------------------------
// TEST 3: both sides changed differently since the last sync -> reported as a conflict, and
// NEITHER row is written.
// ---------------------------------------------------------------------------
test("both changed differently: reported as CONFLICT, both rows left untouched", async () => {
  const inSyncFields = mergeableFields({
    date: "01/08/2026",
    description: "Taxi",
    location: "Hong Kong",
    amount: 40,
    currency: "HKD",
    notes: "",
    paymentMethod: "Cash",
    subcategory: "Taxi",
  });
  const hash0 = computeSyncHash(inSyncFields);

  const originalRaw = rawRow({ date: "01/08/2026", subcategory: "Taxi", description: "Taxi", location: "Hong Kong", amount: 45 /* RAW edited one way */, currency: "HKD", sumHkd: 40, paymentMethod: "Cash", syncHash: hash0 });
  const originalCat = catRow({ date: "01/08/2026", subcategory: "Taxi", description: "Taxi", location: "Hong Kong", amount: 60 /* category edited a different way */, currency: "HKD", rate: 1, sumHkd: 40, paymentMethod: "Cash", syncHash: hash0 });

  resetWorkbook({ RAW: [originalRaw], Restaurants: [originalCat] });
  resetRates({ "HKD|01/08/2026": 1 });

  const report = await runSync({ driveId: "d", itemId: "i", rawSheetName: "RAW", categoryTableMap: CATEGORY_TABLE_MAP });

  assert.strictEqual(report.conflicts.length, 1, "expected exactly one conflict");
  assert.strictEqual(report.propagatedToCategory, 0);
  assert.strictEqual(report.propagatedToRaw, 0);
  assert.deepStrictEqual(workbook.RAW[0], originalRaw, "RAW row must be left completely untouched on conflict");
  assert.deepStrictEqual(workbook.Restaurants[0], originalCat, "category row must be left completely untouched on conflict");
});

// ---------------------------------------------------------------------------
// TEST 4: Date change on RAW triggers a fresh historical rate lookup (new date passed to
// fxRates.getRateToHkd), and Sum (HKD) is recomputed from that historical rate.
// ---------------------------------------------------------------------------
test("date change on RAW: fresh historical FX lookup for the new date drives the new Sum (HKD)", async () => {
  const oldDate = "01/06/2026";
  const newDate = "15/06/2026";
  const inSyncFields = mergeableFields({
    date: oldDate,
    description: "Hotel",
    location: "Bangkok",
    amount: 1000,
    currency: "THB",
    notes: "",
    paymentMethod: "Visa BOC",
    subcategory: "Hotel",
  });
  const hash0 = computeSyncHash(inSyncFields);

  resetWorkbook({
    RAW: [rawRow({ date: newDate /* edited */, subcategory: "Hotel", description: "Hotel", location: "Bangkok", amount: 1000, currency: "THB", sumHkd: 220, paymentMethod: "Visa BOC", syncHash: hash0 })],
    Restaurants: [catRow({ date: oldDate, subcategory: "Hotel", description: "Hotel", location: "Bangkok", amount: 1000, currency: "THB", rate: 0.22, sumHkd: 220, paymentMethod: "Visa BOC", syncHash: hash0 })],
  });
  // Different rate on the new date than the old one, to prove it's a fresh lookup, not reuse.
  resetRates({ [`THB|${oldDate}`]: 0.22, [`THB|${newDate}`]: 0.225 });

  const report = await runSync({ driveId: "d", itemId: "i", rawSheetName: "RAW", categoryTableMap: CATEGORY_TABLE_MAP });

  assert.strictEqual(report.propagatedToCategory, 1);
  assert.ok(
    rateCalls.some((c) => c.currency === "THB" && c.date === newDate),
    "expected a fresh fxRates lookup for the new date"
  );
  const newCat = workbook.Restaurants[0];
  const newRaw = workbook.RAW[0];
  assert.strictEqual(newCat[STANDARD_CAT_COLS.rate], 0.225, "category row's rate should reflect the new date's historical rate");
  assert.strictEqual(newCat[STANDARD_CAT_COLS.sumHkd], 225, "category row's Sum (HKD) = 1000 * 0.225");
  assert.strictEqual(newRaw[RAW_COLS.sumHkd], 225, "RAW's own Sum (HKD) should match too (the fix under test)");
});

// ---------------------------------------------------------------------------
// TEST 5: a RAW row with no RowID (pre-migration row) is skipped, not errored — and produces
// no conflict either, since it's simply never matched against anything.
// ---------------------------------------------------------------------------
test("RAW row with no RowID: skipped silently, no error, no conflict", async () => {
  resetWorkbook({
    RAW: [rawRow({ date: "01/01/2026", subcategory: "Old", description: "Pre-migration row", location: "Hong Kong", amount: 10, currency: "HKD", sumHkd: 10, rowId: "" /* no RowID */, syncHash: "" })],
    Restaurants: [],
  });
  resetRates({});

  const report = await runSync({ driveId: "d", itemId: "i", rawSheetName: "RAW", categoryTableMap: CATEGORY_TABLE_MAP });

  assert.strictEqual(report.errors.length, 0);
  assert.strictEqual(report.conflicts.length, 0);
  assert.strictEqual(report.checked, 0);
  assert.deepStrictEqual(writes, [], "no writes should happen at all");
});

// ---------------------------------------------------------------------------
// TEST 6: a category row with no RowID (pre-migration row on the category side) is likewise
// skipped, not errored.
// ---------------------------------------------------------------------------
test("category row with no RowID: skipped silently, no error", async () => {
  resetWorkbook({
    RAW: [],
    Restaurants: [catRow({ date: "01/01/2026", subcategory: "Old", description: "Pre-migration row", location: "Hong Kong", amount: 10, currency: "HKD", rate: 1, sumHkd: 10, rowId: "", syncHash: "" })],
  });
  resetRates({});

  const report = await runSync({ driveId: "d", itemId: "i", rawSheetName: "RAW", categoryTableMap: CATEGORY_TABLE_MAP });

  assert.strictEqual(report.errors.length, 0);
  assert.strictEqual(report.conflicts.length, 0);
  assert.strictEqual(report.checked, 0);
  assert.deepStrictEqual(writes, []);
});

// ---------------------------------------------------------------------------
// TEST 7: a category row whose RowID no longer exists in RAW (RAW row was deleted) gets
// auto-deleted to mirror it, and the report's changeLines describes exactly that, by
// description, matching Mariya's requested wording.
// ---------------------------------------------------------------------------
test("category row orphaned by a RAW deletion: auto-deleted, with a descriptive changeLine", async () => {
  resetWorkbook({
    RAW: [], // the RAW row was deleted
    Restaurants: [
      catRow({ date: "10/08/2026", subcategory: "Dinner", description: "Sushi dinner", location: "Hong Kong", amount: 100, currency: "USD", rate: 7.8, sumHkd: 780, rowId: "rid-deleted", syncHash: "whatever" }),
    ],
  });
  resetRates({});

  const report = await runSync({ driveId: "d", itemId: "i", rawSheetName: "RAW", categoryTableMap: CATEGORY_TABLE_MAP });

  assert.strictEqual(report.errors.length, 0, `expected no errors, got: ${JSON.stringify(report.errors)}`);
  assert.strictEqual(report.deletedFromCategory, 1, "expected exactly one category-row deletion");
  assert.strictEqual(workbook.Restaurants.length, 0, "the orphaned category row should be gone");
  assert.deepStrictEqual(deletes, [{ tableName: "Restaurants", rowIndex: 0 }]);

  assert.strictEqual(report.changeLines.length, 1);
  assert.strictEqual(
    report.changeLines[0],
    `1 expense deleted in "Restaurants" because it was deleted in RAW (Sushi dinner).`
  );
});

// ---------------------------------------------------------------------------
// TEST 8: multiple orphaned category rows in the same table are all deleted, applied in
// descending index order so earlier deletions never shift the index of a later one out from
// under it (the exact hazard the fake's splice-based deleteTableRowByIndex would expose).
// ---------------------------------------------------------------------------
test("multiple orphaned category rows: all deleted, descending index order avoids index-shift corruption", async () => {
  resetWorkbook({
    RAW: [],
    Restaurants: [
      catRow({ date: "01/08/2026", subcategory: "A", description: "First", location: "HK", amount: 10, currency: "HKD", rate: 1, sumHkd: 10, rowId: "rid-a", syncHash: "x" }),
      catRow({ date: "02/08/2026", subcategory: "B", description: "Second", location: "HK", amount: 20, currency: "HKD", rate: 1, sumHkd: 20, rowId: "rid-b", syncHash: "x" }),
      catRow({ date: "03/08/2026", subcategory: "C", description: "Third", location: "HK", amount: 30, currency: "HKD", rate: 1, sumHkd: 30, rowId: "rid-c", syncHash: "x" }),
    ],
  });
  resetRates({});

  const report = await runSync({ driveId: "d", itemId: "i", rawSheetName: "RAW", categoryTableMap: CATEGORY_TABLE_MAP });

  assert.strictEqual(report.deletedFromCategory, 3);
  assert.strictEqual(workbook.Restaurants.length, 0, "all three orphaned rows should be gone");
  assert.deepStrictEqual(
    deletes.map((d) => d.rowIndex),
    [2, 1, 0],
    "deletes must be applied highest-index-first"
  );
  assert.strictEqual(report.changeLines.length, 3);
});

// ---------------------------------------------------------------------------
// TEST 9: reverse case — a RAW row (type "Expense", matching README's capitalization) whose
// category-table counterpart is missing is reported as a conflict, NOT auto-deleted. This is
// also a regression guard for the type-casing bug fixed alongside this feature: the check used
// to compare against lowercase "expense" while RAW actually stores "Expense".
// ---------------------------------------------------------------------------
test("RAW row orphaned by a category deletion: reported as a conflict, not auto-deleted", async () => {
  resetWorkbook({
    RAW: [
      rawRow({ date: "01/08/2026", subcategory: "Lunch", description: "Lunch", location: "HK", amount: 50, currency: "HKD", sumHkd: 50, rowId: "rid-orphan", syncHash: "x" }),
    ],
    Restaurants: [], // the category row was deleted (or the write failed) — indistinguishable, hence report-only
  });
  resetRates({});

  const report = await runSync({ driveId: "d", itemId: "i", rawSheetName: "RAW", categoryTableMap: CATEGORY_TABLE_MAP });

  assert.strictEqual(report.errors.length, 0);
  assert.strictEqual(report.deletedFromCategory, 0);
  assert.strictEqual(workbook.RAW.length, 1, "RAW row must NOT be auto-deleted");
  assert.strictEqual(report.conflicts.length, 1, "expected exactly one reverse-orphan conflict");
  assert.ok(report.conflicts[0].includes("rid-orphan"), "conflict message should reference the RowID");
});

// ---------------------------------------------------------------------------
// TEST 10: changeLines content for a rawChanged propagation — verifies the exact wording and
// column names match Mariya's requested example format ("Amended columns in X: ...").
// ---------------------------------------------------------------------------
test("changeLines: rawChanged amendment names the exact columns that changed on each side", async () => {
  const inSyncFields = mergeableFields({
    date: "10/08/2026",
    description: "Dinner",
    location: "Hong Kong",
    amount: 100,
    currency: "USD",
    notes: "",
    paymentMethod: "PayMe",
    subcategory: "Dinner",
  });
  const hash0 = computeSyncHash(inSyncFields);

  resetWorkbook({
    RAW: [
      rawRow({ date: "10/08/2026", subcategory: "Dinner", description: "Dinner", location: "Hong Kong", amount: 200, currency: "USD", sumHkd: 780, syncHash: hash0 }),
    ],
    Restaurants: [
      catRow({ date: "10/08/2026", subcategory: "Dinner", description: "Dinner", location: "Hong Kong", amount: 100, currency: "USD", rate: 7.8, sumHkd: 780, syncHash: hash0 }),
    ],
  });
  resetRates({ "USD|10/08/2026": 7.8 });

  const report = await runSync({ driveId: "d", itemId: "i", rawSheetName: "RAW", categoryTableMap: CATEGORY_TABLE_MAP });

  assert.strictEqual(report.changeLines.length, 1);
  const block = report.changeLines[0];
  assert.ok(block.startsWith(`1 expense amended in "Restaurants" because it was changed in RAW (Dinner).`), block);
  assert.ok(block.includes("Amended columns in \"Restaurants\": Сумма, Сумма (HKD)"), block);
  assert.ok(block.includes("Amended columns in RAW: Sum (HKD)"), block);
});

// ---------------------------------------------------------------------------
// TEST 11: changeLines content for a catChanged propagation — the RAW side lists its own
// changed columns; there's no "source changed columns" block since the category table is the
// source here.
// ---------------------------------------------------------------------------
test("changeLines: catChanged amendment names RAW's changed columns", async () => {
  const inSyncFields = mergeableFields({
    date: "05/08/2026",
    description: "Groceries",
    location: "Hong Kong",
    amount: 50,
    currency: "HKD",
    notes: "",
    paymentMethod: "Octopus",
    subcategory: "Groceries",
  });
  const hash0 = computeSyncHash(inSyncFields);

  resetWorkbook({
    RAW: [
      rawRow({ date: "05/08/2026", subcategory: "Groceries", description: "Groceries", location: "Hong Kong", amount: 50, currency: "HKD", sumHkd: 50, paymentMethod: "Octopus", syncHash: hash0 }),
    ],
    Restaurants: [
      catRow({ date: "05/08/2026", subcategory: "Groceries", description: "Groceries", location: "Hong Kong", amount: 80, currency: "HKD", rate: 1, sumHkd: 80, paymentMethod: "Octopus", syncHash: hash0 }),
    ],
  });
  resetRates({ "HKD|05/08/2026": 1 });

  const report = await runSync({ driveId: "d", itemId: "i", rawSheetName: "RAW", categoryTableMap: CATEGORY_TABLE_MAP });

  assert.strictEqual(report.changeLines.length, 1);
  const block = report.changeLines[0];
  assert.ok(block.startsWith(`1 expense amended in RAW because it was changed in "Restaurants" (Groceries).`), block);
  assert.ok(block.includes("Amended columns in RAW: Amount, Sum (HKD)"), block);
  assert.ok(!block.includes(`Amended columns in "Restaurants"`), "no source-side columns block expected in the catChanged case");
});

// ---------------------------------------------------------------------------
// Run
// ---------------------------------------------------------------------------
(async () => {
  let failed = 0;
  for (const { name, fn } of tests) {
    try {
      await fn();
      console.log(`  ✓ ${name}`);
    } catch (err) {
      failed++;
      console.error(`  ✗ ${name}`);
      console.error(`    ${err.message}`);
    }
  }
  console.log(`\n${tests.length - failed}/${tests.length} passed.`);
  process.exit(failed ? 1 : 0);
})();
