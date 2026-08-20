const crypto = require("crypto");
const {
  getTableRows,
  updateTableRowByIndex,
} = require("./graphExcel");
const fxRates = require("./fxRates");

/**
 * Bidirectional RAW <-> category-table sync.
 *
 * Every expense is written to RAW *and* to its category Excel Table as two independent
 * rows (see appendTransactionToExcel in index.js). Both rows share a RowID (written once,
 * at creation time, identical on both sides) and a SyncHash (a fingerprint of the fields
 * that are supposed to mirror between the two rows). Each run:
 *
 *   1. Reads every row of RAW and every category table.
 *   2. Matches rows across tables by RowID.
 *   3. Recomputes each side's current field hash and compares it to the stored SyncHash
 *      (the hash as of the last successful sync):
 *        - neither side changed  -> no-op (just heals a missing/stale SyncHash if needed)
 *        - only RAW changed      -> propagate RAW's fields into the category row
 *        - only the category row changed -> propagate its fields into RAW
 *        - both changed and disagree     -> CONFLICT, left untouched, reported back
 *   4. If Date, Amount, or Currency changed, the frozen Курс/Sum (HKD) are recomputed via
 *      fxRates (date change -> a fresh historical lookup; amount-only change -> the FX
 *      cache returns the same day's already-fetched rate, so it's effectively reused).
 *
 * Deliberately NOT auto-synced (out of scope / too destructive to automate safely):
 *   - Category/Type changes (moving a row to a different category table).
 *   - RAW deletions/insertions that don't have a RowID (pre-migration rows) — skipped.
 *   - Получатель/Сотрудник (Recipient) — RAW has no equivalent column to sync it into.
 */

// Fixed column indices — both tables are written exclusively by buildRawRowValues /
// buildCategoryRowValues in index.js, which always emit columns in this exact order.
const RAW_COLS = {
  date: 0,
  type: 1,
  category: 2,
  subcategory: 3,
  description: 4,
  location: 5,
  amount: 6,
  currency: 7,
  sumHkd: 8,
  notes: 9,
  sender: 10,
  paymentMethod: 11,
  rowId: 12,
  syncHash: 13,
};

// Standard 14-column layout — every category table except Health/MedInsurance (see below).
const STANDARD_CAT_COLS = {
  date: 0,
  subcategory: 1, // Категория
  description: 2, // Описание
  location: 3, // Локация
  amount: 4, // Сумма
  currency: 5, // Валюта
  rate: 6, // Курс
  sumHkd: 7, // Сумма (HKD)
  notes: 8, // Примечание
  paymentMethod: 9, // Метод оплаты
  recipient: 10, // Получатель/Сотрудник
  sender: 11, // Пользователь
  rowId: 12,
  syncHash: 13,
};

/**
 * Health and MedInsurance (Insurance category) kept their 5 insurance columns (Страховка,
 * Статус выплаты, Страховая компания, Период покрытия, Карта) on Mariya's instruction
 * (2026-08-20) instead of getting the same 17->12 cleanup every other category table got on
 * 2026-08-19 — so their column order/width genuinely differs from STANDARD_CAT_COLS. Keep in
 * sync with index.js's buildCategoryRowValues (INSURANCE_CATEGORY_TABLES) and
 * scripts/migrate-workbook.js (INSURANCE_TABLES) — all three must agree.
 *
 * CreditCards was added here 2026-08-20 too, once Mariya confirmed she wants Credit Cards
 * loggable again: a live header check (scripts/migrate-workbook.js) showed CreditCards is
 * currently sitting at the exact same 17-column layout as Health/MedInsurance (it also never got
 * the 2026-08-19 cleanup) — nothing to do with insurance, it just happens to share the identical
 * column positions, so it safely reuses INSURANCE_CAT_COLS rather than needing a third map.
 */
const INSURANCE_CATEGORY_TABLES = ["Health", "MedInsurance", "CreditCards"];
const INSURANCE_CAT_COLS = {
  date: 0,
  subcategory: 1, // Категория
  description: 2, // Описание
  location: 3, // Локация
  amount: 4, // Сумма
  currency: 5, // Валюта
  rate: 6, // Курс
  sumHkd: 7, // Сумма (HKD)
  notes: 8, // Примечание
  paymentMethod: 9, // Метод оплаты
  recipient: 10, // Получатель/Сотрудник
  // 11 Страховка, 12 Статус выплаты — insurance-only, never touched by the sync job
  sender: 13, // Пользователь
  // 14 Страховая компания, 15 Период покрытия, 16 Карта — insurance-only, never touched
  rowId: 17,
  syncHash: 18,
};

/** Picks the right column map for a given category Excel Table name. */
function catColsFor(tableName) {
  return INSURANCE_CATEGORY_TABLES.includes(tableName) ? INSURANCE_CAT_COLS : STANDARD_CAT_COLS;
}

function mergeableFieldsFromRaw(row) {
  return {
    date: row[RAW_COLS.date],
    description: row[RAW_COLS.description],
    location: row[RAW_COLS.location],
    amount: Number(row[RAW_COLS.amount]),
    currency: String(row[RAW_COLS.currency] || "").trim().toUpperCase(),
    notes: row[RAW_COLS.notes],
    paymentMethod: row[RAW_COLS.paymentMethod],
    subcategory: row[RAW_COLS.subcategory],
  };
}

function mergeableFieldsFromCategory(row, catCols) {
  return {
    date: row[catCols.date],
    description: row[catCols.description],
    location: row[catCols.location],
    amount: Number(row[catCols.amount]),
    currency: String(row[catCols.currency] || "").trim().toUpperCase(),
    notes: row[catCols.notes],
    paymentMethod: row[catCols.paymentMethod],
    subcategory: row[catCols.subcategory],
  };
}

function computeSyncHash(fields) {
  return crypto.createHash("sha1").update(JSON.stringify(fields)).digest("hex").slice(0, 16);
}

function roundMoney(n) {
  return Math.round(n * 100) / 100;
}

/**
 * Runs one full sync pass across RAW and every category table.
 * `categoryTableMap`: { categoryLabel: excelTableName } (index.js's CATEGORY_TABLE_MAP).
 */
async function runSync({ driveId, itemId, rawSheetName, categoryTableMap }) {
  const report = {
    checked: 0,
    propagatedToCategory: 0,
    propagatedToRaw: 0,
    healed: 0,
    conflicts: [],
    errors: [],
  };

  let rawRows;
  try {
    rawRows = await getTableRows(driveId, itemId, rawSheetName);
  } catch (err) {
    report.errors.push(`RAW: could not read rows — ${err.message}`);
    return report;
  }

  const rawByRowId = new Map();
  rawRows.forEach((row, idx) => {
    const rowId = row[RAW_COLS.rowId];
    if (rowId) rawByRowId.set(String(rowId), { row: [...row], idx });
  });

  for (const [categoryLabel, tableName] of Object.entries(categoryTableMap)) {
    const catCols = catColsFor(tableName);
    let catRows;
    try {
      catRows = await getTableRows(driveId, itemId, tableName);
    } catch (err) {
      report.errors.push(`${tableName}: could not read rows — ${err.message}`);
      continue;
    }

    for (let ci = 0; ci < catRows.length; ci++) {
      const catRow = [...catRows[ci]];
      const rowId = catRow[catCols.rowId];
      if (!rowId) continue; // pre-migration row (no RowID yet) — nothing to match against

      const rawEntry = rawByRowId.get(String(rowId));
      if (!rawEntry) {
        report.conflicts.push(
          `"${tableName}" row ${ci + 1} (RowID ${rowId}) has no matching row in ${rawSheetName} — check whether it was deleted from RAW.`
        );
        continue;
      }
      report.checked++;

      const rawFields = mergeableFieldsFromRaw(rawEntry.row);
      const catFields = mergeableFieldsFromCategory(catRow, catCols);
      const rawHash = computeSyncHash(rawFields);
      const catHash = computeSyncHash(catFields);
      const storedHash = rawEntry.row[RAW_COLS.syncHash] || catRow[catCols.syncHash] || "";

      if (rawHash === catHash) {
        if (storedHash !== rawHash) {
          // Fields already agree but the stored hash is stale/missing — just heal it, no data change.
          try {
            const healedRaw = [...rawEntry.row];
            healedRaw[RAW_COLS.syncHash] = rawHash;
            await updateTableRowByIndex(driveId, itemId, rawSheetName, rawEntry.idx, healedRaw);
            const healedCat = [...catRow];
            healedCat[catCols.syncHash] = rawHash;
            await updateTableRowByIndex(driveId, itemId, tableName, ci, healedCat);
            report.healed++;
          } catch (err) {
            report.errors.push(`${tableName} row ${ci + 1}: failed to heal SyncHash — ${err.message}`);
          }
        }
        continue;
      }

      const rawChanged = rawHash !== storedHash;
      const catChanged = catHash !== storedHash;

      if (rawChanged && catChanged) {
        report.conflicts.push(
          `"${tableName}" row ${ci + 1} (RowID ${rowId}): both RAW and this tab changed since the last sync — left untouched, please review manually.`
        );
        continue;
      }

      try {
        if (rawChanged) {
          const { sumHkd } = await propagate({
            driveId,
            itemId,
            fromFields: rawFields,
            targetTableName: tableName,
            targetRow: catRow,
            targetRowIndex: ci,
            targetIsRaw: false,
            catCols,
          });
          const newHash = computeSyncHash(rawFields);
          const healedRaw = [...rawEntry.row];
          healedRaw[RAW_COLS.syncHash] = newHash;
          // RAW is the source of truth here (it's what changed), but RAW also carries its own
          // computed Sum (HKD) column (RAW_COLS.sumHkd) that must stay in step with whatever
          // amount/currency/rate propagate() just used for the category row — otherwise RAW's
          // own Sum (HKD) goes stale forever after any RAW-side edit. Reuse the exact sumHkd
          // propagate() just computed (and wrote into the category row) instead of recomputing
          // it a second time, so RAW and the category row can never disagree even if the FX
          // rate happened to change between two separate lookups.
          healedRaw[RAW_COLS.sumHkd] = sumHkd;
          await updateTableRowByIndex(driveId, itemId, rawSheetName, rawEntry.idx, healedRaw);
          report.propagatedToCategory++;
        } else if (catChanged) {
          await propagate({
            driveId,
            itemId,
            fromFields: catFields,
            targetTableName: rawSheetName,
            targetRow: rawEntry.row,
            targetRowIndex: rawEntry.idx,
            targetIsRaw: true,
            catCols,
          });
          const newHash = computeSyncHash(catFields);
          const healedCat = [...catRow];
          healedCat[catCols.syncHash] = newHash;
          await updateTableRowByIndex(driveId, itemId, tableName, ci, healedCat);
          report.propagatedToRaw++;
        }
      } catch (err) {
        report.errors.push(`${tableName} row ${ci + 1} (RowID ${rowId}): propagation failed — ${err.message}`);
      }
    }
  }

  return report;
}

/**
 * Writes `fromFields` into the target row (RAW or category shape) and recomputes the frozen FX
 * conversion, then saves it. `catCols` is required whenever the target (or source) side is a
 * category table, so insurance-preserving tables (Health/MedInsurance) get written at the right
 * column offsets instead of the standard ones.
 */
async function propagate({ driveId, itemId, fromFields, targetTableName, targetRow, targetRowIndex, targetIsRaw, catCols }) {
  const rate = await fxRates.getRateToHkd(fromFields.currency, fromFields.date);
  const sumHkd = roundMoney(fromFields.amount * rate);

  const updated = [...targetRow];
  if (targetIsRaw) {
    updated[RAW_COLS.date] = fromFields.date;
    updated[RAW_COLS.subcategory] = fromFields.subcategory;
    updated[RAW_COLS.description] = fromFields.description;
    updated[RAW_COLS.location] = fromFields.location;
    updated[RAW_COLS.amount] = fromFields.amount;
    updated[RAW_COLS.currency] = fromFields.currency;
    updated[RAW_COLS.sumHkd] = sumHkd;
    updated[RAW_COLS.notes] = fromFields.notes;
    updated[RAW_COLS.paymentMethod] = fromFields.paymentMethod;
  } else {
    updated[catCols.date] = fromFields.date;
    updated[catCols.subcategory] = fromFields.subcategory;
    updated[catCols.description] = fromFields.description;
    updated[catCols.location] = fromFields.location;
    updated[catCols.amount] = fromFields.amount;
    updated[catCols.currency] = fromFields.currency;
    updated[catCols.rate] = rate;
    updated[catCols.sumHkd] = sumHkd;
    updated[catCols.notes] = fromFields.notes;
    updated[catCols.paymentMethod] = fromFields.paymentMethod;
  }

  await updateTableRowByIndex(driveId, itemId, targetTableName, targetRowIndex, updated);
  return { rate, sumHkd };
}

function formatSyncReport(report) {
  const lines = [
    `Checked ${report.checked} matched row${report.checked === 1 ? "" : "s"}.`,
    `→ category: ${report.propagatedToCategory}, → RAW: ${report.propagatedToRaw}, healed: ${report.healed}.`,
  ];
  if (report.conflicts.length) {
    lines.push(`⚠️ ${report.conflicts.length} conflict(s):`);
    report.conflicts.slice(0, 10).forEach((c) => lines.push(`- ${c}`));
    if (report.conflicts.length > 10) lines.push(`...and ${report.conflicts.length - 10} more.`);
  }
  if (report.errors.length) {
    lines.push(`❌ ${report.errors.length} error(s):`);
    report.errors.slice(0, 10).forEach((e) => lines.push(`- ${e}`));
    if (report.errors.length > 10) lines.push(`...and ${report.errors.length - 10} more.`);
  }
  return lines.join("\n");
}

module.exports = {
  runSync,
  formatSyncReport,
  computeSyncHash,
  mergeableFieldsFromRaw,
  RAW_COLS,
  STANDARD_CAT_COLS,
  INSURANCE_CAT_COLS,
  INSURANCE_CATEGORY_TABLES,
  catColsFor,
};
