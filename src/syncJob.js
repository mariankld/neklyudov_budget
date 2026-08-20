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

const CAT_COLS = {
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

function mergeableFieldsFromCategory(row) {
  return {
    date: row[CAT_COLS.date],
    description: row[CAT_COLS.description],
    location: row[CAT_COLS.location],
    amount: Number(row[CAT_COLS.amount]),
    currency: String(row[CAT_COLS.currency] || "").trim().toUpperCase(),
    notes: row[CAT_COLS.notes],
    paymentMethod: row[CAT_COLS.paymentMethod],
    subcategory: row[CAT_COLS.subcategory],
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
    let catRows;
    try {
      catRows = await getTableRows(driveId, itemId, tableName);
    } catch (err) {
      report.errors.push(`${tableName}: could not read rows — ${err.message}`);
      continue;
    }

    for (let ci = 0; ci < catRows.length; ci++) {
      const catRow = [...catRows[ci]];
      const rowId = catRow[CAT_COLS.rowId];
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
      const catFields = mergeableFieldsFromCategory(catRow);
      const rawHash = computeSyncHash(rawFields);
      const catHash = computeSyncHash(catFields);
      const storedHash = rawEntry.row[RAW_COLS.syncHash] || catRow[CAT_COLS.syncHash] || "";

      if (rawHash === catHash) {
        if (storedHash !== rawHash) {
          // Fields already agree but the stored hash is stale/missing — just heal it, no data change.
          try {
            const healedRaw = [...rawEntry.row];
            healedRaw[RAW_COLS.syncHash] = rawHash;
            await updateTableRowByIndex(driveId, itemId, rawSheetName, rawEntry.idx, healedRaw);
            const healedCat = [...catRow];
            healedCat[CAT_COLS.syncHash] = rawHash;
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
          await propagate({
            driveId,
            itemId,
            fromFields: rawFields,
            targetTableName: tableName,
            targetRow: catRow,
            targetRowIndex: ci,
            targetIsRaw: false,
          });
          const newHash = computeSyncHash(rawFields);
          const healedRaw = [...rawEntry.row];
          healedRaw[RAW_COLS.syncHash] = newHash;
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
          });
          const newHash = computeSyncHash(catFields);
          const healedCat = [...catRow];
          healedCat[CAT_COLS.syncHash] = newHash;
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

/** Writes `fromFields` into the target row (RAW or category shape) and recomputes the frozen FX conversion, then saves it. */
async function propagate({ driveId, itemId, fromFields, targetTableName, targetRow, targetRowIndex, targetIsRaw }) {
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
    updated[CAT_COLS.date] = fromFields.date;
    updated[CAT_COLS.subcategory] = fromFields.subcategory;
    updated[CAT_COLS.description] = fromFields.description;
    updated[CAT_COLS.location] = fromFields.location;
    updated[CAT_COLS.amount] = fromFields.amount;
    updated[CAT_COLS.currency] = fromFields.currency;
    updated[CAT_COLS.rate] = rate;
    updated[CAT_COLS.sumHkd] = sumHkd;
    updated[CAT_COLS.notes] = fromFields.notes;
    updated[CAT_COLS.paymentMethod] = fromFields.paymentMethod;
  }

  await updateTableRowByIndex(driveId, itemId, targetTableName, targetRowIndex, updated);
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
  CAT_COLS,
};
