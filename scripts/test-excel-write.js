require('dotenv').config();
const { appendTableRow, listWorkbookTables } = require('../src/graphExcel');

// Sanity check: appends one clearly-marked, easy-to-spot row to RAW to confirm
// the app registration actually has WRITE access (not just read), now that
// admin consent is granted. Safe to delete the row afterwards in Excel.

(async () => {
  const { EXCEL_DRIVE_ID, EXCEL_ITEM_ID } = process.env;
  if (!EXCEL_DRIVE_ID || !EXCEL_ITEM_ID) {
    console.error('Set EXCEL_DRIVE_ID and EXCEL_ITEM_ID in .env first (run npm run resolve-excel-share).');
    process.exit(1);
  }

  const testRow = [
    new Date().toISOString().slice(0, 10), // Date
    'Expense', // Type
    'TEST', // Category
    'Graph API smoke test', // Subcategory
    'Automated test row — safe to delete', // Description
    '', // Location
    0, // Amount
    'HKD', // Currency
    0, // Sum (HKD)
    'Written by scripts/test-excel-write.js — delete me', // Notes
    'Claude', // Sender
    '', // Payment Method
  ];

  try {
    await appendTableRow(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, 'RAW', testRow);
    console.log('✅ Wrote a test row to RAW successfully — write access confirmed.');
    console.log('   Row:', testRow);
    console.log('\nOpen the file in Excel and delete that row when you\'re done verifying.');
  } catch (err) {
    if (err.response) {
      console.error(`❌ Graph error ${err.response.status}:`, JSON.stringify(err.response.data, null, 2));
    } else {
      console.error('❌', err.message);
    }
    process.exit(1);
  }
})();
