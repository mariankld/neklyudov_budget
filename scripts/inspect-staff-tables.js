require("dotenv").config();
const { getTableHeaders, getTableRows, getTableDataBodyRangeAddress } = require("../src/graphExcel");

// Read-only diagnostic (Mariya, 2026-09-13): "Personal Spending" tab currently hosts BOTH
// StaffExpenses and StaffAdvances as two separate Excel Tables. Before splitting them onto two
// tabs, this prints exactly where each table sits (data body range) and how many rows it has, so
// we know whether they're stacked vertically, side by side, and whether StaffAdvances already
// has real data in it (vs. being an empty scaffold).

const TABLES = ["StaffExpenses", "StaffAdvances"];

(async () => {
  const { EXCEL_DRIVE_ID, EXCEL_ITEM_ID } = process.env;
  if (!EXCEL_DRIVE_ID || !EXCEL_ITEM_ID) {
    console.error("Set EXCEL_DRIVE_ID and EXCEL_ITEM_ID in .env first.");
    process.exit(1);
  }

  for (const tableName of TABLES) {
    console.log(`\n=== ${tableName} ===`);
    try {
      const headers = await getTableHeaders(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, tableName);
      console.log(`Columns (${headers.length}): ${headers.join(", ")}`);
      const address = await getTableDataBodyRangeAddress(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, tableName);
      console.log(`Data body range: ${address}`);
      const rows = await getTableRows(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, tableName);
      console.log(`Data rows: ${rows.length}`);
      if (rows.length) {
        console.log("First row:", rows[0]);
        console.log("Last row:", rows[rows.length - 1]);
      }
    } catch (err) {
      console.log(`Could not fully inspect ${tableName}: ${err.message}`);
    }
  }
})();
