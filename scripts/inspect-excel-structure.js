require('dotenv').config();
const { getAccessToken } = require('../src/graphExcel');
const axios = require('axios');

// Dumps the header row of every Excel Table in the live workbook (via the
// Graph headerRowRange endpoint, so it works regardless of which physical
// row the headers sit on). Used to compare each category tab's columns
// against RAW's, since the workbook is the source of truth and the bot's
// category-tab writes need to match it exactly.

(async () => {
  const { EXCEL_DRIVE_ID, EXCEL_ITEM_ID } = process.env;
  if (!EXCEL_DRIVE_ID || !EXCEL_ITEM_ID) {
    console.error('Set EXCEL_DRIVE_ID and EXCEL_ITEM_ID in .env first.');
    process.exit(1);
  }

  const base = `https://graph.microsoft.com/v1.0/drives/${EXCEL_DRIVE_ID}/items/${EXCEL_ITEM_ID}/workbook`;

  try {
    const token = await getAccessToken();
    const headers = { Authorization: `Bearer ${token}` };

    const { data: tablesData } = await axios.get(`${base}/tables?$expand=worksheet`, { headers });
    const tables = tablesData.value;

    // Also grab worksheet order (position) so we can report tabs in the order
    // they physically appear in the workbook, left to right.
    const { data: sheetsData } = await axios.get(`${base}/worksheets?$select=name,position`, { headers });
    const posByName = Object.fromEntries(sheetsData.value.map((s) => [s.name, s.position]));

    const results = [];
    for (const t of tables) {
      const { data: hdr } = await axios.get(
        `${base}/tables('${encodeURIComponent(t.name)}')/headerRowRange`,
        { headers }
      );
      results.push({
        worksheet: t.worksheet ? t.worksheet.name : '(unknown)',
        position: t.worksheet ? posByName[t.worksheet.name] : 999,
        table: t.name,
        columns: hdr.values[0],
      });
    }

    results.sort((a, b) => a.position - b.position);

    console.log(JSON.stringify(results, null, 2));
  } catch (err) {
    if (err.response) {
      console.error(`Graph error ${err.response.status}:`, JSON.stringify(err.response.data, null, 2));
    } else {
      console.error(err.message);
    }
    process.exit(1);
  }
})();
