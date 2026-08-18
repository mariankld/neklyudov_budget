require('dotenv').config();
const { resolveShareToDriveItem, listWorkbookTables } = require('../src/graphExcel');

(async () => {
  const shareUrl = process.env.EXCEL_SHARE_URL;
  if (!shareUrl) {
    console.error('Set EXCEL_SHARE_URL in .env first.');
    process.exit(1);
  }

  try {
    const { driveId, itemId, name } = await resolveShareToDriveItem(shareUrl);
    console.log(`Resolved "${name}"\n`);
    console.log(`EXCEL_DRIVE_ID=${driveId}`);
    console.log(`EXCEL_ITEM_ID=${itemId}`);
    console.log('\nPaste the two lines above into .env.\n');

    const tables = await listWorkbookTables(driveId, itemId);
    console.log('Tables found in the workbook:');
    tables.forEach((t) => console.log(`  - ${t.worksheet} -> ${t.name}`));
  } catch (err) {
    if (err.response) {
      console.error(`Graph error ${err.response.status}:`, JSON.stringify(err.response.data, null, 2));
      if (err.response.status === 401 || err.response.status === 403) {
        console.error(
          "\nThis usually means admin consent hasn't been granted yet for the app's API permissions " +
            '(Entra ID > App registrations > Neklyudov Budget > API permissions > Grant admin consent).'
        );
      }
    } else {
      console.error(err.message);
    }
    process.exit(1);
  }
})();
