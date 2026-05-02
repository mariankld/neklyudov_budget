/**
 * Lists all tab titles in GOOGLE_SPREADSHEET_ID (same auth as the app).
 * Usage: node scripts/list-sheet-tabs.js
 */
require("dotenv").config();
const { google } = require("googleapis");

async function main() {
  const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;
  if (!spreadsheetId || !process.env.GOOGLE_REFRESH_TOKEN) {
    console.error("Set GOOGLE_SPREADSHEET_ID and GOOGLE_REFRESH_TOKEN in .env");
    process.exit(1);
  }

  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_OAUTH_REDIRECT_URI || "http://127.0.0.1:3001/oauth/callback"
  );
  oauth2Client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });

  const sheets = google.sheets({ version: "v4", auth: oauth2Client });
  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: "properties(title),sheets(properties(sheetId,title,index))",
  });

  console.log("Spreadsheet:", meta.data.properties?.title || spreadsheetId);
  console.log("--- tabs ---");
  for (const s of meta.data.sheets || []) {
    const p = s.properties;
    console.log(`${p.index}\t${p.title}\t(sheetId=${p.sheetId})`);
  }
}

main().catch((e) => {
  console.error(e.message);
  process.exit(1);
});
