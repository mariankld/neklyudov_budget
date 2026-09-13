require("dotenv").config();
const { listWorksheetsWithPosition, setWorksheetPosition } = require("../src/graphExcel");

// One-off (Mariya, 2026-09-13): move the "Staff Advances" tab so it sits immediately before
// "Staff Spending" in the tab strip, instead of wherever addWorksheet dropped it (typically the
// far right end of the workbook).
//
// Run:
//   node scripts/reorder-staff-tabs.js            (dry run — shows current + target order)
//   node scripts/reorder-staff-tabs.js --apply     (actually moves the tab)

const TAB_TO_MOVE = "Staff Advances";
const TARGET_BEFORE = "Staff Spending";

const APPLY = process.argv.includes("--apply");

(async () => {
  const { EXCEL_DRIVE_ID, EXCEL_ITEM_ID } = process.env;
  if (!EXCEL_DRIVE_ID || !EXCEL_ITEM_ID) {
    console.error("Set EXCEL_DRIVE_ID and EXCEL_ITEM_ID in .env first.");
    process.exit(1);
  }

  console.log(`\n=== Reorder tabs: put "${TAB_TO_MOVE}" right before "${TARGET_BEFORE}" — ${APPLY ? "APPLY" : "DRY RUN"} ===\n`);

  const sheets = await listWorksheetsWithPosition(EXCEL_DRIVE_ID, EXCEL_ITEM_ID);
  console.log("Current tab order:");
  sheets.forEach((s) => console.log(`  ${s.position}: ${s.name}`));
  console.log("");

  const moving = sheets.find((s) => s.name === TAB_TO_MOVE);
  const target = sheets.find((s) => s.name === TARGET_BEFORE);

  if (!moving) {
    console.error(`Could not find a tab named "${TAB_TO_MOVE}".`);
    process.exit(1);
  }
  if (!target) {
    console.error(`Could not find a tab named "${TARGET_BEFORE}".`);
    process.exit(1);
  }

  if (moving.position === target.position - 1) {
    console.log(`"${TAB_TO_MOVE}" is already directly before "${TARGET_BEFORE}" — nothing to do.`);
    return;
  }

  // Setting position to target's CURRENT position slots the moved tab in right there, pushing
  // target (and everything after it) one step to the right — which is exactly "right before".
  const newPosition = target.position;
  console.log(`Will move "${TAB_TO_MOVE}" (currently position ${moving.position}) to position ${newPosition}, i.e. directly before "${TARGET_BEFORE}".\n`);

  if (!APPLY) {
    console.log("Dry run only — nothing was moved. Re-run with --apply once this looks right.");
    return;
  }

  await setWorksheetPosition(EXCEL_DRIVE_ID, EXCEL_ITEM_ID, TAB_TO_MOVE, newPosition);
  console.log(`✅ Moved "${TAB_TO_MOVE}".`);

  const after = await listWorksheetsWithPosition(EXCEL_DRIVE_ID, EXCEL_ITEM_ID);
  console.log("\nNew tab order:");
  after.forEach((s) => console.log(`  ${s.position}: ${s.name}`));
})().catch((err) => {
  if (err.response) {
    console.error(`Graph error ${err.response.status}:`, JSON.stringify(err.response.data, null, 2));
  } else {
    console.error(err.message);
  }
  process.exit(1);
});
