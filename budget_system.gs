/**
 * 2Up Budget Tracker — monthly tab automation
 * -------------------------------------------
 * Cycles run from CYCLE_DAY to CYCLE_DAY of the next month (currently the
 * 14th). The end date is exclusive, so a 14 Oct – 14 Nov tab covers 14 Oct
 * through 13 Nov and the 14th belongs to the next cycle only.
 *
 * Categories live in the "tags" column of the "2Up Spending" tab, with
 * "override" (column H) winning over tags and "split" (column I) winning
 * over amount. Budgets come from the Config tab, which is one row per line
 * item — a category can span several rows and they are summed.
 *
 * The hidden Template tab is rebuilt from Config each time a tab is created,
 * so Config is the source of truth and Template is just a cache.
 *
 * MENU
 *   Create Next Month Tab              — chains on from the latest tab's end date
 *   Enable Auto-Create ...             — does the above automatically each cycle
 */

const SOURCE_SHEET = '2Up Spending';   // tab holding raw synced transactions
const CONFIG_NAME = 'Config';
const TEMPLATE_NAME = 'Template';
const CYCLE_DAY = 14;                  // cycles run from this day of the month to the same day next month

// Shows a UI alert when run manually (menu click); falls back to a log entry
// when run from a time-based trigger, where there's no UI to alert into.
function notify(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log(message);
  }
}

// ---------- Menu ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Budget')
    .addItem('Create Next Month Tab', 'createNextMonthTab')
    .addSeparator()
    .addItem('Enable Auto-Create the day before each cycle', 'createMonthlyTrigger')
    .addToUi();
}

// ---------- Monthly tab creation ----------

function createNextMonthTab() {
  const ss = SpreadsheetApp.getActive();
  const template = ss.getSheetByName(TEMPLATE_NAME);
  if (!template) {
    notify(`No "${TEMPLATE_NAME}" tab found. It is hidden by default — unhide it via View > Show, or restore it before creating a tab.`);
    return;
  }

  // Rebuild the template from Config first, so budget changes are picked up
  buildTemplateLayout(template);

  const { start, end } = getNextCycleDates(ss);
  const tabName = formatTabName(start, end);

  if (ss.getSheetByName(tabName)) {
    notify('Tab already exists: ' + tabName);
    return;
  }

  const newSheet = template.copyTo(ss);
  newSheet.setName(tabName);
  newSheet.showSheet();
  newSheet.getRange('B2').setValue(start);
  newSheet.getRange('D2').setValue(end);

  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(ss.getNumSheets());

  notify('Created tab: ' + tabName);
}

// Chains on from the end date of the most recent cycle tab, rather than from
// today's date — so a skipped or early run still produces the correct next
// cycle instead of drifting.
function getNextCycleDates(ss) {
  let latestEnd = null;

  ss.getSheets().forEach(sh => {
    if (sh.getRange('A1').getValue() !== 'Cycle') return; // only our generated month tabs
    if (sh.getName() === TEMPLATE_NAME) return;
    const endVal = sh.getRange('D2').getValue();
    if (endVal instanceof Date && (!latestEnd || endVal > latestEnd)) {
      latestEnd = endVal;
    }
  });

  let start;
  if (latestEnd) {
    start = new Date(latestEnd);
  } else {
    // Only used if no cycle tab exists at all
    const today = new Date();
    start = new Date(today.getFullYear(), today.getMonth(), CYCLE_DAY);
    if (today.getDate() < CYCLE_DAY) {
      start = new Date(today.getFullYear(), today.getMonth() - 1, CYCLE_DAY);
    }
  }
  const end = new Date(start);
  end.setMonth(end.getMonth() + 1);
  return { start, end };
}

function formatTabName(start, end) {
  const tz = Session.getScriptTimeZone();
  const startStr = Utilities.formatDate(start, tz, 'd MMM');
  const endStr = Utilities.formatDate(end, tz, 'd MMM yyyy');
  return `${startStr} - ${endStr}`;
}

function createMonthlyTrigger() {
  // Removes any existing trigger for this function first, so re-running doesn't duplicate it
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'createNextMonthTab') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('createNextMonthTab')
    .timeBased()
    .onMonthDay(CYCLE_DAY - 1)
    .atHour(20)
    .create();
  notify('Auto-create enabled: a new month tab will be generated on day ' + (CYCLE_DAY - 1) + ' of each month.');
}

// ---------- Template layout ----------

function buildTemplateLayout(sheet) {
  sheet.clear();
  sheet.getRange('A1').setValue('Cycle');
  sheet.getRange('A2').setValue('Start');
  sheet.getRange('B2').setValue(new Date()).setNumberFormat('yyyy-mm-dd');
  sheet.getRange('C2').setValue('End');
  sheet.getRange('D2').setValue(new Date()).setNumberFormat('yyyy-mm-dd');

  sheet.getRange('A3').setValue('Expected Incoming');
  sheet.getRange('B3').setValue(0).setNumberFormat('$#,##0.00');
  sheet.getRange('C3').setValue('Actual Incoming');
  sheet.getRange('D3').setFormula(buildActualIncomingFormula()).setNumberFormat('$#,##0.00');

  const headers = ['Category', 'Budget', 'Remaining', 'Auto Usage', 'Manual Usage', 'Total Usage', '% Used'];
  sheet.getRange('A5:G5').setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(5);

  populateCategoryRows(sheet);
}

function populateCategoryRows(sheet) {
  const categories = readConfigBudgets(); // [[category, summed total], ...]

  // clear any previous category rows before repopulating
  const lastRow = sheet.getLastRow();
  if (lastRow > 5) {
    sheet.getRange(6, 1, lastRow - 5, 7).clearContent();
  }

  let row = 6;
  categories.forEach(([cat, budget]) => {
    sheet.getRange(row, 1).setValue(cat);
    sheet.getRange(row, 2).setValue(budget || 0).setNumberFormat('$#,##0.00');
    sheet.getRange(row, 5).setValue(0).setNumberFormat('$#,##0.00'); // manual usage — you fill this in
    writeCategoryRowFormulas(sheet, row);
    row++;
  });

  const totalRow = row + 1;
  sheet.getRange(totalRow, 1).setValue('TOTAL').setFontWeight('bold');
  ['B', 'C', 'D', 'E', 'F'].forEach(col => {
    sheet.getRange(`${col}${totalRow}`)
      .setFormula(`=SUM(${col}6:${col}${row - 1})`)
      .setNumberFormat('$#,##0.00')
      .setFontWeight('bold');
  });
}

// Column layout (one place, so every writer agrees):
// A Category | B Budget | C Remaining | D Auto Usage | E Manual Usage
// F Total Usage | G % Used
function writeCategoryRowFormulas(sheet, row) {
  sheet.getRange(row, 3).setFormula(`=B${row}-F${row}`).setNumberFormat('$#,##0.00');
  sheet.getRange(row, 4).setFormula(buildAutoUsageFormula(row)).setNumberFormat('$#,##0.00');
  sheet.getRange(row, 6).setFormula(`=D${row}+E${row}`).setNumberFormat('$#,##0.00');
  sheet.getRange(row, 7).setFormula(`=IF(B${row}=0,"",F${row}/B${row})`).setNumberFormat('0%');
}

// ---------- Config ----------
//
// Config layout:  A = Category | B = Line Item | C = Amount
// A category can span as many rows as you like (e.g. several loans);
// the month tab's Budget is the sum of that category's rows.
// Category matching is case-insensitive ("Loans" and "loans" are one group).

// Returns [[categoryName, total], ...] in first-seen order, summing line items.
function readConfigBudgets() {
  const config = SpreadsheetApp.getActive().getSheetByName(CONFIG_NAME);
  if (!config) return [];
  const lastRow = config.getLastRow();
  if (lastRow < 2) return [];

  const data = config.getRange(2, 1, lastRow - 1, 3).getValues();
  const order = [];
  const totals = {};

  data.forEach(row => {
    const name = String(row[0]).trim();
    if (!name) return;
    const amount = Number(row[2]) || 0;
    const key = name.toLowerCase();
    if (!(key in totals)) {
      totals[key] = { name, total: 0 };
      order.push(key);
    }
    totals[key].total += amount;
  });

  return order.map(k => [totals[k].name, Math.round(totals[k].total * 100) / 100]);
}

// ---------- Formula builders ----------

// Source-sheet column range, e.g. srcRange('D').
//
// Wrapped in INDIRECT on purpose. The sync inserts new transactions at the TOP
// of the sheet, and a normal reference anchored at row 2 (e.g. $D$2:$D$20000)
// gets SHIFTED DOWN by Sheets on every insert, so the newest rows silently fall
// outside the range and stop being counted. INDIRECT stores the range as text,
// which Sheets cannot rewrite, so it always means "row 2 to the bottom".
// Open-ended ($D$2:$D) so it grows with the sheet and can never overflow.
function srcRange(col) {
  return `INDIRECT("'${SOURCE_SHEET}'!$${col}$2:$${col}")`;
}

// Effective amount for each transaction: the split column (I) overrides the
// synced amount (D) when it is filled in. The split keeps the transaction's
// original direction, so you can type the portion that is yours as either
// 50 or -50 on a spend row and it counts as spend either way.
function buildEffectiveAmount() {
  const amount = srcRange('D');
  const split = srcRange('I');
  return `IF(${split}<>"",SIGN(${amount})*ABS(${split}),${amount})`;
}

// Effective tag: override (column H) if present, else tags (column G).
function buildEffectiveTags() {
  return `IF(${srcRange('H')}<>"",${srcRange('H')},${srcRange('G')})`;
}

// Builds the Auto Usage formula for a given category row.
// SEARCH is wrapped in commas on both sides so matching is exact-tag,
// not a loose substring match (avoids "car" matching inside
// "car-insurance-and-maintenance", for example). TRIM guards against a
// stray leading/trailing space in a tag breaking that exact match.
function buildAutoUsageFormula(row) {
  const tags = buildEffectiveTags();
  const amt = buildEffectiveAmount();
  const date = srcRange('B');
  return (
    `=SUMPRODUCT((IFERROR(SEARCH(","&TRIM($A${row})&",",","&TRIM(${tags})&","),0)>0)*` +
    `(${date}>=$B$2)*(${date}<$D$2)*` +
    `(${amt}<0)*(-(${amt})))`
  );
}

// Sums positive effective amounts inside the cycle (respects split too).
function buildActualIncomingFormula() {
  const amt = buildEffectiveAmount();
  const date = srcRange('B');
  return (
    `=SUMPRODUCT((${date}>=$B$2)*(${date}<$D$2)*(${amt}>0)*(${amt}))`
  );
}

// ---------- Maintenance (no menu entry — run from the Apps Script editor) ----------

// Re-applies the current Auto Usage and Actual Incoming formulas to every
// existing month tab. Each tab is a static copy of Template made at creation
// time, so it does NOT pick up changes made to the formula builders above.
// Run this from the editor after changing any of them, or the older tabs will
// quietly keep using the formulas they were born with.
function refreshAutoUsageFormulas() {
  const ss = SpreadsheetApp.getActive();
  let tabsUpdated = 0;

  ss.getSheets().forEach(sh => {
    if (sh.getRange('A1').getValue() !== 'Cycle') return; // only our generated tabs (Template included)
    sh.getRange('D3').setFormula(buildActualIncomingFormula());
    const lastRow = sh.getLastRow();
    for (let row = 6; row <= lastRow; row++) {
      const cat = sh.getRange(row, 1).getValue();
      if (!cat || cat === 'TOTAL') break;
      sh.getRange(row, 4).setFormula(buildAutoUsageFormula(row));
    }
    tabsUpdated++;
  });

  notify(`Auto Usage and Actual Incoming formulas refreshed on ${tabsUpdated} tab(s).`);
}
