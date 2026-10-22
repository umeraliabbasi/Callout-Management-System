/*  Shift Call-Out Management System — backend
 *  Container-bound to the Master Callout Sheet.
 *
 *  SETUP CHECKLIST (see SETUP.md for details):
 *   1. Project Settings → Time zone = "America/New_York"  (REQUIRED — classification depends on it)
 *   2. Ambassador Names tab headers:  A=Name  B=Email  C=Proper Callouts  D=Improper Callouts
 *   3. Add a "Config" tab (key in col A, value in col B) with the emails/phones below.
 *   4. Create the Coverage sheet (separate file) and paste its ID into COVERAGE_SPREADSHEET_ID.
 *   5. Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone with the link.
 */

// ───────────────────────── CONFIG ─────────────────────────
const COVERAGE_SPREADSHEET_ID = 'PASTE_COVERAGE_SHEET_ID';   // the separate Coverage Google Sheet

const TAB = {
  intake:   'All Callouts',
  roster:   'Ambassador Names',
  config:   'Config',
  coverage: 'Coverage',            // tab name inside the Coverage spreadsheet
};

// Roster columns (1-indexed). Names live from row 2 down.
const R = { name:1, email:2, proper:3, improper:4, firstRow:2 };

// Canonical "Type of Shift" strings the master sheet's tabs filter on — do not change.
const TYPE_OF_SHIFT = {
  regular:  'Regularly Scheduled Shift (Most Shifts)',
  training: 'Tour Training Shift',
  marquee:  'Marquee Event Shift',
  invited:  'Invited Shift',
};

// The four canonical "When is Shift" strings (9am day-of boundary). Written to column K.
const WHEN = {
  proper:       'It is more than 72 hours from the start of the shift',
  imp_before10: 'It is less than 72 hours until the start of the shift, but it is before 10:00pm the night before the shift',
  imp_after10:  'It is after 10:00pm the night before the shift, but before 9:00am the day of the shift',
  imp_dayof:    'It is after 9:00am the day of the shift',
};

// Tabs that get the orange/yellow row highlighting.
const HIGHLIGHT_TABS = ['All Callouts','Archived All Callouts','Proper Callouts','Improper Callouts',
  'Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday',
  'Tour Training Callouts','Invited Shift Callouts'];

// Filtered tabs, each populated by a QUERY over All Callouts (columns A–M).
const FILTER_TABS = {
  'Proper Callouts':        "select * where Col11 contains 'more'",
  'Improper Callouts':      "select * where Col11 contains 'less' or Col11 contains 'after'",
  'Monday':                 "select * where Col7 contains 'Monday'",
  'Tuesday':                "select * where Col7 contains 'Tuesday'",
  'Wednesday':              "select * where Col7 contains 'Wednesday'",
  'Thursday':               "select * where Col7 contains 'Thursday'",
  'Friday':                 "select * where Col7 contains 'Friday'",
  'Saturday':               "select * where Col7 contains 'Saturday'",
  'Sunday':                 "select * where Col7 contains 'Sunday'",
  'Tour Training Callouts': "select * where Col5 contains 'Tour Training Shift'",
  'Invited Shift Callouts': "select * where Col5 contains 'Invited Shift'",
};

// ───────────────────────── WEB APP ─────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Shift Call-Out')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Names for the pick-your-name dropdown (read live from the roster). */
function getRosterNames() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.roster);
  const last = sh.getLastRow();
  if (last < R.firstRow) return [];
  const vals = sh.getRange(R.firstRow, R.name, last - R.firstRow + 1, 1).getValues();
  const names = vals.map(r => String(r[0] || '').trim()).filter(Boolean);
  return [...new Set(names)].sort((a, b) => a.localeCompare(b));
}

// ───────────────────────── SUBMIT ─────────────────────────
/**
 * payload = { name, category, shiftType, shiftName, date 'YYYY-MM-DD',
 *             start 'HH:MM', end 'HH:MM', reason, detail }
 * Classification is recomputed here (server, NYC time) — the client value is ignored.
 */
function submitCallout(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const p = payload || {};
  const name = String(p.name || '').trim();
  if (!name) throw new Error('Missing ambassador name.');

  const person = findAmbassador_(name);
  if (!person) throw new Error('Ambassador not found in roster: ' + name);

  const cls = classifyServer_(p.date, p.start);      // proper | imp_before10 | imp_after10 | imp_dayof
  const improper = cls !== 'proper';
  const dow = dayOfWeek_(p.date);
  let shiftName;
  if (p.category === 'marquee')
    shiftName = (String(p.marqueeEvent || '').trim() + (p.shiftName ? ' — ' + String(p.shiftName).trim() : '')).trim();
  else if (p.category === 'invited')
    shiftName = String(p.shiftName || '').trim();
  else  // regular / training ("Other/My shift is not listed" carries its name in shiftName)
    shiftName = (p.shiftType === 'Other/My shift is not listed')
      ? String(p.shiftName || '').trim()
      : String(p.shiftType || '').trim();

  // All Callouts row (columns A–M).
  const row = [
    new Date(),                                   // A Timestamp
    person.email,                                 // B Email Address
    name,                                         // C Name
    'I understand the callout procedure and will follow it accurately',  // D
    TYPE_OF_SHIFT[p.category] || '',              // E Type of Shift
    shiftName,                                    // F Shift
    dow,                                          // G Day
    formatDate_(p.date),                          // H Date
    formatTime12_(p.start),                       // I Start Time
    formatTime12_(p.end),                         // J End Time
    WHEN[cls],                                    // K When is Shift
    String(p.reason || ''),                       // L Why
    String(p.detail || ''),                       // M Details
  ];
  ss.getSheetByName(TAB.intake).appendRow(row);

  // Proper regular/marquee callouts → Coverage sheet (training & invited excluded).
  const coverageCat = (p.category === 'regular' || p.category === 'marquee');
  if (cls === 'proper' && coverageCat && COVERAGE_SPREADSHEET_ID && COVERAGE_SPREADSHEET_ID !== 'PASTE_COVERAGE_SHEET_ID') {
    SpreadsheetApp.openById(COVERAGE_SPREADSHEET_ID).getSheetByName(TAB.coverage)
      .appendRow(['', name, shiftName, dow, formatDate_(p.date), formatTime12_(p.start), formatTime12_(p.end)]);
  }

  // Counters: every callout counts as 1 (proper -> Proper, improper -> Improper).
  const weight = 1;
  bumpCounter_(person.rowIndex, cls === 'proper' ? R.proper : R.improper, weight);

  sendNotifications_({ cls: cls, improper: improper, dow: dow, name: name, email: person.email,
    category: p.category, shiftName: shiftName, marqueeEvent: p.marqueeEvent, date: formatDate_(p.date),
    start: formatTime12_(p.start), end: formatTime12_(p.end) });

  return { ok: true, classification: cls, weight: weight, coverage: cls === 'proper' };
}

// ───────────────────── CLASSIFY (server, NYC time) ─────────────────────
// Requires the project time zone to be America/New_York so wall-clock dates are NYC.
function classifyServer_(dateStr, startStr) {
  const d = String(dateStr).split('-').map(Number);
  const t = String(startStr).split(':').map(Number);
  const start       = new Date(d[0], d[1]-1, d[2],   t[0], t[1], 0);
  const now         = new Date();
  const hrs         = (start - now) / 36e5;
  const nightBefore = new Date(d[0], d[1]-1, d[2]-1, 22, 0, 0);   // 10pm the night before
  const dayOf9      = new Date(d[0], d[1]-1, d[2],    9, 0, 0);   // 9am day-of
  if (hrs > 72)          return 'proper';
  if (now < nightBefore) return 'imp_before10';
  if (now < dayOf9)      return 'imp_after10';
  return 'imp_dayof';
}

// Marquee events routed to GA/XME + Scheduling (all other marquee events → DTS + SOD-if-late).
const MARQUEE_GROUP_A = ['Weekend on the Square', 'Open House: Spotlight Saturday'];
function isMarqueeGroupA_(ev) { return MARQUEE_GROUP_A.indexOf(String(ev || '')) !== -1; }
// Which callouts require the ambassador to CALL the SOD on a day-of: regular/training, or marquee Group B.
function needsSODCall_(cat, marqueeEvent) {
  if (cat === 'regular' || cat === 'training') return true;
  if (cat === 'marquee' && !isMarqueeGroupA_(marqueeEvent)) return true;
  return false;
}

// ───────────────────────── NOTIFICATIONS ─────────────────────────
function sendNotifications_(info) {
  const cfg = readConfig_();
  const fromName = cfg['from name'] || 'Call-Out System';
  const subj = 'Call-out: ' + info.name + ' — ' + info.shiftName + ' (' + info.dow + ' ' + info.date + ')';
  const details =
    'Ambassador: ' + info.name + '\nShift: ' + info.shiftName +
    '\nWhen: ' + info.dow + ' ' + info.date + ', ' + info.start + '–' + info.end +
    '\nClassification: ' + labelFor_(info.cls) + '\n';

  const cat = info.category;
  const improper = info.improper;
  const late = (info.cls === 'imp_after10' || info.cls === 'imp_dayof');
  const dayof = info.cls === 'imp_dayof';

  const dts   = cfg[(info.dow + ' dts').toLowerCase()];
  const sod   = cfg['sod email'];
  const gaxme = cfg['ga/xme email'];
  const sched = cfg['scheduling email'];
  const tsup  = cfg['tour supervisors email'];
  const mk = function (who, line) { return 'Dear ' + who + ',\n\n' + line + '\n\n' + details; };

  // Confirmation to the ambassador, with the category-appropriate required action.
  let action = '';
  if (dayof && needsSODCall_(cat, info.marqueeEvent)) action = '\nIMPORTANT: You must also CALL ' + callContactText_(cfg, info.shiftName) + '.\n';
  else if (cat === 'invited') action = '\nRequired: message the Program Administrator who invited you (or a GA) on Slack.\n';
  safeMail_(info.email, 'Your shift call-out is confirmed',
    'Hi ' + info.name + ',\n\nYour call-out has been recorded.\n\n' + details + action + '\nThis is an automated message.', fromName);

  if (cat === 'regular' || cat === 'training') {
    if (improper && dts) safeMail_(dts, subj, mk(info.dow + ' DTS', 'An improper call-out was submitted.'), fromName);
    if (late) {
      if (sod)   safeMail_(sod,   subj, mk('SOD', 'A short-notice call-out was submitted.'), fromName);
      if (gaxme) safeMail_(gaxme, subj, mk('GA/XME', 'A short-notice call-out was submitted.'), fromName);
    }
    if (cat === 'training' && tsup) safeMail_(tsup, subj, mk('Tour Supervisors', 'A Tour Training call-out was submitted.'), fromName);

  } else if (cat === 'invited') {
    if (improper && dts) safeMail_(dts, subj, mk(info.dow + ' DTS', 'An invited-shift improper call-out was submitted.'), fromName);
    // Ambassador Slacks the inviting administrator (shown in the app + confirmation). No other system emails.

  } else if (cat === 'marquee') {
    if (improper) {
      if (isMarqueeGroupA_(info.marqueeEvent)) {                 // WOTS, Spotlight Saturday
        if (gaxme) safeMail_(gaxme, subj, mk('GA/XME', 'A marquee call-out was submitted.'), fromName);
        if (sched) safeMail_(sched, subj, mk('Scheduling', 'A marquee call-out was submitted.'), fromName);
      } else {                                                   // Senior Day, MOTS/MOTC, LaunchPad, Other
        if (dts) safeMail_(dts, subj, mk(info.dow + ' DTS', 'A marquee call-out was submitted.'), fromName);
        if (late && sod) safeMail_(sod, subj, mk('SOD', 'A short-notice marquee call-out was submitted.'), fromName);
      }
    }
  }
}

function callContactText_(cfg, shiftName) {
  return 'the SOD — Bonomi ' + (cfg['bonomi sod phone'] || '') + ' or 400 ' + (cfg['400 sod phone'] || '');
}

// ───────────────────────── ROSTER / COUNTERS ─────────────────────────
function findAmbassador_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.roster);
  const last = sh.getLastRow();
  if (last < R.firstRow) return null;
  const vals = sh.getRange(R.firstRow, R.name, last - R.firstRow + 1, R.improper).getValues();
  const target = name.toLowerCase();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][R.name-1] || '').trim().toLowerCase() === target) {
      return { rowIndex: R.firstRow + i, email: String(vals[i][R.email-1] || '').trim() };
    }
  }
  return null;
}

function bumpCounter_(rowIndex, col, amount) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.roster);
  const cell = sh.getRange(rowIndex, col);
  cell.setValue((Number(cell.getValue()) || 0) + amount);
}

function readConfig_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.config);
  const map = {};
  if (!sh) return map;
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.config).getDataRange().getValues().forEach(r => {
    const k = String(r[0] || '').trim().toLowerCase();
    if (k) map[k] = String(r[1] || '').trim();
  });
  return map;
}

// ───────────────────────── MENU: reset / highlight ─────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Callout Tools')
    .addItem('Set up / rebuild filter tabs', 'rebuildFilterTabs')
    .addItem('Apply row highlighting', 'applyHighlighting')
    .addItem('Remove row highlighting', 'removeHighlighting')
    .addSeparator()
    .addItem('Archive selected call-outs', 'archiveSelected')
    .addSeparator()
    .addItem('Fix append gap (trim empty rows)', 'trimEmptyRows')
    .addSeparator()
    .addItem('Reset ALL counters to 0', 'resetCounters')
    .addToUi();
}

/**
 * Moves the rows you've selected in the "All Callouts" tab to "Archived All Callouts",
 * then removes them from All Callouts (which also drops them from the filter tabs).
 * Counters are NOT affected. Select whole rows or any cells in the rows you want to move.
 */
function archiveSelected() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  if (sh.getName() !== TAB.intake) {
    ui.alert('Open the "' + TAB.intake + '" tab, select the call-out rows to archive, then run this.');
    return;
  }
  // Collect the selected row numbers (skip the header row 1).
  const ranges = (sh.getActiveRangeList() ? sh.getActiveRangeList().getRanges() : [sh.getActiveRange()]);
  const rowSet = {};
  ranges.forEach(r => { const start = r.getRow(); for (let i = 0; i < r.getNumRows(); i++) { const row = start + i; if (row > 1) rowSet[row] = true; } });
  const rows = Object.keys(rowSet).map(Number).sort((a, b) => a - b);
  if (!rows.length) { ui.alert('No call-out rows selected (don\'t select the header row).'); return; }

  if (ui.alert('Archive ' + rows.length + ' call-out(s)? They move to "Archived All Callouts" and leave All Callouts.',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const arch = ss.getSheetByName('Archived All Callouts');
  if (!arch) { ui.alert('No "Archived All Callouts" tab found.'); return; }

  const data = rows.map(r => sh.getRange(r, 1, 1, 13).getValues()[0]);   // columns A–M
  arch.getRange(arch.getLastRow() + 1, 1, data.length, 13).setValues(data);
  rows.slice().reverse().forEach(r => sh.deleteRow(r));                  // delete bottom-up so indexes hold
  ui.alert('Archived ' + rows.length + ' call-out(s).');
}

/**
 * Fixes new call-outs appending far down the sheet (e.g. row 1008).
 * Copying a sheet can leave blank-but-not-empty rows below the data, which pushes
 * appendRow past them. This finds the last row that actually holds call-out data
 * (columns A–M) and deletes every row below it on All Callouts.
 */
function trimEmptyRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sh = ss.getSheetByName(TAB.intake);
  if (!sh) { ui.alert('No "' + TAB.intake + '" tab found.'); return; }

  // Walk up from the bottom to find the last row with real content in A–M.
  const maxRows = sh.getMaxRows();
  const values = sh.getRange(1, 1, maxRows, 13).getValues();
  let lastData = 1;                                  // keep the header row
  for (let r = values.length - 1; r >= 1; r--) {
    if (values[r].some(v => String(v).trim() !== '')) { lastData = r + 1; break; }
  }

  const toDelete = maxRows - lastData;
  if (toDelete <= 0) { ui.alert('Nothing to trim — data already ends at the last row.'); return; }

  if (ui.alert('Last call-out data is in row ' + lastData + '.\n\nDelete the ' + toDelete +
      ' row(s) below it so new call-outs append correctly?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  sh.deleteRows(lastData + 1, toDelete);
  ui.alert('Trimmed ' + toDelete + ' row(s). New call-outs will now land in row ' + (lastData + 1) + '.');
}

/** Writes the QUERY formula into each filtered tab (run once after import). */
function rebuildFilterTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(FILTER_TABS).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.getRange('A2').setFormula('=QUERY(\'All Callouts\'!A2:M, "' + FILTER_TABS[name] + '", 0)');
  });
  SpreadsheetApp.getUi().alert('Filter tabs wired up. They will fill as call-outs come in.');
}

function resetCounters() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Reset every ambassador\'s Proper and Improper counters to 0?',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.roster);
  const last = sh.getLastRow();
  if (last >= R.firstRow) sh.getRange(R.firstRow, R.proper, last - R.firstRow + 1, 2).setValue(0);
  ui.alert('All counters reset to 0.');
}

/**
 * Orange if the ambassador's improper counter > 5, else yellow if proper > 10 (orange wins).
 *
 * Google Sheets forbids conditional-format formulas that reference another sheet, so each
 * tab gets two hidden helper columns (O = improper, P = proper) that VLOOKUP the roster
 * locally; the format rules then only look at those same-sheet cells.
 */
function applyHighlighting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ORANGE = '#f4b183', YELLOW = '#ffe599';
  const HELP_IMP = 15, HELP_PRO = 16;            // columns O and P

  HIGHLIGHT_TABS.forEach(tabName => {
    const sh = ss.getSheetByName(tabName);
    if (!sh) return;
    const lastRow = Math.max(sh.getMaxRows(), 2);

    // Helper columns: pull this row's counters from the roster by Name (column C).
    sh.getRange(2, HELP_IMP).setFormula(
      '=ARRAYFORMULA(IF(LEN($C2:$C)=0,"",IFERROR(VLOOKUP($C2:$C,\'' + TAB.roster + '\'!$A:$D,4,FALSE),0)))');
    sh.getRange(2, HELP_PRO).setFormula(
      '=ARRAYFORMULA(IF(LEN($C2:$C)=0,"",IFERROR(VLOOKUP($C2:$C,\'' + TAB.roster + '\'!$A:$D,3,FALSE),0)))');
    sh.getRange(1, HELP_IMP).setValue('_impCount');
    sh.getRange(1, HELP_PRO).setValue('_properCount');
    sh.hideColumns(HELP_IMP, 2);

    // Rules reference only same-sheet cells ($O2 / $P2) — allowed.
    const rng = sh.getRange(2, 1, lastRow - 1, 13);              // A2:M
    const ruleO = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($C2<>"",N($O2)>5)')
      .setBackground(ORANGE).setRanges([rng]).build();
    const ruleY = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($C2<>"",N($P2)>10)')
      .setBackground(YELLOW).setRanges([rng]).build();

    // Drop our previous rules (they reference $O2/$P2 or the roster), keep any others.
    const keep = sh.getConditionalFormatRules().filter(r => {
      const bc = r.getBooleanCondition && r.getBooleanCondition();
      const vals = bc ? (bc.getCriteriaValues() || []).join(' ') : '';
      return vals.indexOf('$O2') === -1 && vals.indexOf('$P2') === -1
          && vals.indexOf("'" + TAB.roster + "'") === -1;
    });
    sh.setConditionalFormatRules([ruleO, ruleY].concat(keep));   // orange first → wins
  });
  SpreadsheetApp.getUi().alert('Row highlighting applied. (Helper columns O–P are hidden on each tab.)');
}

/**
 * Undoes applyHighlighting: removes our orange/yellow rules and clears + unhides
 * the O/P helper columns. Any other conditional-format rules on the tab are kept.
 */
function removeHighlighting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Remove the orange/yellow row highlighting from all call-out tabs?',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const HELP_IMP = 15, HELP_PRO = 16;            // columns O and P
  let cleaned = 0;

  HIGHLIGHT_TABS.forEach(tabName => {
    const sh = ss.getSheetByName(tabName);
    if (!sh) return;

    // Keep every rule except the two we created.
    const keep = sh.getConditionalFormatRules().filter(r => {
      const bc = r.getBooleanCondition && r.getBooleanCondition();
      const vals = bc ? (bc.getCriteriaValues() || []).join(' ') : '';
      return vals.indexOf('$O2') === -1 && vals.indexOf('$P2') === -1
          && vals.indexOf("'" + TAB.roster + "'") === -1;
    });
    sh.setConditionalFormatRules(keep);

    // Clear and unhide the helper columns.
    if (sh.getMaxColumns() >= HELP_PRO) {
      sh.getRange(1, HELP_IMP, sh.getMaxRows(), 2).clearContent();
      sh.showColumns(HELP_IMP, 2);
    }
    cleaned++;
  });
  ui.alert('Highlighting removed from ' + cleaned + ' tab(s). Helper columns cleared.');
}

// ───────────────────────── small helpers ─────────────────────────
function labelFor_(cls) {
  return { proper:'Proper (72h+)', imp_before10:'Improper — under 72h',
    imp_after10:'Improper — after 10pm', imp_dayof:'Improper — day-of' }[cls] || cls;
}
function dayOfWeek_(dateStr) {
  const d = String(dateStr).split('-').map(Number);
  return ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'][new Date(d[0], d[1]-1, d[2]).getDay()];
}
function formatDate_(dateStr) {
  const d = String(dateStr).split('-').map(Number);
  return d[1] + '/' + d[2] + '/' + d[0];
}
function formatTime12_(hhmm) {
  if (!hhmm) return '';
  const t = String(hhmm).split(':').map(Number);
  const ap = t[0] >= 12 ? 'PM' : 'AM';
  const h12 = ((t[0] + 11) % 12) + 1;
  return h12 + ':' + String(t[1]).padStart(2, '0') + ' ' + ap;
}
function safeMail_(to, subject, body, fromName) {
  if (!to) return;
  try { MailApp.sendEmail({ to: to, subject: subject, body: body, name: fromName || 'Call-Out System' }); }
  catch (e) { console.error('Mail failed to ' + to + ': ' + e); }
}
