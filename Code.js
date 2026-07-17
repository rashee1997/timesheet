/**
 * TIMESHEET ENTRY SYSTEM (HYBRID BULK VERSION) — v2
 * --------------------------------------------------
 * Writes to whichever month tab is ACTIVE when you open it.
 *
 * WHAT CHANGED IN v2:
 *  - Single + Bulk now share ONE validation/write core (coreProcessEntries),
 *    so both modes get identical protection.
 *  - ALL warnings (month mismatch, long shift, overnight, overwrite) are
 *    collected and shown in ONE combined confirmation instead of three
 *    sequential round-trips. This also fixes the old bug where confirming
 *    one warning type could re-trigger another in a loop.
 *  - Hard errors are collected across ALL entries and reported together,
 *    instead of aborting on the first bad entry.
 *  - Zero-duration shifts (start == end) are rejected — previously they
 *    silently became 24-hour shifts.
 *  - Shifts that cross midnight now surface a soft "overnight" warning
 *    (they used to be assumed silently).
 *  - In-batch duplicate detection: same employee + same date twice in one
 *    submission is now an error (the second used to silently clobber the first).
 *  - Overwrite check now looks at the WHOLE 6-column employee block, not
 *    just the START cell.
 *  - Employee columns coming from the client are validated server-side
 *    (must sit on a block boundary with a real name in the header row).
 *  - All validation reads happen from ONE bulk read of the sheet — no more
 *    per-entry getRange().getValue() calls.
 *  - Writes are wrapped in LockService so two people saving at once can't
 *    interleave.
 *  - getJobOrderSuggestions() now only scans real timesheet sheets and
 *    de-duplicates case-insensitively.
 *
 * SHEET LAYOUT ASSUMPTIONS (unchanged):
 *   Row 1 (HEADER_ROW)      -> Employee name, merged across 6 columns
 *   Row 2 (SUBHEADER_ROW)   -> "TIME" / "HOURS" merged labels (not read by code)
 *   Row 3 (LABEL_ROW)       -> START | END | TOTAL HRS | NORMAL | OT | JOB ORDER
 *   Row 4+ (DATA_START_ROW) -> daily data, column A = DATE
 *   Each employee occupies 6 columns starting at their START column.
 */

const CONFIG = {
  HEADER_ROW: 1,           // row with merged employee names
  LABEL_ROW: 3,            // row with START / END / TOTAL HRS / NORMAL / OT / JOB ORDER
  DATA_START_ROW: 4,       // first row of actual daily data
  DATE_COL: 1,             // column A
  COLS_PER_EMPLOYEE: 6,    // START, END, TOTAL HRS, NORMAL, OT, JOB ORDER
  MAX_REASONABLE_HOURS: 16 // shifts longer than this need explicit confirmation
};

// Full name + short name for every month, used for the soft sanity check
const MONTHS = [
  { full: 'JANUARY', short: 'JAN' },
  { full: 'FEBRUARY', short: 'FEB' },
  { full: 'MARCH', short: 'MAR' },
  { full: 'APRIL', short: 'APR' },
  { full: 'MAY', short: 'MAY' },
  { full: 'JUNE', short: 'JUN' },
  { full: 'JULY', short: 'JUL' },
  { full: 'AUGUST', short: 'AUG' },
  { full: 'SEPTEMBER', short: 'SEP' },
  { full: 'OCTOBER', short: 'OCT' },
  { full: 'NOVEMBER', short: 'NOV' },
  { full: 'DECEMBER', short: 'DEC' }
];

// ---------------------------------------------------------------------
// MENU / UI ENTRY POINT
// ---------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⏱ Timesheet')
    .addItem('Add Time Entry', 'showTimeEntryForm')
    .addItem('Send Timesheet Report', 'showReportForm')
    .addToUi();
}

function showTimeEntryForm() {
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const template = HtmlService.createTemplateFromFile('Form');
  template.activeSheetName = activeSheet.getName();
  const html = template.evaluate().setWidth(520).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Time Entry');
}

// ---------------------------------------------------------------------
// SERVER FUNCTIONS CALLED FROM THE FORM
// ---------------------------------------------------------------------

/**
 * Returns the employee list read live from row 1 of the active sheet.
 */
function getEmployeesForSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { error: 'The sheet "' + sheetName + '" no longer exists. Close this form and reopen it from the tab you want.' };
  }

  const employees = findEmployeeBlocks(sheet);

  if (employees.length === 0) {
    return {
      error: 'No employee columns found in "' + sheet.getName() + '" - row ' + CONFIG.LABEL_ROW + ' needs "START" directly above each employee\'s first column.'
    };
  }

  return { sheetName: sheet.getName(), employees: employees };
}

/**
 * Finds every real employee block by reading row LABEL_ROW ("START" marks
 * a block's first column) instead of assuming a rigid COLS_PER_EMPLOYEE
 * grid starting at a fixed column. A block added slightly off-grid (an
 * extra/missing column earlier on the sheet) is still found correctly,
 * and this is the single source of truth shared with the save-time
 * validation in coreProcessEntries and the report scan in ReportCode.js.
 */
function findEmployeeBlocks(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < CONFIG.DATE_COL + 1) return [];

  const headerRow = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const labelRow = sheet.getRange(CONFIG.LABEL_ROW, 1, 1, lastCol).getDisplayValues()[0];

  const blocks = [];
  for (let i = CONFIG.DATE_COL; i < lastCol; i++) {
    if (String(labelRow[i] || '').trim().toUpperCase() !== 'START') continue;
    const name = String(headerRow[i] || '').trim();
    if (name) blocks.push({ name: name, column: i + 1 });
  }
  return blocks;
}

/**
 * Distinct job order values already used, sorted by frequency (most-used first).
 * v2: only scans sheets that structurally look like timesheets
 * (isTimesheetSheet lives in Report.gs — same project, same global scope),
 * and de-duplicates case-insensitively (keeps first-seen casing).
 * v3: frequency-sorted via getJobOrdersByFrequency() in UiUtils.gs.
 */
function getJobOrderSuggestions() {
  return getJobOrdersByFrequency();
}

// ---------------------------------------------------------------------
// SUBMISSION HANDLERS (both funnel into coreProcessEntries)
// ---------------------------------------------------------------------

/** Single Time Entry submission handler. */
function submitTimeEntry(formData) {
  try {
    return processSubmission(formData);
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function processSubmission(formData) {
  if (!formData || !formData.sheetName) throw new Error('Missing target sheet - please close and reopen the form.');
  if (!formData.date) throw new Error('Please select a date.');
  if (!formData.startTime || !formData.endTime) throw new Error('Please enter both a start time and an end time.');
  if (!formData.employeeColumns || formData.employeeColumns.length === 0) {
    throw new Error('Please select at least one employee.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(formData.sheetName);
  if (!sheet) throw new Error('The sheet "' + formData.sheetName + '" no longer exists.');

  // Expand the multi-employee selection into flat entries and reuse the
  // exact same core as bulk mode.
  const entries = formData.employeeColumns.map(function (col) {
    return {
      date: formData.date,
      employeeColumn: col,
      startTime: formData.startTime,
      endTime: formData.endTime,
      jobOrder: formData.jobOrder
    };
  });

  return coreProcessEntries(sheet, entries, { confirmWarnings: !!formData.confirmWarnings, actorEmail: formData.actorEmail });
}

/** Hybrid bulk submission handler. */
function submitBulkTimeEntries(formData) {
  try {
    return processBulkSubmission(formData);
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function processBulkSubmission(formData) {
  if (!formData || !formData.sheetName) throw new Error('Missing target sheet - please close and reopen the form.');
  if (!formData.entries || formData.entries.length === 0) {
    throw new Error('Please add at least one shift entry.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(formData.sheetName);
  if (!sheet) throw new Error('The sheet "' + formData.sheetName + '" no longer exists.');

  return coreProcessEntries(sheet, formData.entries, { confirmWarnings: !!formData.confirmWarnings, actorEmail: formData.actorEmail });
}

// ---------------------------------------------------------------------
// SHARED VALIDATION + WRITE CORE
// ---------------------------------------------------------------------

/**
 * Validates and writes a flat list of entries:
 *   { date: 'yyyy-mm-dd', employeeColumn, startTime: 'HH:mm', endTime: 'HH:mm', jobOrder }
 *
 * Behavior:
 *  - Hard errors (bad date, bad column, zero duration, in-batch duplicates,
 *    date not found) are collected for ALL entries and thrown together.
 *  - Soft warnings (month mismatch, long shift, overnight, overwrite) are
 *    collected and returned as ONE combined confirmation. The client
 *    resubmits with flags.confirmWarnings = true to proceed.
 *  - Writes happen under a document lock.
 */
function coreProcessEntries(sheet, rawEntries, flags) {
  flags = flags || {};

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW) {
    throw new Error('Sheet "' + sheet.getName() + '" has no data rows below row ' + (CONFIG.DATA_START_ROW - 1) + '.');
  }

  // ---- ONE bulk read for everything: header, dates, conflict cells ----
  const allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headerRow = allValues[CONFIG.HEADER_ROW - 1];

  const tz = Session.getScriptTimeZone();
  const dateMap = {}; // 'yyyy-MM-dd' -> sheet row number
  for (let r = CONFIG.DATA_START_ROW - 1; r < lastRow; r++) {
    const d = parseCellDate(allValues[r][CONFIG.DATE_COL - 1]);
    if (d && !isNaN(d.getTime())) {
      dateMap[Utilities.formatDate(d, tz, 'yyyy-MM-dd')] = r + 1;
    }
  }

  const detected = parseSheetNameForMonthYear(sheet.getName());
  const labelRow = allValues[CONFIG.LABEL_ROW - 1];
  const validStartCols = {}; // column -> true, wherever LABEL_ROW says "START"
  for (let i = CONFIG.DATE_COL; i < lastCol; i++) {
    if (String(labelRow[i] || '').trim().toUpperCase() === 'START') validStartCols[i + 1] = true;
  }

  const errors = [];
  const warnMonth = [];
  const warnLong = [];
  const warnOvernight = [];
  const warnOverwrite = [];
  const seenKeys = {}; // in-batch duplicate detection: dateKey|empCol
  const validated = [];
  let totalHoursSum = 0;

  rawEntries.forEach(function (raw, i) {
    const label = 'Entry #' + (i + 1);

    // Date
    let date;
    try {
      if (!raw.date) throw new Error('date is missing');
      date = parseIsoDate(raw.date);
    } catch (e) {
      errors.push(label + ': invalid or missing date.');
      return;
    }

    // Times
    let start, end;
    try {
      if (!raw.startTime || !raw.endTime) throw new Error('start/end missing');
      start = parseTimeString(raw.startTime);
      end = parseTimeString(raw.endTime);
    } catch (e) {
      errors.push(label + ': enter valid start and end times.');
      return;
    }

    // Employee column: must be a real "START" column and have a name in the header
    const empCol = parseInt(raw.employeeColumn, 10);
    const onBoundary = !isNaN(empCol) && !!validStartCols[empCol];
    const headerName = onBoundary ? String(headerRow[empCol - 1] || '').trim() : '';
    if (!onBoundary || !headerName) {
      errors.push(label + ': employee column ' + raw.employeeColumn + ' is not valid on this sheet — the layout may have changed. Close and reopen the form.');
      return;
    }

    // Duration
    let diffMs = end.getTime() - start.getTime();
    let overnight = false;
    if (diffMs === 0) {
      errors.push(label + ' (' + headerName + '): start and end times are identical (' + raw.startTime + ') — probably a typo.');
      return;
    }
    if (diffMs < 0) {
      diffMs += 24 * 60 * 60 * 1000;
      overnight = true;
    }
    const totalHrs = diffMs / (1000 * 60 * 60);

    // Target row
    const dateKey = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
    const targetRow = dateMap[dateKey];
    if (!targetRow) {
      errors.push(label + ': date ' + formatDDMMYYYY(date) + ' not found in column A of "' + sheet.getName() + '".');
      return;
    }

    // In-batch duplicate
    const dupKey = dateKey + '|' + empCol;
    if (seenKeys[dupKey]) {
      errors.push(label + ': ' + headerName + ' already has another entry for ' + formatDDMMYYYY(date) + ' in this same batch — remove one of them.');
      return;
    }
    seenKeys[dupKey] = true;

    // Soft warnings
    if (detected && (detected.monthIndex !== date.getMonth() || detected.year !== date.getFullYear())) {
      warnMonth.push('  • ' + headerName + ' on ' + formatDDMMYYYY(date));
    }
    if (totalHrs > CONFIG.MAX_REASONABLE_HOURS) {
      warnLong.push('  • ' + headerName + ' on ' + formatDDMMYYYY(date) + ' — ' + totalHrs.toFixed(1) + ' hrs');
    }
    if (overnight) {
      warnOvernight.push('  • ' + headerName + ' on ' + formatDDMMYYYY(date) + ' — ' + raw.startTime + ' → ' + raw.endTime + ' (next day)');
    }

    // Overwrite check across the WHOLE 6-column block (in memory)
    const rowVals = allValues[targetRow - 1];
    let hasExisting = false;
    for (let c = 0; c < CONFIG.COLS_PER_EMPLOYEE; c++) {
      const v = rowVals[empCol - 1 + c];
      if (v !== '' && v !== null && v !== undefined) { hasExisting = true; break; }
    }
    if (hasExisting) {
      warnOverwrite.push('  • ' + headerName + ' on ' + formatDDMMYYYY(date));
    }

    let jobOrder = raw.jobOrder ? String(raw.jobOrder).trim() : '';
    if (jobOrder.length > 200) {
      errors.push(label + ' (' + headerName + '): job order text is too long (max 200 characters).');
      return;
    }
    jobOrder = sanitizeSheetText(jobOrder);

    totalHoursSum += totalHrs;
    validated.push({
      targetRow: targetRow,
      employeeColumn: empCol,
      startTime: start,
      endTime: end,
      jobOrder: jobOrder
    });
  });

  // ---- Hard errors: report them ALL at once ----
  if (errors.length > 0) {
    throw new Error('Fix these before saving:\n' + errors.map(function (e) { return '• ' + e; }).join('\n'));
  }

  // ---- Soft warnings: ONE combined confirmation ----
  const sections = [];
  if (warnMonth.length > 0) {
    sections.push('📅 Outside the "' + sheet.getName() + '" month window:\n' + warnMonth.join('\n'));
  }
  if (warnLong.length > 0) {
    sections.push('⏱ Unusually long shifts (over ' + CONFIG.MAX_REASONABLE_HOURS + ' hrs):\n' + warnLong.join('\n'));
  }
  if (warnOvernight.length > 0) {
    sections.push('🌙 Crosses midnight — double-check these aren\u2019t AM/PM typos:\n' + warnOvernight.join('\n'));
  }
  if (warnOverwrite.length > 0) {
    sections.push('⚠ Existing entries will be OVERWRITTEN:\n' + warnOverwrite.join('\n'));
  }

  if (sections.length > 0 && !flags.confirmWarnings) {
    return {
      success: false,
      needsConfirmation: true,
      confirmType: 'combined',
      message: sections.join('\n\n') + '\n\nSave anyway?'
    };
  }

  // ---- Write under a document lock ----
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    throw new Error('The sheet is busy right now (another save may be in progress). Wait a few seconds and try again — nothing was written.');
  }
  var undoEntries = [];
  try {
    validated.forEach(function (entry) {
      var prevValues = [];
      try {
        prevValues = sheet.getRange(entry.targetRow, entry.employeeColumn, 1, CONFIG.COLS_PER_EMPLOYEE).getValues()[0];
      } catch (e) { prevValues = []; }
      undoEntries.push({ targetRow: entry.targetRow, employeeColumn: entry.employeeColumn, prevValues: prevValues });
      writeFormulaRow(sheet, entry.targetRow, entry.employeeColumn, entry.startTime, entry.endTime, entry.jobOrder);
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  saveUndoData(sheet.getName(), undoEntries);
  validated.forEach(function (entry) {
    logAudit('write', sheet.getName(), entry.targetRow, getEmployeeNameAtColumn(sheet, entry.employeeColumn),
      entry.startTime + '-' + entry.endTime + ' ' + (entry.jobOrder || ''), flags.actorEmail);
  });

  return {
    success: true,
    message: 'Saved ' + validated.length + ' entr' + (validated.length === 1 ? 'y' : 'ies') +
      ' (' + totalHoursSum.toFixed(1) + ' hrs total) on "' + sheet.getName() + '".'
  };
}

/**
 * Builds the TOTAL/NORMAL/OT formulas for one employee's daily row.
 * Shared by writeFormulaRow (on save) and onEdit (auto-revert guard in
 * Guard.gs) so the two can't drift out of sync.
 */
function buildHourFormulas(targetRow, col) {
  var dateCell = getColumnLetter(CONFIG.DATE_COL) + targetRow;
  var startCell = getColumnLetter(col) + targetRow;
  var endCell = getColumnLetter(col + 1) + targetRow;
  var totalCell = getColumnLetter(col + 2) + targetRow;
  var normalCell = getColumnLetter(col + 3) + targetRow;
  var isFriday = 'WEEKDAY(' + dateCell + ',2)=5'; // Friday = rest day, full OT, no Normal

  return {
    total: '=IF(OR(ISBLANK(' + startCell + '), ISBLANK(' + endCell + ')), "", ' + endCell + '-' + startCell + ' + IF(' + endCell + '<' + startCell + ', 1, 0))',
    normal: '=IF(' + totalCell + '="", "", IF(' + isFriday + ', 0, MIN(TIME(8,0,0), ' + totalCell + ')))',
    ot: '=IF(' + totalCell + '="", "", IF(' + isFriday + ', ' + totalCell + ', ' + totalCell + '-' + normalCell + '))'
  };
}

/**
 * Universal Formula Writer for a single employee block.
 */
function writeFormulaRow(sheet, targetRow, col, start, end, jobOrder) {
  var f = buildHourFormulas(targetRow, col);
  var rowValues = [start, end, f.total, f.normal, f.ot, jobOrder];
  sheet.getRange(targetRow, col, 1, CONFIG.COLS_PER_EMPLOYEE).setValues([rowValues]);
  // Force consistent formats every save so a stray manual edit (e.g. cell set to
  // Text/General) upstream can't throw off how times/durations render or sum.
  sheet.getRange(targetRow, col, 1, 2).setNumberFormat('h:mm AM/PM');
  sheet.getRange(targetRow, col + 2, 1, 3).setNumberFormat('[h]:mm');
  sheet.getRange(targetRow, col + 5, 1, 1).setNumberFormat('@');
}

/**
 * Lists every timesheet-shaped sheet with its detected month/year, so a
 * standalone client (no "active sheet" concept) can auto-match a sheet to
 * a date the user picks instead of offering a manual tab picker.
 */
function listTimesheetSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = [];
  ss.getSheets().forEach(function (sheet) {
    if (typeof isTimesheetSheet === 'function' && !isTimesheetSheet(sheet)) return;
    var detected = parseSheetNameForMonthYear(sheet.getName());
    sheets.push({
      name: sheet.getName(),
      monthIndex: detected ? detected.monthIndex : null,
      year: detected ? detected.year : null
    });
  });
  return sheets;
}

// ---------------------------------------------------------------------
// MONTH-NAME PARSING
// ---------------------------------------------------------------------

function parseSheetNameForMonthYear(sheetName) {
  const upper = String(sheetName).toUpperCase();
  let monthIndex = -1;
  let matchedText = '';

  for (let i = 0; i < MONTHS.length && monthIndex === -1; i++) {
    if (upper.indexOf(MONTHS[i].full) !== -1) {
      monthIndex = i;
      matchedText = MONTHS[i].full;
    }
  }
  if (monthIndex === -1) {
    for (let i = 0; i < MONTHS.length && monthIndex === -1; i++) {
      if (upper.indexOf(MONTHS[i].short) !== -1) {
        monthIndex = i;
        matchedText = MONTHS[i].short;
      }
    }
  }

  let year = null;
  if (monthIndex !== -1) {
    const leftoverDigits = upper.replace(matchedText, '').replace(/[^0-9]/g, '');
    year = extractYearFromDigits(leftoverDigits);
  } else {
    const numeric = upper.match(/(\d{1,4})\s*[\/\-. ]\s*(\d{1,4})/);
    if (numeric) {
      const guess = resolveNumericMonthYear(numeric[1], numeric[2]);
      if (guess) {
        monthIndex = guess.monthIndex;
        year = guess.year;
      }
    }
  }

  if (monthIndex === -1 || year === null) return null;
  return { monthIndex: monthIndex, year: year };
}

function extractYearFromDigits(digits) {
  if (digits.length >= 4) {
    const m = digits.match(/\d{4}/);
    if (m) return parseInt(m[0], 10);
  }
  if (digits.length >= 2) {
    const yy = parseInt(digits.slice(0, 2), 10);
    if (!isNaN(yy)) return yy <= 49 ? 2000 + yy : 1900 + yy;
  }
  return null;
}

function resolveNumericMonthYear(a, b) {
  const numA = parseInt(a, 10), numB = parseInt(b, 10);
  let monthNum = null, yearNum = null;

  if (numA >= 1 && numA <= 12 && (b.length === 4 || numB > 12)) {
    monthNum = numA; yearNum = numB;
  } else if (numB >= 1 && numB <= 12 && (a.length === 4 || numA > 12)) {
    monthNum = numB; yearNum = numA;
  } else {
    return null;
  }

  const year = yearNum < 100 ? (yearNum <= 49 ? 2000 + yearNum : 1900 + yearNum) : yearNum;
  return { monthIndex: monthNum - 1, year: year };
}

// ---------------------------------------------------------------------
// HELPERS
// ---------------------------------------------------------------------

/**
 * Neutralizes leading =, +, -, @ so free-text saved into a cell can never be
 * interpreted as a formula (formula/DDE injection via job order text, which
 * later also flows into the Excel export).
 */
function sanitizeSheetText(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function getColumnLetter(col) {
  var letter = '';
  while (col > 0) {
    var temp = (col - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    col = Math.floor((col - temp - 1) / 26);
  }
  return letter;
}

function parseIsoDate(dateString) {
  const date = new Date(dateString + 'T00:00:00');
  if (isNaN(date.getTime())) throw new Error('The date entered is invalid.');
  return date;
}

function parseTimeString(t) {
  const parts = String(t).split(':');
  const h = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10);
  if (isNaN(h) || isNaN(m) || h < 0 || h > 23 || m < 0 || m > 59) {
    throw new Error('Invalid time value: ' + t);
  }
  return new Date(1899, 11, 30, h, m, 0);
}

function findRowForDate(sheet, date) {
  const tz = Session.getScriptTimeZone();
  const targetStr = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return -1;

  const dateValues = sheet
    .getRange(CONFIG.DATA_START_ROW, CONFIG.DATE_COL, lastRow - CONFIG.DATA_START_ROW + 1, 1)
    .getValues();

  for (let i = 0; i < dateValues.length; i++) {
    const cell = parseCellDate(dateValues[i][0]);
    if (cell && Utilities.formatDate(cell, tz, 'yyyy-MM-dd') === targetStr) {
      return CONFIG.DATA_START_ROW + i;
    }
  }
  return -1;
}

function parseCellDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  if (typeof value === 'string' && value.trim() !== '') {
    const s = value.trim();
    let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (m) {
      const day = parseInt(m[1], 10), month = parseInt(m[2], 10), year = parseInt(m[3], 10);
      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        return new Date(year, month - 1, day);
      }
    }
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) {
      return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
    }
  }
  return null;
}

function getEmployeeNameAtColumn(sheet, col) {
  const lastCol = sheet.getLastColumn();
  const headerRow = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0];
  for (let c = col - 1; c >= 0; c--) {
    if (headerRow[c] && String(headerRow[c]).trim() !== '') return String(headerRow[c]).trim();
  }
  return 'Column ' + col;
}

function formatMonthYear(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM yyyy');
}

function formatDDMMYYYY(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}