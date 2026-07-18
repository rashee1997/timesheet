# Excel Report: Per-Employee Tabs, Summary Charts and Stats — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Turn the OT report's "Attach report as Excel" output from a single totals-table workbook into a multi-tab workbook: a renamed "Summary" tab (existing table + a stat banner + a stacked bar chart + a pie chart) followed by one tab per employee with their full daily entry breakdown for the report's date range.

**Architecture:** Extract the raw per-shift sheet scan already inside `EntriesList.gs`'s `listTimesheetEntries` into a shared `collectTimesheetEntries`, thread those raw entries through `processSendReport` into a rewritten `buildReportExcel(report, entries)` in `ExcelGen.gs`, which builds the Summary tab and one tab per `report.rows` employee — all sharing the same header/border/auto-width styling via a new `styleReportTable` helper so every tab stays visually consistent.

**Tech Stack:** Google Apps Script (`.gs`/`.js`, ES5-style, no build step), `clasp` for push/deploy.

Spec: `docs/superpowers/specs/2026-07-18-excel-report-monthly-tabs-charts-design.md`

---

### Task 1: Extract `collectTimesheetEntries` out of `listTimesheetEntries`

**Files:**
- Modify: `F:\timesheet\EntriesList.gs`

- [ ] **Step 1: Replace the file with the scan extracted into its own function**

Replace the entire contents of `EntriesList.gs` with:

```js
/**
 * TIMESHEETS LIST — server-side filtered, searched, and paginated.
 * Reuses the same sheet/row-scanning helpers as ReportCode.js (findEmployeeBlocks,
 * getHoursBetweenTimes, parseCellDate, isTimesheetSheet) but returns raw per-shift
 * rows instead of per-employee aggregate totals.
 */

/**
 * Raw per-shift rows for every employee across all timesheet sheets touching
 * [startDate, endDate], unfiltered and unpaginated. Shared by
 * listTimesheetEntries (which filters/sorts/paginates on top of this) and
 * buildReportExcel (which groups rows by employee for per-employee tabs).
 *
 * @param {string} startDateStr yyyy-MM-dd
 * @param {string} endDateStr   yyyy-MM-dd
 * @returns {{entries: Array, skippedSheets: string[]}}
 */
function collectTimesheetEntries(startDateStr, endDateStr) {
  var startDate = parseIsoDate(startDateStr);
  var endDate = parseIsoDate(endDateStr);
  if (startDate.getTime() > endDate.getTime()) {
    throw new Error('Start date must be on or before the end date.');
  }

  var tz = Session.getScriptTimeZone();
  var startStr = Utilities.formatDate(startDate, tz, 'yyyy-MM-dd');
  var endStr = Utilities.formatDate(endDate, tz, 'yyyy-MM-dd');
  var startKey = startDate.getFullYear() * 12 + startDate.getMonth();
  var endKey = endDate.getFullYear() * 12 + endDate.getMonth();

  var entries = [];
  var skippedSheets = [];

  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (sheet) {
    var sheetName = '(unknown sheet)';
    try {
      sheetName = sheet.getName();
      if (!isTimesheetSheet(sheet)) return;

      // Skip whole sheets outside the requested range whenever the month/year
      // is parseable from the sheet name - avoids opening every historical tab
      // for a narrow (e.g. current-month) query. Unparseable names are read
      // anyway, matching how generateOtReport already handles them (safe default).
      var detected = parseSheetNameForMonthYear(sheetName);
      if (detected) {
        var sheetKey = detected.year * 12 + detected.monthIndex;
        if (sheetKey < startKey || sheetKey > endKey) return;
      }

      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow < CONFIG.DATA_START_ROW) return;

      var employeeBlocks = findEmployeeBlocks(sheet);
      if (employeeBlocks.length === 0) return;

      var numRows = lastRow - CONFIG.DATA_START_ROW + 1;
      var data = sheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, lastCol).getValues();

      data.forEach(function (row) {
        var rowDate = parseCellDate(row[0]);
        if (!rowDate || isNaN(rowDate.getTime())) return;
        var rowDateStr = Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd');
        if (rowDateStr < startStr || rowDateStr > endStr) return;

        employeeBlocks.forEach(function (emp) {
          var idx = emp.column - 1;
          if (idx + 5 >= row.length) return;

          var startVal = row[idx];
          var endVal = row[idx + 1];
          var jobOrder = row[idx + 5] ? String(row[idx + 5]).trim() : '';

          var totalHours;
          try {
            totalHours = getHoursBetweenTimes(startVal, endVal);
          } catch (timeErr) {
            return;
          }
          if (!totalHours || totalHours <= 0) return;

          var normalHours, otHours;
          if (rowDate.getDay() === 5) { // Friday = rest day, all OT
            normalHours = 0;
            otHours = totalHours;
          } else {
            normalHours = Math.min(8, totalHours);
            otHours = totalHours > 8 ? totalHours - 8 : 0;
          }

          entries.push({
            date: rowDateStr,
            sheetName: sheetName,
            employeeName: emp.name,
            employeeKey: normalizeNameKey(emp.name),
            startTime: formatTimeCell(startVal, tz),
            endTime: formatTimeCell(endVal, tz),
            jobOrder: jobOrder,
            totalHours: round2(totalHours),
            normalHours: round2(normalHours),
            otHours: round2(otHours)
          });
        });
      });
    } catch (sheetErr) {
      Logger.log('collectTimesheetEntries: skipped sheet "' + sheetName + '" due to error: ' + sheetErr);
      skippedSheets.push(sheetName);
    }
  });

  return { entries: entries, skippedSheets: skippedSheets };
}

/**
 * @param {Object} filters
 * @param {string} filters.startDate  yyyy-MM-dd (required)
 * @param {string} filters.endDate    yyyy-MM-dd (required)
 * @param {string[]} [filters.employees]  exact known names to include
 * @param {string} [filters.jobOrder]     exact job order to match (case-insensitive)
 * @param {number} [filters.minHours]
 * @param {number} [filters.maxHours]
 * @param {boolean} [filters.otOnly]
 * @param {string} [filters.search]       fuzzy match against employee name or job order
 * @param {number} [filters.page]         1-indexed, default 1
 * @param {number} [filters.pageSize]     default 25, capped at 100
 */
function listTimesheetEntries(filters) {
  filters = filters || {};
  if (!filters.startDate || !filters.endDate) {
    return { success: false, error: 'Both a start date and end date are required.' };
  }

  var scan;
  try {
    scan = collectTimesheetEntries(filters.startDate, filters.endDate);
  } catch (e) {
    return { success: false, error: e.message };
  }

  var selectedKeys = null;
  if (filters.employees && filters.employees.length > 0) {
    selectedKeys = {};
    filters.employees.forEach(function (name) {
      if (name) selectedKeys[normalizeNameKey(name)] = true;
    });
  }
  var jobOrderFilter = filters.jobOrder ? String(filters.jobOrder).trim().toLowerCase() : '';
  var minHours = typeof filters.minHours === 'number' ? filters.minHours : null;
  var maxHours = typeof filters.maxHours === 'number' ? filters.maxHours : null;
  var otOnly = !!filters.otOnly;
  var search = filters.search ? String(filters.search).trim() : '';

  var page = Math.max(1, parseInt(filters.page, 10) || 1);
  var pageSize = Math.min(100, Math.max(1, parseInt(filters.pageSize, 10) || 25));

  var filtered = scan.entries.filter(function (e) {
    if (selectedKeys && !selectedKeys[e.employeeKey]) return false;
    if (jobOrderFilter && e.jobOrder.toLowerCase() !== jobOrderFilter) return false;
    if (minHours !== null && e.totalHours < minHours) return false;
    if (maxHours !== null && e.totalHours > maxHours) return false;
    if (otOnly && e.otHours <= 0) return false;
    if (search && !fuzzyMatches(search, e.employeeName) && !fuzzyMatches(search, e.jobOrder)) return false;
    return true;
  });

  filtered.sort(function (a, b) {
    if (a.date !== b.date) return a.date < b.date ? 1 : -1; // newest first
    return a.employeeName.localeCompare(b.employeeName);
  });

  var total = filtered.length;
  var totalPages = Math.max(1, Math.ceil(total / pageSize));
  page = Math.min(page, totalPages);
  var pageEntries = filtered.slice((page - 1) * pageSize, page * pageSize).map(function (e) {
    return {
      date: e.date,
      sheetName: e.sheetName,
      employeeName: e.employeeName,
      startTime: e.startTime,
      endTime: e.endTime,
      jobOrder: e.jobOrder,
      totalHours: e.totalHours,
      normalHours: e.normalHours,
      otHours: e.otHours
    };
  });

  return {
    success: true,
    entries: pageEntries,
    total: total,
    page: page,
    pageSize: pageSize,
    totalPages: totalPages,
    skippedSheets: scan.skippedSheets
  };
}

function formatTimeCell(val, tz) {
  if (typeof val === 'string') return val;
  try {
    return Utilities.formatDate(val, tz, 'HH:mm');
  } catch (e) {
    return String(val);
  }
}

/** Case-insensitive substring match, falling back to per-word edit-distance <=2 for typo tolerance. */
function fuzzyMatches(query, target) {
  if (!query || !target) return false;
  var q = String(query).trim().toLowerCase();
  var t = String(target).trim().toLowerCase();
  if (q === '') return false;
  if (t.indexOf(q) !== -1) return true;

  var words = t.split(/\s+/);
  for (var i = 0; i < words.length; i++) {
    if (levenshtein(q, words[i]) <= 2) return true;
  }
  return false;
}

function levenshtein(a, b) {
  var m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  var prev = [];
  for (var j = 0; j <= n; j++) prev[j] = j;
  for (var i = 1; i <= m; i++) {
    var curr = [i];
    for (var j2 = 1; j2 <= n; j2++) {
      var cost = a.charAt(i - 1) === b.charAt(j2 - 1) ? 0 : 1;
      curr[j2] = Math.min(prev[j2] + 1, curr[j2 - 1] + 1, prev[j2 - 1] + cost);
    }
    prev = curr;
  }
  return prev[n];
}
```

Note: `listTimesheetEntries`'s validation previously threw via `parseIsoDate` inside its own try/catch with a raw `e.message`; that behavior is preserved because `collectTimesheetEntries` throws the same way and the caller still catches it and returns `{ success: false, error }`.

- [ ] **Step 2: Push to Apps Script and sanity-check via the editor**

```bash
cd /f/timesheet && clasp push
```

Expected: `Pushed X files.` with no syntax errors reported.

- [ ] **Step 3: Commit**

```bash
git add EntriesList.gs
git commit -m "Extract collectTimesheetEntries out of listTimesheetEntries

Shares the raw per-shift sheet scan with the upcoming per-employee
Excel tabs instead of duplicating the ~70-line scan loop a third time."
```

---

### Task 2: Add `sanitizeSheetName` next to `sanitizeSheetText`

**Files:**
- Modify: `F:\timesheet\Code.js:575-578`

- [ ] **Step 1: Add the helper immediately after `sanitizeSheetText`**

In `Code.js`, find:

```js
function sanitizeSheetText(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
```

Add directly after it:

```js
/**
 * Makes an employee name safe as an Excel/Sheets tab name (Excel's 31-char
 * cap is the binding constraint) and de-dupes against names already used in
 * the same workbook by appending " (2)", " (3)", etc.
 */
function sanitizeSheetName(name, usedNames) {
  var s = String(name || '').replace(/[\[\]*?/\\:]/g, ' ').trim();
  if (s.length > 31) s = s.substring(0, 31).trim();
  if (!s) s = 'Employee';
  var base = s, n = 2;
  while (usedNames[s]) {
    var suffix = ' (' + n + ')';
    s = base.substring(0, 31 - suffix.length) + suffix;
    n++;
  }
  usedNames[s] = true;
  return s;
}
```

- [ ] **Step 2: Push and commit**

```bash
cd /f/timesheet && clasp push
git add Code.js
git commit -m "Add sanitizeSheetName for Excel-safe, de-duped employee tab names"
```

---

### Task 3: Rewrite `ExcelGen.gs` — Summary tab + per-employee tabs + charts

**Files:**
- Modify: `F:\timesheet\ExcelGen.gs` (full rewrite)

- [ ] **Step 1: Replace the entire file**

```js
function setColumnWidthsToFitContent(sheet, numCols, headers, data, startCol) {
  startCol = startCol || 1;
  var PX_PER_CHAR = 7;
  var PADDING_PX = 24;
  var MIN_WIDTH = 60;
  for (var c = 0; c < numCols; c++) {
    var maxLen = String(headers[c]).length;
    for (var r = 0; r < data.length; r++) {
      var v = data[r][c];
      var len = (v === null || v === undefined) ? 0 : String(v).length;
      if (len > maxLen) maxLen = len;
    }
    sheet.setColumnWidth(startCol + c, Math.max(MIN_WIDTH, maxLen * PX_PER_CHAR + PADDING_PX));
  }
}

/** Applies the report's standard header-row + border styling. Used by every
 * tab so Summary and per-employee tables never visually drift apart. */
function styleReportTable(sheet, headerRow, numCols, dataRowCount, startCol) {
  startCol = startCol || 1;
  sheet.getRange(headerRow, startCol, 1, numCols)
    .setBackground('#fbf0e6').setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center');
  if (dataRowCount > 0) {
    sheet.getRange(headerRow, startCol, dataRowCount + 1, numCols)
      .setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);
  }
}

/**
 * Builds the multi-tab OT report workbook: a Summary tab (totals table, stat
 * banner, bar + pie charts) followed by one tab per employee with their full
 * daily entries for the report's date range.
 *
 * @param {Object} report   result of generateOtReport
 * @param {Array} [entries] raw per-shift rows from collectTimesheetEntries,
 *                          used to build the per-employee tabs. Each employee
 *                          in report.rows with no matching entries still gets
 *                          a tab with a "no entries found" note instead of
 *                          being skipped or throwing.
 */
function buildReportExcel(report, entries) {
  entries = entries || [];
  var info = getCompanyInfo();
  var ss = SpreadsheetApp.create('Timesheet_Report_' + Date.now());
  var usedNames = { 'Summary': true };

  var summarySheet = ss.getSheets()[0].setName('Summary');
  buildSummaryTab_(summarySheet, info, report);

  report.rows.forEach(function (row) {
    var empSheet = ss.insertSheet(sanitizeSheetName(row.name, usedNames));
    var key = normalizeNameKey(row.name);
    var empEntries = entries.filter(function (e) { return e.employeeKey === key; });
    buildEmployeeTab_(empSheet, row, empEntries, report);
  });

  SpreadsheetApp.flush();

  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  });
  var safeStart = report.startDate.replace(/[/]/g, '-');
  var safeEnd = report.endDate.replace(/[/]/g, '-');
  var blob = response.getBlob().setName('Timesheet_Report_' + safeStart + '_to_' + safeEnd + '.xlsx');

  DriveApp.getFileById(ss.getId()).setTrashed(true);
  return [blob];
}

function buildSummaryTab_(sheet, info, report) {
  var COLS = 6;
  var normalHours = report.rows.reduce(function (s, r) { return s + r.normalHours; }, 0);
  var otHours = report.rows.reduce(function (s, r) { return s + r.otHours; }, 0);
  var totalHours = normalHours + otHours;
  var otShare = totalHours > 0 ? (otHours / totalHours) * 100 : 0;

  sheet.getRange(1, 1, 1, COLS).merge().setValue(info.name)
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 1, 1, COLS).merge().setValue('Timesheet OT Report')
    .setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(3, 1, 1, COLS).merge().setValue(report.startDate + ' to ' + report.endDate)
    .setFontSize(9).setFontColor('#6b7280').setHorizontalAlignment('center');

  var statLine = 'Total Hours: ' + totalHours.toFixed(2) +
    '   |   OT Hours: ' + otHours.toFixed(2) + ' (' + otShare.toFixed(1) + '%)' +
    '   |   Active Employees: ' + report.rows.length +
    '   |   Days Covered: ' + report.rangeCalendarDays;
  sheet.getRange(4, 1, 1, COLS).merge().setValue(statLine)
    .setFontSize(10).setFontWeight('bold').setFontColor('#6b7280').setHorizontalAlignment('center');

  var headerRow = 6;
  var headers = ['Employee', 'Normal Hrs', 'OT Hrs', 'Total Hrs', 'Days', 'Entry Range'];
  sheet.getRange(headerRow, 1, 1, COLS).setValues([headers]);

  var data = report.rows.map(function (r) {
    return [sanitizeSheetText(r.name), r.normalHours, r.otHours, r.totalHours, r.daysWorked, r.firstEntry + ' - ' + r.lastEntry];
  });
  if (data.length > 0) {
    sheet.getRange(headerRow + 1, 1, data.length, COLS).setValues(data).setFontSize(9);
    sheet.getRange(headerRow + 1, 2, data.length, 3).setNumberFormat('0.00').setHorizontalAlignment('right');
    sheet.getRange(headerRow + 1, 5, data.length, 1).setNumberFormat('0').setHorizontalAlignment('right');
  }
  styleReportTable(sheet, headerRow, COLS, data.length, 1);

  var lastRow = headerRow + data.length;
  if (report.skippedSheets && report.skippedSheets.length > 0) {
    lastRow += 1;
    sheet.getRange(lastRow, 1, 1, COLS).merge()
      .setValue('⚠ ' + report.skippedSheets.length + ' sheet(s) had errors and were excluded.')
      .setFontSize(8).setFontColor('#dc2626');
  }

  sheet.setFrozenRows(headerRow);
  setColumnWidthsToFitContent(sheet, COLS, headers, data, 1);

  // Charts need at least one employee row - an empty report gets no chart
  // source table and no charts, just the (empty) totals table above.
  if (data.length > 0) {
    var pieHeaders = ['Split', 'Hours'];
    var pieData = [['Normal', round2(normalHours)], ['OT', round2(otHours)]];
    sheet.getRange(headerRow, 8, 1, 2).setValues([pieHeaders]);
    sheet.getRange(headerRow + 1, 8, 2, 2).setValues(pieData);
    styleReportTable(sheet, headerRow, 2, pieData.length, 8);
    setColumnWidthsToFitContent(sheet, 2, pieHeaders, pieData, 8);

    var barChart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(headerRow, 1, data.length + 1, 3))
      .setNumHeaders(1)
      .setOption('title', 'Hours by employee — Normal vs OT')
      .setOption('isStacked', true)
      .setOption('legend', { position: 'top' })
      .setOption('width', 640)
      .setOption('height', 320)
      .setPosition(lastRow + 3, 1, 0, 0)
      .build();
    sheet.insertChart(barChart);

    var pieChart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange(headerRow, 8, 3, 2))
      .setNumHeaders(1)
      .setOption('title', 'OT share of total hours')
      .setOption('width', 360)
      .setOption('height', 320)
      .setPosition(lastRow + 3, 8, 0, 0)
      .build();
    sheet.insertChart(pieChart);
  }
}

function buildEmployeeTab_(sheet, row, empEntries, report) {
  var COLS = 8;
  var tz = Session.getScriptTimeZone();

  sheet.getRange(1, 1, 1, COLS).merge().setValue(row.name)
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 1, 1, COLS).merge().setValue(report.startDate + ' to ' + report.endDate)
    .setFontSize(9).setFontColor('#6b7280').setHorizontalAlignment('center');

  var headerRow = 4;
  var headers = ['Date', 'Day', 'Start', 'End', 'Job Order', 'Normal', 'OT', 'Total'];
  sheet.getRange(headerRow, 1, 1, COLS).setValues([headers]);

  var sorted = empEntries.slice().sort(function (a, b) {
    return a.date < b.date ? -1 : a.date > b.date ? 1 : 0;
  });

  if (sorted.length === 0) {
    sheet.getRange(headerRow + 1, 1, 1, COLS).merge()
      .setValue('No entries found for this range.')
      .setFontSize(9).setFontColor('#6b7280').setHorizontalAlignment('center');
    styleReportTable(sheet, headerRow, COLS, 0, 1);
    sheet.setFrozenRows(headerRow);
    setColumnWidthsToFitContent(sheet, COLS, headers, [], 1);
    return;
  }

  var data = sorted.map(function (e) {
    var dayLabel = Utilities.formatDate(parseIsoDate(e.date), tz, 'EEE');
    return [e.date, dayLabel, e.startTime, e.endTime, sanitizeSheetText(e.jobOrder), e.normalHours, e.otHours, e.totalHours];
  });
  sheet.getRange(headerRow + 1, 1, data.length, COLS).setValues(data).setFontSize(9);
  sheet.getRange(headerRow + 1, 6, data.length, 3).setNumberFormat('0.00').setHorizontalAlignment('right');

  var totalsRow = headerRow + 1 + data.length;
  sheet.getRange(totalsRow, 1, 1, 5).merge().setValue('Total (' + row.daysWorked + ' day(s))')
    .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('right');
  sheet.getRange(totalsRow, 6, 1, 3)
    .setValues([[row.normalHours, row.otHours, row.totalHours]])
    .setFontWeight('bold').setFontSize(9).setNumberFormat('0.00').setHorizontalAlignment('right');

  // Border box covers header + data + totals row.
  styleReportTable(sheet, headerRow, COLS, data.length + 1, 1);
  sheet.setFrozenRows(headerRow);
  setColumnWidthsToFitContent(sheet, COLS, headers, data, 1);
}
```

- [ ] **Step 2: Push to Apps Script**

```bash
cd /f/timesheet && clasp push
```

Expected: `Pushed X files.` with no syntax errors.

- [ ] **Step 3: Commit**

```bash
git add ExcelGen.gs
git commit -m "Rewrite buildReportExcel: Summary tab charts/stats + per-employee tabs

Summary tab keeps the existing totals table, renamed from 'OT Report',
and adds a stat banner plus a stacked bar chart (Normal vs OT per
employee) and a pie chart (org-wide OT share). One additional tab per
employee lists their full daily entries for the report's date range,
sharing the same header/border/auto-width styling via the new
styleReportTable helper so every tab stays visually consistent."
```

---

### Task 4: Wire raw entries into `processSendReport`

**Files:**
- Modify: `F:\timesheet\ReportCode.js:492-501`

- [ ] **Step 1: Update the `attachPdf` block**

Find:

```js
  var attachments = [];
  if (payload.attachPdf) {
    try {
      attachments = buildReportExcel(report);
    } catch (xlsxErr) {
      Logger.log('Excel generation failed: ' + xlsxErr);
      throw new Error('Excel generation failed: ' + (xlsxErr.message || xlsxErr) +
        '. The report was generated but the Excel file could not be created. Uncheck "Attach report as Excel" and try again.');
    }
  }
```

Replace with:

```js
  var attachments = [];
  if (payload.attachPdf) {
    try {
      var scan = collectTimesheetEntries(payload.startDate, payload.endDate);
      attachments = buildReportExcel(report, scan.entries);
    } catch (xlsxErr) {
      Logger.log('Excel generation failed: ' + xlsxErr);
      throw new Error('Excel generation failed: ' + (xlsxErr.message || xlsxErr) +
        '. The report was generated but the Excel file could not be created. Uncheck "Attach report as Excel" and try again.');
    }
  }
```

- [ ] **Step 2: Push to Apps Script**

```bash
cd /f/timesheet && clasp push
```

Expected: `Pushed X files.` with no syntax errors.

- [ ] **Step 3: Commit**

```bash
git add ReportCode.js
git commit -m "Pass raw entries into buildReportExcel for per-employee tabs"
```

---

### Task 5: Manual verification (no Apps Script test harness exists in this repo)

**Files:** none — verification only.

- [ ] **Step 1: Send a real report with the Excel attachment**

From the deployed web app's `/reports` page (or the Sheets-bound Report dialog), pick a date range covering multiple employees with entries, check "Attach report as Excel", and send to a test address you control.

- [ ] **Step 2: Open the downloaded `.xlsx` and check the Summary tab**

- Tab is named "Summary" (not "OT Report").
- Stat banner numbers (Total Hours, OT Hours + %, Active Employees, Days Covered) match a manual sum of the totals table below it.
- Bar chart shows one bar-pair (Normal/OT, stacked) per employee, legend visible.
- Pie chart shows the Normal/OT split matching the stat banner's OT %.
- All columns are auto-fit: no clipped employee names, no excessively wide columns.

- [ ] **Step 3: Check the employee tabs**

- One tab per employee that appears in the Summary table, named after them (truncated + de-duped if needed).
- Row count matches that employee's `Days` column from the Summary table.
- Totals row at the bottom matches that employee's Normal/OT/Total from the Summary table.
- Columns are auto-fit, same visual style (header shading, borders, font size) as the Summary table.

- [ ] **Step 4: Check the empty-range edge case**

Send a report for a date range with zero entries (with the "send anyway" confirmation). Confirm it still sends successfully with just a Summary tab (empty totals table, no stat banner chart-source table, no charts, no employee tabs) — no thrown error.

- [ ] **Step 5: Deploy the live Web App version**

```bash
cd /f/timesheet && clasp deployments
```

Find the production deployment ID (matches the currently live `/exec` URL used by `F:\timesheet-web`), then:

```bash
clasp deploy -i <productionDeploymentId> -d "Per-employee Excel tabs + Summary charts/stats"
```

`clasp push` alone does not update the live Web App — this step is required for the Sheets-bound send flow and the `/reports` page's Excel attachment to actually use the new code.
