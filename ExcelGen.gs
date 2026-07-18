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

/** formatDDMMYYYY (shared with the email body/UI) produces 'dd/MM/yyyy';
 * Excel output uses dashes instead, so this just swaps the separator rather
 * than reparsing the date. */
function toExcelDate_(ddmmyyyySlash) {
  return String(ddmmyyyySlash).replace(/\//g, '-');
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
  sheet.getRange(3, 1, 1, COLS).merge().setValue(toExcelDate_(report.startDate) + ' to ' + toExcelDate_(report.endDate))
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
    return [sanitizeSheetText(r.name), r.normalHours, r.otHours, r.totalHours, r.daysWorked, toExcelDate_(r.firstEntry) + ' - ' + toExcelDate_(r.lastEntry)];
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

    // Both charts anchor to the right of the table (column 11) instead of
    // below it, stacked vertically so they don't need the table's row count.
    var barChart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(headerRow, 1, data.length + 1, 3))
      .setNumHeaders(1)
      .setOption('title', 'Hours by employee — Normal vs OT')
      .setOption('isStacked', true)
      .setOption('legend', { position: 'top' })
      .setOption('width', 640)
      .setOption('height', 320)
      .setPosition(headerRow, 11, 0, 0)
      .build();
    sheet.insertChart(barChart);

    var pieChart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange(headerRow, 8, 3, 2))
      .setNumHeaders(1)
      .setOption('title', 'OT share of total hours')
      .setOption('width', 360)
      .setOption('height', 320)
      .setPosition(headerRow + 17, 11, 0, 0)
      .build();
    sheet.insertChart(pieChart);
  }
}

function buildEmployeeTab_(sheet, row, empEntries, report) {
  var COLS = 8;
  var tz = Session.getScriptTimeZone();

  sheet.getRange(1, 1, 1, COLS).merge().setValue(row.name)
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 1, 1, COLS).merge().setValue(toExcelDate_(report.startDate) + ' to ' + toExcelDate_(report.endDate))
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
    var d = parseIsoDate(e.date);
    var dateLabel = Utilities.formatDate(d, tz, 'dd-MM-yyyy');
    var dayLabel = Utilities.formatDate(d, tz, 'EEE');
    return [dateLabel, dayLabel, e.startTime, e.endTime, sanitizeSheetText(e.jobOrder), e.normalHours, e.otHours, e.totalHours];
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
