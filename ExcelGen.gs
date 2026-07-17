function buildReportExcel(report) {
  var info = getCompanyInfo();
  var COLS = 6;
  var ss = SpreadsheetApp.create('Timesheet_Report_' + Date.now());
  var sheet = ss.getSheets()[0].setName('OT Report');

  sheet.getRange(1, 1, 1, COLS).merge().setValue(info.name)
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 1, 1, COLS).merge().setValue('Timesheet OT Report')
    .setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(3, 1, 1, COLS).merge().setValue(report.startDate + ' to ' + report.endDate)
    .setFontSize(9).setFontColor('#6b7280').setHorizontalAlignment('center');

  var headers = ['Employee', 'Normal Hrs', 'OT Hrs', 'Total Hrs', 'Days', 'Entry Range'];
  var headerRow = 5;
  var headerRange = sheet.getRange(headerRow, 1, 1, COLS).setValues([headers]);
  headerRange.setBackground('#fbf0e6').setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center');

  if (report.rows.length > 0) {
    var data = report.rows.map(function (r) {
      return [sanitizeSheetText(r.name), r.normalHours, r.otHours, r.totalHours, r.daysWorked, r.firstEntry + ' - ' + r.lastEntry];
    });
    var dataRange = sheet.getRange(headerRow + 1, 1, data.length, COLS).setValues(data);
    dataRange.setFontSize(9);
    sheet.getRange(headerRow + 1, 2, data.length, 3).setNumberFormat('0.00').setHorizontalAlignment('right');
    sheet.getRange(headerRow + 1, 5, data.length, 1).setNumberFormat('0').setHorizontalAlignment('right');
    sheet.getRange(headerRow, 1, data.length + 1, COLS).setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);
  }

  var lastRow = headerRow + Math.max(report.rows.length, 0);
  if (report.skippedSheets && report.skippedSheets.length > 0) {
    lastRow += 1;
    sheet.getRange(lastRow, 1, 1, COLS).merge()
      .setValue('⚠ ' + report.skippedSheets.length + ' sheet(s) had errors and were excluded.')
      .setFontSize(8).setFontColor('#dc2626');
  }

  sheet.setFrozenRows(headerRow);
  sheet.autoResizeColumns(1, COLS);
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
