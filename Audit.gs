var AUDIT_SHEET_NAME = '_AuditLog';
var AUDIT_HEADERS = ['Timestamp', 'User', 'Action', 'SheetName', 'TargetRow', 'Employee', 'Details'];

function ensureAuditSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(AUDIT_SHEET_NAME);
  if (sheet) return sheet;
  sheet = ss.insertSheet(AUDIT_SHEET_NAME);
  sheet.getRange(1, 1, 1, AUDIT_HEADERS.length).setValues([AUDIT_HEADERS]);
  sheet.getRange(1, 1, 1, AUDIT_HEADERS.length).setFontWeight('bold');
  sheet.setColumnWidths(1, AUDIT_HEADERS.length, 120);
  sheet.hideSheet();
  return sheet;
}

function logAudit(action, sheetName, targetRow, employee, details, actorEmail) {
  try {
    var sheet = ensureAuditSheet_();
    var user = actorEmail || Session.getActiveUser().getEmail() || 'unknown';
    sheet.appendRow([new Date(), user, action, sheetName, targetRow || '', employee || '', details || '']);
  } catch (e) {
    Logger.log('Audit log failed: ' + e);
  }
}

function getRecentAuditEntries(limit) {
  limit = limit || 50;
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AUDIT_SHEET_NAME);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var startRow = Math.max(2, lastRow - limit + 1);
    var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, AUDIT_HEADERS.length).getValues();
    return data.map(function (r) {
      return { timestamp: r[0], user: r[1], action: r[2], sheetName: r[3], targetRow: r[4], employee: r[5], details: r[6] };
    }).reverse();
  } catch (e) {
    Logger.log('getRecentAuditEntries failed: ' + e);
    return [];
  }
}
