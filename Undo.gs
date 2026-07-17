var UNDO_PROPERTY_KEY = 'lastWriteUndoData';
var UNDO_TTL_MS = 60 * 1000; // 1 minute

function saveUndoData(sheetName, entries) {
  var data = {
    timestamp: Date.now(),
    sheetName: sheetName,
    entries: entries.map(function (e) {
      return { row: e.targetRow, col: e.employeeColumn, prevValues: e.prevValues || null };
    })
  };
  PropertiesService.getScriptProperties().setProperty(UNDO_PROPERTY_KEY, JSON.stringify(data));
}

function getUndoData() {
  var raw = PropertiesService.getScriptProperties().getProperty(UNDO_PROPERTY_KEY);
  if (!raw) return null;
  var data = JSON.parse(raw);
  if (Date.now() - data.timestamp > UNDO_TTL_MS) {
    PropertiesService.getScriptProperties().deleteProperty(UNDO_PROPERTY_KEY);
    return null;
  }
  return data;
}

function undoLastSave(actorEmail) {
  var data = getUndoData();
  if (!data) return { success: false, error: 'No recent save to undo (undo expires after 1 minute).' };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(data.sheetName);
    if (!sheet) return { success: false, error: 'Sheet "' + data.sheetName + '" no longer exists.' };

    var restored = 0;
    data.entries.forEach(function (entry) {
      if (entry.prevValues && entry.prevValues.length > 0) {
        sheet.getRange(entry.row, entry.col, 1, entry.prevValues.length).setValues([entry.prevValues]);
        restored++;
      }
    });

    SpreadsheetApp.flush();
    PropertiesService.getScriptProperties().deleteProperty(UNDO_PROPERTY_KEY);
    logAudit('undo', data.sheetName, null, null, 'Undid ' + restored + ' entry blocks', actorEmail);
    return { success: true, message: 'Undid last save (' + restored + ' entries restored).' };
  } catch (e) {
    return { success: false, error: 'Undo failed: ' + e.message };
  }
}
