function checkAiCredentials() {
  const accountId = PropertiesService.getScriptProperties().getProperty('CLOUDFLARE_ACCOUNT_ID');
  const apiToken = PropertiesService.getScriptProperties().getProperty('CLOUDFLARE_API_TOKEN');
  return !!(accountId && apiToken);
}

function getRecentEntriesPreview(sheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.DATA_START_ROW) return [];
    var lastCol = sheet.getLastColumn();
    var data = sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, lastCol).getValues();
    var tz = Session.getScriptTimeZone();
    var employeeBlocks = findEmployeeBlocks(sheet);
    var results = [];
    for (var r = data.length - 1; r >= 0 && results.length < 20; r--) {
      var dateVal = data[r][0];
      if (!dateVal) continue;
      var d = parseCellDate(dateVal);
      if (!d) continue;
      var dateStr = Utilities.formatDate(d, tz, 'dd/MM');
      for (var b = 0; b < employeeBlocks.length; b++) {
        var col = employeeBlocks[b].column;
        var name = employeeBlocks[b].name;
        var startCell = data[r][col - 1];
        var endCell = data[r][col];
        var jobCell = data[r][col + 5 - 1];
        if (startCell && endCell) {
          var startStr = typeof startCell === 'string' ? startCell : Utilities.formatDate(startCell, tz, 'HH:mm');
          var endStr = typeof endCell === 'string' ? endCell : Utilities.formatDate(endCell, tz, 'HH:mm');
          results.push({ date: dateStr, name: name, start: startStr, end: endStr, job: jobCell || '' });
        }
      }
    }
    return results;
  } catch (e) {
    Logger.log('getRecentEntriesPreview failed: ' + e);
    return [];
  }
}

/**
 * Distinct job order values across every timesheet sheet, most-used first.
 * Cached (same CacheService pattern as collectTimesheetEntries in
 * EntriesList.gs) since this has no date-range shortcut - every call scans
 * every historical sheet in full, and the entry form re-triggers it on
 * every mount.
 */
function getJobOrdersByFrequency() {
  return getCachedTimesheetScan_('jobOrderFrequency', scanJobOrdersByFrequency_);
}

function scanJobOrdersByFrequency_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const freq = {};

  ss.getSheets().forEach(function (sheet) {
    try {
      if (typeof isTimesheetSheet === 'function' && !isTimesheetSheet(sheet)) return;
      const lastRow = sheet.getLastRow();
      if (lastRow < CONFIG.DATA_START_ROW) return;
      const lastCol = sheet.getLastColumn();
      const blocks = findEmployeeBlocks(sheet);
      if (blocks.length === 0) return;

      const allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      for (let b = 0; b < blocks.length; b++) {
        const colIdx = blocks[b].column + 5 - 1; // JOB ORDER sits 5 columns after this block's START
        for (let r = CONFIG.DATA_START_ROW - 1; r < lastRow; r++) {
          if (allValues[r] && allValues[r][colIdx] !== undefined) {
            const v = allValues[r][colIdx];
            if (v && String(v).trim() !== '') {
              const key = String(v).trim().toUpperCase();
              if (!freq[key]) freq[key] = { name: String(v).trim(), count: 0 };
              freq[key].count++;
            }
          }
        }
      }
    } catch (e) {
      Logger.log('getJobOrdersByFrequency: skipped sheet due to: ' + e);
    }
  });

  return Object.keys(freq)
    .sort(function (a, b) { return freq[b].count - freq[a].count || a.localeCompare(b); })
    .map(function (k) { return freq[k].name; });
}
