/**
 * SmartCheck.gs — AI-powered pre-save validation
 * ----------------------------------------------
 * Analyzes pending entries for anomalies before they reach
 * coreProcessEntries. Reuses callCloudflareAiJson from ai.gs
 * and getContextAndRules-style helpers from Code.gs/UiUtils.gs.
 */

/**
 * Analyzes a batch of pending entries for anomalies.
 * @param {Object} payload  { entries, sheetName }
 *   entries format: [{ date, employeeColumn, startTime, endTime, jobOrder }]
 * @returns {Object} { success, warnings: [...], error? }
 */
function smartCheckEntries(payload) {
  try {
    if (!payload || !payload.entries || !payload.entries.length) {
      return { success: false, error: 'No entries to check.' };
    }
    if (!checkAiCredentials()) {
      return { success: false, error: 'AI not configured.' };
    }

    var sheetName = payload.sheetName;
    var entries = payload.entries;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet "' + sheetName + '" not found.' };

    // Resolve employee column numbers to names
    var blocks = findEmployeeBlocks(sheet);
    var colToName = {};
    blocks.forEach(function (b) { colToName[b.column] = b.name; });

    var enriched = entries.map(function (e) {
      return {
        date: e.date,
        employee: colToName[e.employeeColumn] || ('Column ' + e.employeeColumn),
        startTime: e.startTime,
        endTime: e.endTime,
        jobOrder: e.jobOrder || ''
      };
    });

    var jobOrders = getJobOrderSuggestions().slice(0, 20);
    var tz = Session.getScriptTimeZone();
    var now = new Date();

    var prompt = [
      'You are a timesheet validation assistant. Return ONLY valid JSON.',
      '',
      'CONTEXT:',
      '- Today: ' + Utilities.formatDate(now, tz, 'yyyy-MM-dd') + ' (' + Utilities.formatDate(now, tz, 'EEEE') + ')',
      '- Sheet: "' + sheetName + '"',
      '- Known job orders: ' + JSON.stringify(jobOrders),
      '',
      'ENTRIES:',
      JSON.stringify(enriched, null, 2),
      '',
      'Check for these anomalies:',
      '- long_shift: shift > 14 hours (high severity)',
      '- weekend: Friday entry (rest day in Qatar) (medium severity)',
      '- missing_job: employee has no job order while others on same day do (low severity)',
      '- duplicate_pattern: same employee+hours+job order across 3+ days (low severity)',
      '- anomaly: any other suspicious pattern (severity as appropriate)',
      '',
      'Return: { "warnings": [{ "type": "...", "severity": "low|medium|high", "message": "..." }] }',
      'If nothing is suspicious: { "warnings": [] }'
    ].join('\n');

    var result = callCloudflareAiJson(prompt, 0.1);
    if (!result.success) return { success: false, error: result.error };

    var warnings = Array.isArray(result.data.warnings)
      ? result.data.warnings.filter(function (w) { return w.type && w.severity && w.message; })
      : [];

    return { success: true, warnings: warnings };
  } catch (err) {
    Logger.log('smartCheckEntries error: ' + err);
    return { success: false, error: err.message || String(err) };
  }
}
