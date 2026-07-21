/**
 * DashboardAI.gs — AI-powered dashboard narrative
 * ------------------------------------------------
 * Generates a natural-language summary comparing the current period
 * to the prior period. Reuses generateOtReport (ReportCode.gs),
 * collectTimesheetEntries (EntriesList.gs), and callCloudflareAiJson
 * (ai.gs).
 */

/**
 * @param {Object} payload  { startDate, endDate }  yyyy-MM-dd
 * @returns {Object}
 *   { success, narrative: { summary, highlights, comparisons, anomalies } | null, error? }
 */
function getDashboardNarrative(payload) {
  try {
    if (!checkAiCredentials()) {
      return { success: true, narrative: null };
    }

    var startDate = payload.startDate;
    var endDate = payload.endDate;
    if (!startDate || !endDate) {
      return { success: false, error: 'Both startDate and endDate are required.' };
    }

    // Current period summary
    var currentReport = generateOtReport(startDate, endDate, []);
    var currentStats = buildPeriodStats(currentReport);

    // Previous period (same-length window before startDate)
    var rangeDays = dayDiff(startDate, endDate) + 1;
    var prevEnd = shiftDate(startDate, -1);
    var prevStart = shiftDate(prevEnd, -rangeDays + 1);
    var prevStats = null;
    try {
      var prevReport = generateOtReport(prevStart, prevEnd, []);
      prevStats = buildPeriodStats(prevReport);
    } catch (e) {
      Logger.log('DashboardAI: no prior period data (' + e + ')');
    }

    var prompt = [
      'You are a timesheet analytics assistant. Compare two periods and return ONLY valid JSON.',
      '',
      'CURRENT PERIOD (' + startDate + ' to ' + endDate + '):',
      JSON.stringify(currentStats, null, 2),
      '',
      prevStats
        ? 'PREVIOUS PERIOD (' + prevStart + ' to ' + prevEnd + '):\n' + JSON.stringify(prevStats, null, 2)
        : 'PREVIOUS PERIOD: No data available.',
      '',
      'Return: {',
      '  "summary": "1-2 sentence overview",',
      '  "highlights": ["notable point 1", "notable point 2"],',
      '  "comparisons": ["vs last period: up/down X% in Y"],',
      '  "anomalies": ["anything unusual"]',
      '}',
      'Max 5 bullets total across all arrays. Empty arrays allowed. Skip comparisons if no prior data.'
    ].join('\n');

    var result = callCloudflareAiJson(prompt, 0.3);
    if (!result.success) return { success: true, narrative: null };

    return {
      success: true,
      narrative: {
        summary: result.data.summary || '',
        highlights: coerceArray(result.data.highlights),
        comparisons: coerceArray(result.data.comparisons),
        anomalies: coerceArray(result.data.anomalies)
      }
    };
  } catch (err) {
    Logger.log('getDashboardNarrative error: ' + err);
    return { success: false, error: err.message || String(err) };
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function buildPeriodStats(report) {
  var rows = report.rows || [];
  var totalHours = 0, normalHours = 0, otHours = 0, totalDays = 0;
  rows.forEach(function (r) {
    totalHours += r.totalHours;
    normalHours += r.normalHours;
    otHours += r.otHours;
    totalDays += r.daysWorked;
  });

  var sorted = rows.slice().sort(function (a, b) { return b.otHours - a.otHours; });
  var topOt = sorted.slice(0, 3).map(function (r) {
    return { name: r.name, otHours: Math.round(r.otHours * 100) / 100 };
  });

  return {
    employeeCount: rows.length,
    totalHours: round2(totalHours),
    normalHours: round2(normalHours),
    otHours: round2(otHours),
    otShare: totalHours > 0 ? round2((otHours / totalHours) * 100) : 0,
    avgPerEmployee: rows.length > 0 ? round2(totalHours / rows.length) : 0,
    topOtEmployees: topOt
  };
}

function dayDiff(startStr, endStr) {
  var s = new Date(startStr + 'T00:00:00');
  var e = new Date(endStr + 'T00:00:00');
  return Math.round((e.getTime() - s.getTime()) / 86400000);
}

function shiftDate(dateStr, offsetDays) {
  var d = new Date(dateStr + 'T00:00:00');
  d.setDate(d.getDate() + offsetDays);
  var y = d.getFullYear();
  var m = String(d.getMonth() + 1).padStart(2, '0');
  var day = String(d.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + day;
}

function coerceArray(v) {
  return Array.isArray(v) ? v.filter(function (x) { return typeof x === 'string' && x.trim(); }) : [];
}
