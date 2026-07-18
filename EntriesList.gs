/**
 * TIMESHEETS LIST — server-side filtered, searched, and paginated.
 * Reuses the same sheet/row-scanning helpers as ReportCode.js (findEmployeeBlocks,
 * getHoursBetweenTimes, parseCellDate, isTimesheetSheet) but returns raw per-shift
 * rows instead of per-employee aggregate totals.
 */

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

  var startDate, endDate;
  try {
    startDate = parseIsoDate(filters.startDate);
    endDate = parseIsoDate(filters.endDate);
  } catch (e) {
    return { success: false, error: e.message };
  }
  if (startDate.getTime() > endDate.getTime()) {
    return { success: false, error: 'Start date must be on or before the end date.' };
  }

  var tz = Session.getScriptTimeZone();
  var startStr = Utilities.formatDate(startDate, tz, 'yyyy-MM-dd');
  var endStr = Utilities.formatDate(endDate, tz, 'yyyy-MM-dd');
  var startKey = startDate.getFullYear() * 12 + startDate.getMonth();
  var endKey = endDate.getFullYear() * 12 + endDate.getMonth();

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
      Logger.log('listTimesheetEntries: skipped sheet "' + sheetName + '" due to error: ' + sheetErr);
      skippedSheets.push(sheetName);
    }
  });

  var filtered = entries.filter(function (e) {
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
    skippedSheets: skippedSheets
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
