/**
 * TIMESHEETS LIST — server-side filtered, searched, and paginated.
 * Reuses the same sheet/row-scanning helpers as ReportCode.gs (findEmployeeBlocks,
 * getHoursBetweenTimes, parseCellDate, isTimesheetSheet) but returns raw per-shift
 * rows instead of per-employee aggregate totals.
 */

/**
 * Raw per-shift rows for every employee across all timesheet sheets touching
 * [startDate, endDate], unfiltered and unpaginated. Shared by
 * listTimesheetEntries (which filters/sorts/paginates on top of this),
 * generateOtReport (which aggregates it into per-employee totals), and
 * buildReportExcel (which groups rows by employee for per-employee tabs).
 *
 * Cached per date range (CacheService, 60s TTL) so the several near-simultaneous
 * calls one page load makes (paginated entries list + OT report preview) share
 * one sheet scan instead of each re-reading every sheet from scratch.
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

  return getCachedTimesheetScan_('scan_' + startDateStr + '_' + endDateStr, function () {
    return scanTimesheetEntries_(startDate, endDate);
  });
}

/** ponytail: 60s TTL over exact write-time invalidation of every date-range key; raise if staleness ever bites. */
var TIMESHEET_SCAN_CACHE_TTL_SEC = 60;
var TIMESHEET_SCAN_CACHE_INDEX_KEY = '_timesheetScanCacheKeys';
// Apps Script caps each cache entry at 100KB; leave headroom for key overhead.
var TIMESHEET_SCAN_CACHE_MAX_BYTES = 95000;

/**
 * Cache-or-compute wrapper around a sheet scan. Mirrors the common Apps Script
 * CacheService pattern (get -> miss -> compute -> put, size-guarded under the
 * 100KB/entry cache limit) so callers don't need to know caching exists.
 */
function getCachedTimesheetScan_(cacheKey, computeFn) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // fall through and recompute on a corrupt/incompatible cache entry
    }
  }

  var fresh = computeFn();

  try {
    var serialized = JSON.stringify(fresh);
    if (serialized.length < TIMESHEET_SCAN_CACHE_MAX_BYTES) {
      cache.put(cacheKey, serialized, TIMESHEET_SCAN_CACHE_TTL_SEC);
      var indexRaw = cache.get(TIMESHEET_SCAN_CACHE_INDEX_KEY);
      var index = indexRaw ? JSON.parse(indexRaw) : [];
      if (index.indexOf(cacheKey) === -1) {
        index.push(cacheKey);
        cache.put(TIMESHEET_SCAN_CACHE_INDEX_KEY, JSON.stringify(index), TIMESHEET_SCAN_CACHE_TTL_SEC);
      }
    } else {
      // Big date ranges silently re-scan every sheet on every call — leave a
      // trace so the pattern is visible in the execution logs.
      Logger.log('Timesheet scan too large to cache (' + serialized.length + ' bytes) for key ' + cacheKey);
    }
  } catch (e) {
    // cache write failures are non-fatal - fresh data is already computed
  }

  return fresh;
}

/** Called right after a save/undo commits so the next read isn't served stale data from within the TTL window. */
function invalidateTimesheetScanCache_() {
  var cache = CacheService.getScriptCache();
  var indexRaw = cache.get(TIMESHEET_SCAN_CACHE_INDEX_KEY);
  if (!indexRaw) return;
  try {
    cache.removeAll(JSON.parse(indexRaw));
  } catch (e) {
    // ignore - worst case the TTL expires it within 60s anyway
  }
  cache.remove(TIMESHEET_SCAN_CACHE_INDEX_KEY);
}

function scanTimesheetEntries_(startDate, endDate) {
  var tz = Session.getScriptTimeZone();
  var startStr = Utilities.formatDate(startDate, tz, 'yyyy-MM-dd');
  var endStr = Utilities.formatDate(endDate, tz, 'yyyy-MM-dd');
  var startKey = startDate.getFullYear() * 12 + startDate.getMonth();
  var endKey = endDate.getFullYear() * 12 + endDate.getMonth();
  var jsRestDays = getRestDaysJs();

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
          if (jsRestDays.indexOf(rowDate.getDay()) !== -1) { // configured rest day, all OT
            normalHours = 0;
            otHours = totalHours;
          } else {
            normalHours = Math.min(CONFIG.NORMAL_HOURS_PER_DAY, totalHours);
            otHours = totalHours > CONFIG.NORMAL_HOURS_PER_DAY ? totalHours - CONFIG.NORMAL_HOURS_PER_DAY : 0;
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
 * @param {number} [filters.pageSize]     default 25, capped at 1000
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
  var pageSize = Math.min(1000, Math.max(1, parseInt(filters.pageSize, 10) || 25));

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

/** Case-insensitive substring match, falling back to per-word edit-distance <=2 for typo tolerance.
 * (levenshtein lives in Shared.gs.) */
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
