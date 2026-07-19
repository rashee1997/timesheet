/**
 * TIMESHEET OT REPORT — v2
 * -------------------------
 * Scans every sheet that structurally looks like a timesheet (checked via
 * isTimesheetSheet), reads the actual Date values in column A, and sums up
 * hours for whichever rows fall inside the date range you pick - regardless
 * of what any tab happens to be called or how many tabs a range spans.
 *
 * Shares CONFIG, MONTHS, parseIsoDate(), formatDDMMYYYY(), etc. from Code.gs.
 */

// ---------------------------------------------------------------------
// DIALOG ENTRY POINT
// ---------------------------------------------------------------------

/**
 * Opens the report dialog. Doesn't depend on which sheet is active -
 * a report can span multiple month tabs.
 */
function showReportForm() {
  const html = HtmlService.createTemplateFromFile('ReportForm')
    .evaluate()
    .setWidth(480)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Send Timesheet Report');
}

/**
 * Default range: the 25th of last month to the 25th of this month, matching
 * a typical 26th-to-25th payroll cycle.
 */
function getDefaultReportRange() {
  const today = new Date();
  const end = new Date(today.getFullYear(), today.getMonth(), 25);
  const start = new Date(today.getFullYear(), today.getMonth() - 1, 25);
  const tz = Session.getScriptTimeZone();
  return {
    start: Utilities.formatDate(start, tz, 'yyyy-MM-dd'),
    end: Utilities.formatDate(end, tz, 'yyyy-MM-dd')
  };
}

// ---------------------------------------------------------------------
// SHEET DETECTION
// ---------------------------------------------------------------------

/** A sheet "looks like" a timesheet if its label row matches our known layout - cheap structural check, not a name check. */
function isTimesheetSheet(sheet) {
  try {
    const lastCol = sheet.getLastColumn();
    if (lastCol < 6) return false;

    const labels = sheet
      .getRange(CONFIG.LABEL_ROW, 1, 1, lastCol)
      .getDisplayValues()[0]
      .join(' ')
      .toUpperCase();

    return (
      labels.indexOf('START') !== -1 &&
      labels.indexOf('END') !== -1 &&
      labels.indexOf('NORMAL') !== -1
    );
  } catch (e) {
    Logger.log('isTimesheetSheet failed for sheet "' + (sheet && sheet.getName ? sheet.getName() : '?') + '": ' + e);
    return false;
  }
}

/** Returns a list of all unique employee names across all timesheets. Never throws - returns [] on failure. */
function getAllEmployees() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const found = {};

    ss.getSheets().forEach(function (sheet) {
      try {
        if (!isTimesheetSheet(sheet)) return;
        findEmployeeBlocks(sheet).forEach(function (emp) {
          found[emp.name] = true;
        });
      } catch (innerErr) {
        Logger.log('getAllEmployees: skipped sheet "' + sheet.getName() + '" due to error: ' + innerErr);
      }
    });

    return Object.keys(found).sort(function (a, b) { return a.localeCompare(b); });
  } catch (e) {
    Logger.log('getAllEmployees failed entirely: ' + e);
    return [];
  }
}

// ---------------------------------------------------------------------
// CORE REPORT BUILDER
// ---------------------------------------------------------------------

/**
 * Core report builder: scans every timesheet-shaped sheet, sums Normal/OT/Total
 * hours per employee for rows whose date falls within [startDateStr, endDateStr]
 * (inclusive), by reading the actual date in column A - not by matching tab names.
 * Throws on genuinely invalid input; skips/logs per-sheet problems instead of
 * failing the whole report.
 */
function generateOtReport(startDateStr, endDateStr, selectedEmployees) {
  if (!startDateStr || !endDateStr) {
    throw new Error('Both a start date and end date are required.');
  }

  const startDate = parseIsoDate(startDateStr);
  const endDate = parseIsoDate(endDateStr);

  if (!startDate || isNaN(startDate.getTime())) {
    throw new Error('Start date "' + startDateStr + '" could not be parsed.');
  }
  if (!endDate || isNaN(endDate.getTime())) {
    throw new Error('End date "' + endDateStr + '" could not be parsed.');
  }
  if (startDate.getTime() > endDate.getTime()) {
    throw new Error('Start date must be on or before the end date.');
  }

  const tz = Session.getScriptTimeZone();
  const startStr = Utilities.formatDate(startDate, tz, 'yyyy-MM-dd');
  const endStr = Utilities.formatDate(endDate, tz, 'yyyy-MM-dd');

  // Reuses the same (cached) sheet scan as listTimesheetEntries instead of
  // independently re-reading every sheet - see EntriesList.gs.
  const scan = collectTimesheetEntries(startDateStr, endDateStr);
  const totals = {};
  const skippedSheets = scan.skippedSheets;
  // Latest date (yyyy-MM-dd) where ANY employee - filtered or not - has an
  // actual logged entry, across every timesheet sheet.
  let latestDataDateStr = null;

  const selectedKeys = {};
  if (selectedEmployees && selectedEmployees.length > 0) {
    selectedEmployees.forEach(function (name) {
      if (name) selectedKeys[normalizeNameKey(name)] = true;
    });
  }
  const hasFilter = Object.keys(selectedKeys).length > 0;

  scan.entries.forEach(function (e) {
    // Data-frontier check runs over EVERY employee, not just the filtered set.
    if (!latestDataDateStr || e.date > latestDataDateStr) {
      latestDataDateStr = e.date;
    }

    if (hasFilter && !selectedKeys[e.employeeKey]) return;

    if (!totals[e.employeeKey]) {
      totals[e.employeeKey] = {
        displayName: e.employeeName,
        normal: 0,
        ot: 0,
        days: 0,
        firstDate: e.date,
        lastDate: e.date
      };
    }

    const t = totals[e.employeeKey];
    t.normal += e.normalHours;
    t.ot += e.otHours;
    t.days++;
    if (e.date < t.firstDate) t.firstDate = e.date;
    if (e.date > t.lastDate) t.lastDate = e.date;
  });

  const rangeCalendarDays = Math.round(
    (endDate.getTime() - startDate.getTime()) / (24 * 60 * 60 * 1000)
  ) + 1;

  // ---- Future-date-aware "expected days" calculation ----
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const rangeIncludesFuture = endStr > todayStr;
  let effectiveEndStr = rangeIncludesFuture ? todayStr : endStr;

  let dataLagNote = null;
  if (latestDataDateStr === null) {
    // Nobody - filtered or not - has any entry at all in this range yet.
    effectiveEndStr = null;
  } else if (latestDataDateStr < effectiveEndStr) {
    dataLagNote = 'No timesheet data found anywhere past ' + formatDDMMYYYY(parseIsoDate(latestDataDateStr)) +
      ' yet, so partial-coverage checks stop there instead of today - entries may just not be filled in yet.';
    effectiveEndStr = latestDataDateStr;
  }

  let expectedCalendarDays = 0;
  if (effectiveEndStr && effectiveEndStr >= startStr) {
    const effectiveEndDate = parseIsoDate(effectiveEndStr);
    expectedCalendarDays = Math.round(
      (effectiveEndDate.getTime() - startDate.getTime()) / (24 * 60 * 60 * 1000)
    ) + 1;
  }

  // Partial-coverage flag: proportional to the period length rather than a flat
  // day-count cushion, so it doesn't over-flag long ranges (a few normal leave
  // days across 6 months tripped the old flat "-10" cushion) or fire on ranges
  // too short/just-started to carry any real signal.
  const MIN_EVALUABLE_DAYS = 10;        // shorter/less-elapsed than this: not enough signal, never flag
  const ABSENCE_TOLERANCE_RATIO = 0.30; // tolerate up to ~30% of the period as normal days off
  const MIN_ABSOLUTE_GAP_DAYS = 5;      // even if the ratio math says "flag", require a real gap this big

  const rows = Object.keys(totals)
    .sort(function (a, b) { return totals[a].displayName.localeCompare(totals[b].displayName); })
    .map(function (key) {
      const t = totals[key];
      const gap = expectedCalendarDays - t.days;
      return {
        name: t.displayName,
        normalHours: round2(t.normal),
        otHours: round2(t.ot),
        totalHours: round2(t.normal + t.ot),
        daysWorked: t.days,
        firstEntry: formatDDMMYYYY(parseIsoDate(t.firstDate)),
        lastEntry: formatDDMMYYYY(parseIsoDate(t.lastDate)),
        partialCoverage: expectedCalendarDays >= MIN_EVALUABLE_DAYS &&
          gap >= MIN_ABSOLUTE_GAP_DAYS &&
          t.days < expectedCalendarDays * (1 - ABSENCE_TOLERANCE_RATIO)
      };
    });

  let futureRangeNote = null;
  if (rangeIncludesFuture) {
    futureRangeNote = expectedCalendarDays > 0
      ? 'This range includes dates after today (' + formatDDMMYYYY(parseIsoDate(todayStr)) +
        '), so this period isn\u2019t complete yet. Partial-coverage flags only consider data through today.'
      : 'This entire date range is in the future (today is ' + formatDDMMYYYY(parseIsoDate(todayStr)) +
        ') - there is no data to report yet.';
  }

  return {
    startDate: formatDDMMYYYY(startDate),
    endDate: formatDDMMYYYY(endDate),
    rangeCalendarDays: rangeCalendarDays,
    rangeIncludesFuture: rangeIncludesFuture,
    futureRangeNote: futureRangeNote,
    dataLagNote: dataLagNote,
    rows: rows,
    skippedSheets: skippedSheets
  };
}

/** Collapses casing/spacing differences so the same person merges across sheets. */
function normalizeNameKey(name) {
  return String(name).trim().toUpperCase().replace(/\s+/g, ' ');
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

// ---------------------------------------------------------------------
// CLIENT-FACING WRAPPERS (preview / send)
// ---------------------------------------------------------------------

/**
 * Client-facing wrapper for previewing the report inline before sending.
 */
function previewOtReport(startDateStr, endDateStr, selectedEmployees) {
  try {
    const report = generateOtReport(startDateStr, endDateStr, selectedEmployees);
    return { success: true, report: report };
  } catch (err) {
    Logger.log('previewOtReport error: ' + err);
    return { success: false, error: err.message || String(err) };
  }
}

// ---------------------------------------------------------------------
// EMAIL RECIPIENT SEARCH
// ---------------------------------------------------------------------

/**
 * Auto-fetches recipients matching query, aggregating from:
 *  1. Personal Google Contacts.
 *  2. Domain directory profiles via the People API Advanced service.
 *  3. AdminDirectory API (Directory SDK) fallback.
 */
function searchContacts(query) {
  const results = [];
  const seen = {};
  if (!query) return results;
  const q = String(query).trim().toLowerCase();
  if (!q) return results;

  // 1. Personal Google Contacts
  try {
    const contacts = ContactsApp.getContacts();
    for (let i = 0; i < contacts.length && results.length < 8; i++) {
      const c = contacts[i];
      const fullName = (c.getGivenName() + ' ' + c.getFamilyName()).trim();
      const nameMatches = fullName.toLowerCase().indexOf(q) !== -1;
      const emails = c.getEmails();
      for (let j = 0; j < emails.length; j++) {
        const email = emails[j].getAddress();
        if (!email || seen[email]) continue;
        const emailMatches = email.toLowerCase().indexOf(q) !== -1;
        if (nameMatches || emailMatches) {
          seen[email] = true;
          results.push({ name: fullName || email, email: email });
        }
      }
    }
  } catch (e) {
    Logger.log('searchContacts: Contacts service unavailable: ' + e);
  }

  // 2. Organization Directory search via the People API Advanced service
  try {
    if (typeof People !== 'undefined' && results.length < 12) {
      const peopleSearch = People.People.searchDirectoryPeople({
        query: q,
        readMask: 'names,emailAddresses',
        sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'],
        pageSize: 10
      });
      if (peopleSearch && peopleSearch.people) {
        peopleSearch.people.forEach(function (person) {
          const nameObj = person.names && person.names[0];
          const emailObj = person.emailAddresses && person.emailAddresses[0];
          if (emailObj && emailObj.value && !seen[emailObj.value] && results.length < 12) {
            seen[emailObj.value] = true;
            results.push({ name: nameObj ? nameObj.displayName : emailObj.value, email: emailObj.value });
          }
        });
      }
    }
  } catch (e) {
    Logger.log('searchContacts: People API advanced service unavailable: ' + e);
  }

  // 3. Fallback: Admin SDK Directory API
  try {
    if (typeof AdminDirectory !== 'undefined' && results.length < 12) {
      const domain = Session.getActiveUser().getEmail().split('@')[1];
      const safeQ = q.replace(/[^a-z0-9 ]/gi, '');
      const page = AdminDirectory.Users.list({ domain: domain, query: 'name:' + safeQ + '*', maxResults: 8 });
      if (page && page.users) {
        page.users.forEach(function (u) {
          const email = u.primaryEmail;
          if (email && !seen[email] && results.length < 12) {
            seen[email] = true;
            results.push({ name: u.name ? u.name.fullName : email, email: email });
          }
        });
      }
    }
  } catch (e) {
    Logger.log('searchContacts: Directory API advanced service unavailable: ' + e);
  }

  return results;
}

// ---------------------------------------------------------------------
// SEND REPORT
// ---------------------------------------------------------------------

const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const MAX_RECIPIENTS = 50; // sane upper bound to avoid fat-fingered mass sends

function sendOtReport(payload) {
  try {
    return processSendReport(payload);
  } catch (err) {
    Logger.log('sendOtReport error: ' + err);
    return { success: false, error: err.message || String(err) };
  }
}

function processSendReport(payload) {
  if (!payload || !payload.startDate || !payload.endDate) {
    throw new Error('Please select a date range.');
  }
  if (!payload.emails || payload.emails.length === 0) {
    throw new Error('Please add at least one recipient email.');
  }
  if (payload.emails.length > MAX_RECIPIENTS) {
    throw new Error('Too many recipients (' + payload.emails.length + '). Max is ' + MAX_RECIPIENTS + ' per send.');
  }

  // De-dupe defensively even though the UI shouldn't allow duplicates.
  const uniqueEmails = Array.from(new Set(payload.emails.map(function (e) { return String(e).trim(); })));

  const invalid = uniqueEmails.filter(function (e) { return !EMAIL_PATTERN.test(e); });
  if (invalid.length > 0) {
    throw new Error('These don\'t look like valid email addresses: ' + invalid.join(', '));
  }

  const report = generateOtReport(payload.startDate, payload.endDate, payload.selectedEmployees);

  if (report.rows.length === 0 && !payload.confirmEmpty) {
    const baseMsg = 'No timesheet entries found between ' + report.startDate + ' and ' + report.endDate + '.';
    const extraNote = report.futureRangeNote || report.dataLagNote || '';
    const message = extraNote
      ? baseMsg + ' ' + extraNote + ' Send an empty report anyway?'
      : baseMsg + ' Send an empty report anyway?';
    return {
      success: false,
      needsConfirmation: true,
      confirmType: 'noData',
      message: message
    };
  }

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

  try {
    MailApp.sendEmail({
      to: uniqueEmails.join(','),
      subject: 'Timesheet OT Report (' + report.startDate + ' - ' + report.endDate + ')',
      body: buildReportText(report),
      htmlBody: buildReportHtml(report),
      attachments: attachments.length > 0 ? attachments : undefined
    });
  } catch (mailErr) {
    Logger.log('MailApp.sendEmail failed: ' + mailErr);
    throw new Error('Report was generated but the email failed to send: ' + (mailErr.message || mailErr) +
      '. This is often a daily MailApp quota limit - check Apps Script quotas.');
  }

  logAudit('send', null, null, null, 'Sent report ' + report.startDate + '-' + report.endDate + ' to ' + uniqueEmails.length + ' recipients' + (payload.attachPdf ? ' with Excel' : ''));

  let msg = 'Sent to ' + uniqueEmails.length + ' recipient(s) - ' + report.rows.length + ' employee(s) covered.';
  if (attachments.length > 0) msg += ' (Excel attached)';
  if (report.skippedSheets && report.skippedSheets.length > 0) {
    msg += ' Note: ' + report.skippedSheets.length + ' sheet(s) had errors and were skipped - check Apps Script logs.';
  }

  return { success: true, message: msg };
}

function buildReportHtml(report) {
  const rowsHtml = report.rows.map(function (r) {
    const coverageNote = r.partialCoverage
      ? '<div style="font-size:10.5px;color:#b45309;margin-top:2px;">⚑ partial - check for leave</div>'
      : '';
    return '<tr>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #eee;">' + escapeHtml(r.name) + coverageNote + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">' + r.normalHours.toFixed(2) + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;font-weight:600;color:#b56f39;">' + r.otHours.toFixed(2) + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">' + r.totalHours.toFixed(2) + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">' + r.daysWorked + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;font-size:11px;color:#6b7280;">' + r.firstEntry + ' – ' + r.lastEntry + '</td>' +
      '</tr>';
  }).join('');

  const emptyRow = report.rows.length === 0
    ? '<tr><td colspan="6" style="padding:12px;color:#6b7280;">No entries found for this date range.</td></tr>'
    : '';

  const skippedNote = (report.skippedSheets && report.skippedSheets.length > 0)
    ? '<p style="font-size:11px;color:#dc2626;margin-top:4px;">⚠ ' + report.skippedSheets.length +
      ' sheet(s) could not be read and were excluded from this report - check Apps Script logs for details.</p>'
    : '';

  const futureNote = report.futureRangeNote
    ? '<p style="font-size:11.5px;color:#b45309;background:#fffbeb;border:1px solid #f59e0b;padding:6px 10px;border-radius:6px;margin:0 0 10px;">⏳ ' +
      escapeHtml(report.futureRangeNote) + '</p>'
    : '';

  const lagNote = report.dataLagNote
    ? '<p style="font-size:11.5px;color:#0369a1;background:#f0f9ff;border:1px solid #7dd3fc;padding:6px 10px;border-radius:6px;margin:0 0 14px;">ℹ ' +
      escapeHtml(report.dataLagNote) + '</p>'
    : '';

  return '<div style="font-family:Arial,sans-serif;color:#1f2937;">' +
    '<h2 style="color:#b56f39;margin-bottom:4px;">Timesheet OT Report</h2>' +
    '<p style="color:#6b7280;margin-top:0;margin-bottom:16px;">' + report.startDate + ' to ' + report.endDate + '</p>' +
    futureNote +
    lagNote +
    '<table style="border-collapse:collapse;width:100%;max-width:680px;">' +
    '<thead><tr style="background:#fbf0e6;">' +
    '<th style="padding:8px 12px;text-align:left;">Employee</th>' +
    '<th style="padding:8px 12px;text-align:right;">Normal Hrs</th>' +
    '<th style="padding:8px 12px;text-align:right;">OT Hrs</th>' +
    '<th style="padding:8px 12px;text-align:right;">Total Hrs</th>' +
    '<th style="padding:8px 12px;text-align:right;">Days</th>' +
    '<th style="padding:8px 12px;text-align:right;">Entry Range</th>' +
    '</tr></thead><tbody>' + rowsHtml + emptyRow + '</tbody></table>' +
    '<p style="font-size:11px;color:#9ca3af;margin-top:10px;">⚑ = worked noticeably fewer days than the report period covers - worth checking for leave, resignation, or a late join date before finalizing.</p>' +
    skippedNote +
    '</div>';
}

function buildReportText(report) {
  let text = 'TIMESHEET OT REPORT\n' + report.startDate + ' to ' + report.endDate + '\n\n';

  if (report.futureRangeNote) {
    text += '⏳ ' + report.futureRangeNote + '\n\n';
  }
  if (report.dataLagNote) {
    text += 'ℹ ' + report.dataLagNote + '\n\n';
  }

  if (report.rows.length === 0) {
    text += 'No entries found for this date range.\n';
    return text;
  }

  text += 'Employee | Normal Hrs | OT Hrs | Total Hrs | Days | Entry Range\n';
  report.rows.forEach(function (r) {
    const flag = r.partialCoverage ? ' [partial - check leave]' : '';
    text += r.name + flag + ' | ' + r.normalHours.toFixed(2) + ' | ' + r.otHours.toFixed(2) + ' | ' +
      r.totalHours.toFixed(2) + ' | ' + r.daysWorked + ' | ' + r.firstEntry + '-' + r.lastEntry + '\n';
  });

  if (report.skippedSheets && report.skippedSheets.length > 0) {
    text += '\n⚠ ' + report.skippedSheets.length + ' sheet(s) could not be read and were excluded - check Apps Script logs.\n';
  }

  return text;
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}



// ---------------------------------------------------------------------
// DURATION & TIME PARSING HELPERS
// ---------------------------------------------------------------------

/** Calculates total decimal hours between a start time and end time value. Returns 0 (never throws) on bad input. */
function getHoursBetweenTimes(startVal, endVal) {
  if (startVal === null || startVal === undefined || startVal === '') return 0;
  if (endVal === null || endVal === undefined || endVal === '') return 0;

  const startMs = parseTimeToMs(startVal);
  const endMs = parseTimeToMs(endVal);
  if (startMs === null || endMs === null) return 0;

  let diffMs = endMs - startMs;
  if (diffMs <= 0) {
    diffMs += 24 * 60 * 60 * 1000; // Account for shifts crossing midnight
  }

  return diffMs / (1000 * 60 * 60);
}

/** Parses times from multiple types (Date objects, raw numbers, or formatted strings) into MS-since-midnight. Returns null if unparseable. */
function parseTimeToMs(val) {
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    return (val.getHours() * 60 + val.getMinutes()) * 60 * 1000;
  }
  if (typeof val === 'number') {
    if (isNaN(val) || val < 0) return null;
    // Sheets stores time-of-day as a fraction of a day; guard against
    // someone accidentally putting a full serial date number in here.
    const fractionalDay = val % 1;
    return fractionalDay * 24 * 60 * 60 * 1000;
  }
  if (typeof val === 'string' && val.trim() !== '') {
    const s = val.trim().toUpperCase();
    const match = s.match(/^(\d{1,2}):(\d{2})(?:\s*(AM|PM))?$/);
    if (match) {
      let hours = parseInt(match[1], 10);
      const minutes = parseInt(match[2], 10);
      const ampm = match[3];
      if (hours > 23 || minutes > 59) return null;
      if (ampm === 'PM' && hours < 12) hours += 12;
      if (ampm === 'AM' && hours === 12) hours = 0;
      return (hours * 60 + minutes) * 60 * 1000;
    }
  }
  return null;
}