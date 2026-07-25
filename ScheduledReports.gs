var SCHEDULE_PROPERTY_KEY = 'scheduledReportConfig';

function getScheduledReportConfig() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(SCHEDULE_PROPERTY_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function setScheduledReportConfig(config) {
  if (!config || !config.emails || config.emails.length === 0) {
    return { success: false, error: 'At least one recipient email is required.' };
  }
  var dayOfMonth = parseInt(config.dayOfMonth, 10) || 25;
  var hour = config.hour === 0 ? 0 : (parseInt(config.hour, 10) || 8);
  // Days 29-31 never fire in shorter months, so Apps Script only supports 1-28.
  if (dayOfMonth < 1 || dayOfMonth > 28) {
    return { success: false, error: 'Day of month must be between 1 and 28 (later days never fire in short months).' };
  }
  if (hour < 0 || hour > 23) {
    return { success: false, error: 'Hour must be between 0 and 23.' };
  }
  for (var i = 0; i < config.emails.length; i++) {
    if (!EMAIL_PATTERN.test(String(config.emails[i]).trim())) {
      return { success: false, error: 'Invalid email address: ' + config.emails[i] };
    }
  }
  try {
    withScriptLock_(function () {
      var existingTriggers = ScriptApp.getProjectTriggers().filter(function (t) {
        return t.getHandlerFunction() === 'runScheduledReport';
      });
      existingTriggers.forEach(function (t) { ScriptApp.deleteTrigger(t); });

      ScriptApp.newTrigger('runScheduledReport')
        .timeBased()
        .onMonthDay(dayOfMonth)
        .atHour(hour)
        .create();

      config.enabled = true;
      config.dayOfMonth = dayOfMonth;
      config.hour = hour;
      PropertiesService.getScriptProperties().setProperty(SCHEDULE_PROPERTY_KEY, JSON.stringify(config));
    });
    return { success: true, message: 'Scheduled report set for day ' + dayOfMonth + ' at ' + hour + ':00.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function disableScheduledReport() {
  try {
    var triggers = ScriptApp.getProjectTriggers().filter(function (t) {
      return t.getHandlerFunction() === 'runScheduledReport';
    });
    triggers.forEach(function (t) { ScriptApp.deleteTrigger(t); });
    PropertiesService.getScriptProperties().deleteProperty(SCHEDULE_PROPERTY_KEY);
    return { success: true, message: 'Scheduled report disabled.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function runScheduledReport() {
  try {
    var config = getScheduledReportConfig();
    if (!config) return;

    var range = getDefaultReportRange();
    var selectedEmployees = config.selectedEmployees || [];
    var report = generateOtReport(range.start, range.end, selectedEmployees);

    MailApp.sendEmail({
      to: config.emails.join(','),
      subject: 'Scheduled Timesheet OT Report (' + report.startDate + ' - ' + report.endDate + ')',
      body: buildReportText(report),
      htmlBody: buildReportHtml(report)
    });
    Logger.log('Scheduled report sent to ' + config.emails.length + ' recipient(s).');

    if (config.sendPerEmployee) {
      try {
        sendPerEmployeeReports({
          startDate: range.start,
          endDate: range.end,
          attachExcel: !!config.attachExcelPerEmployee
        });
        Logger.log('Scheduled per-employee reports sent.');
      } catch (perEmpErr) {
        Logger.log('Scheduled per-employee reports failed: ' + perEmpErr);
      }
    }
  } catch (e) {
    Logger.log('runScheduledReport failed: ' + e);
    // Nobody watches the execution log — tell the recipients the report
    // didn't go out, and leave an audit trail.
    try {
      var failConfig = getScheduledReportConfig();
      if (failConfig && failConfig.emails && failConfig.emails.length > 0) {
        MailApp.sendEmail({
          to: failConfig.emails.join(','),
          subject: 'Scheduled Timesheet OT Report FAILED',
          body: 'The scheduled timesheet report could not be generated or sent.\n\nError: ' +
            (e.message || e) + '\n\nPlease check the timesheet system or re-run the report manually.'
        });
      }
    } catch (mailErr) {
      Logger.log('runScheduledReport failure notice also failed: ' + mailErr);
    }
    try {
      logAudit('scheduled_report_failed', '', null, null, String(e.message || e), 'system');
    } catch (auditErr) {
      Logger.log('runScheduledReport failure audit failed: ' + auditErr);
    }
  }
}
