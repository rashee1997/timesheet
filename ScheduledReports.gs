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
  try {
    var existingTriggers = ScriptApp.getProjectTriggers().filter(function (t) {
      return t.getHandlerFunction() === 'runScheduledReport';
    });
    existingTriggers.forEach(function (t) { ScriptApp.deleteTrigger(t); });

    ScriptApp.newTrigger('runScheduledReport')
      .timeBased()
      .onMonthDay(config.dayOfMonth || 25)
      .atHour(config.hour || 8)
      .create();

    config.enabled = true;
    PropertiesService.getScriptProperties().setProperty(SCHEDULE_PROPERTY_KEY, JSON.stringify(config));
    return { success: true, message: 'Scheduled report set for day ' + (config.dayOfMonth || 25) + ' at ' + (config.hour || 8) + ':00.' };
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
  } catch (e) {
    Logger.log('runScheduledReport failed: ' + e);
  }
}
