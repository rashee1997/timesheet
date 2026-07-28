var REST_DAYS_PROP = 'restDays';
var EMPLOYEE_EMAILS_PROP = 'employeeEmails';

var WEEKDAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

// Per-execution memo: buildHourFormulas and Guard.onEdit call getRestDays for
// every entry/cell, and each call was a fresh PropertiesService read.
var _restDaysMemo = null;

function getRestDays() {
  if (_restDaysMemo) return _restDaysMemo;
  var raw = PropertiesService.getScriptProperties().getProperty(REST_DAYS_PROP);
  var result = [5];
  if (raw) {
    try {
      var parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) result = parsed;
    } catch (e) {
      // fall back to default
    }
  }
  _restDaysMemo = result;
  return result;
}

function saveRestDays(days) {
  if (!Array.isArray(days) || days.length === 0) {
    return { success: false, error: 'Select at least one rest day.' };
  }
  for (var i = 0; i < days.length; i++) {
    if (typeof days[i] !== 'number' || days[i] < 1 || days[i] > 7) {
      return { success: false, error: 'Invalid day number: ' + days[i] };
    }
  }
  PropertiesService.getScriptProperties().setProperty(REST_DAYS_PROP, JSON.stringify(days));
  _restDaysMemo = null;
  invalidateTimesheetScanCache_();
  return { success: true, message: 'Rest days updated.' };
}

function getRestDayNames() {
  return getRestDays().map(function (d) { return WEEKDAY_NAMES[d - 1]; });
}

function weekdayToJsDay(weekdayNum) {
  return weekdayNum === 7 ? 0 : weekdayNum;
}

function getRestDaysJs() {
  return getRestDays().map(weekdayToJsDay);
}

function getEmployeeEmailMap() {
  var raw = PropertiesService.getScriptProperties().getProperty(EMPLOYEE_EMAILS_PROP);
  if (!raw) return {};
  try {
    return JSON.parse(raw);
  } catch (e) {
    return {};
  }
}

function saveEmployeeEmails(map) {
  if (typeof map !== 'object' || map === null) {
    return { success: false, error: 'Invalid email map.' };
  }
  try {
    withScriptLock_(function () {
      PropertiesService.getScriptProperties().setProperty(EMPLOYEE_EMAILS_PROP, JSON.stringify(map));
    });
    return { success: true, message: 'Employee emails saved.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getEmployeeEmailSettings() {
  var employees = getAllEmployees();
  var emailMap = getEmployeeEmailMap();
  return { employees: employees, emailMap: emailMap };
}
