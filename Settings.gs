var REST_DAYS_PROP = 'restDays';
var EMPLOYEE_EMAILS_PROP = 'employeeEmails';

var WEEKDAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

function getRestDays() {
  var raw = PropertiesService.getScriptProperties().getProperty(REST_DAYS_PROP);
  if (!raw) return [5];
  try {
    var parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [5];
  } catch (e) {
    return [5];
  }
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
  PropertiesService.getScriptProperties().setProperty(EMPLOYEE_EMAILS_PROP, JSON.stringify(map));
  return { success: true, message: 'Employee emails saved.' };
}

function getEmployeeEmailSettings() {
  var employees = getAllEmployees();
  var emailMap = getEmployeeEmailMap();
  return { employees: employees, emailMap: emailMap };
}

function findEmailForEmployee(name, emailMap) {
  if (emailMap[name]) return emailMap[name];
  var lower = String(name).toLowerCase().trim();
  var keys = Object.keys(emailMap);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].toLowerCase().trim() === lower) return emailMap[keys[i]];
  }
  return null;
}
