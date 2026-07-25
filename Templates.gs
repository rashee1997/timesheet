var TEMPLATES_PROPERTY_KEY = 'shiftTemplates';

function getShiftTemplates() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(TEMPLATES_PROPERTY_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    Logger.log('getShiftTemplates failed: ' + e);
    return [];
  }
}

function saveShiftTemplate(name, startTime, endTime, jobOrder) {
  if (!name || !startTime || !endTime) {
    return { success: false, error: 'Name, start time, and end time are required.' };
  }
  try {
    withScriptLock_(function () {
      var templates = getShiftTemplates();
      templates.push({
        id: 'tpl_' + Date.now(),
        name: String(name).trim(),
        startTime: String(startTime).trim(),
        endTime: String(endTime).trim(),
        jobOrder: String(jobOrder || '').trim()
      });
      PropertiesService.getScriptProperties().setProperty(TEMPLATES_PROPERTY_KEY, JSON.stringify(templates));
    });
    return { success: true, message: 'Template "' + name + '" saved.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteShiftTemplate(id) {
  try {
    withScriptLock_(function () {
      var templates = getShiftTemplates().filter(function (t) { return t.id !== id; });
      PropertiesService.getScriptProperties().setProperty(TEMPLATES_PROPERTY_KEY, JSON.stringify(templates));
    });
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
