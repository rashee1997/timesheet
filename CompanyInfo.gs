var COMPANY_PROP_NAME = 'companyName';
var COMPANY_LOGO_PROP = 'companyLogoFileId';

function getCompanyInfo() {
  var props = PropertiesService.getScriptProperties();
  return {
    name: props.getProperty(COMPANY_PROP_NAME) || 'United Ocean Trading and Services',
    logoFileId: props.getProperty(COMPANY_LOGO_PROP) || null
  };
}

function saveCompanyInfo(formData) {
  var props = PropertiesService.getScriptProperties();
  var name = formData.companyName && String(formData.companyName).trim();
  if (name) props.setProperty(COMPANY_PROP_NAME, name);

  if (formData.logo && typeof formData.logo === 'object' && formData.logo.bytes) {
    try {
      var decodedBytes = Utilities.base64Decode(formData.logo.bytes);
      var blob = Utilities.newBlob(decodedBytes, formData.logo.mimeType || 'image/png', 'logo');
      var file = DriveApp.createFile(blob);
      var oldId = props.getProperty(COMPANY_LOGO_PROP);
      if (oldId) {
        try { DriveApp.getFileById(oldId).setTrashed(true); } catch (e) {}
      }
      props.setProperty(COMPANY_LOGO_PROP, file.getId());
    } catch (e) {
      return { success: false, error: 'Logo upload failed: ' + e.message };
    }
  }

  return { success: true, message: 'Company information saved.' };
}

function removeCompanyLogo() {
  var props = PropertiesService.getScriptProperties();
  var oldId = props.getProperty(COMPANY_LOGO_PROP);
  if (oldId) {
    try { DriveApp.getFileById(oldId).setTrashed(true); } catch (e) {}
  }
  props.deleteProperty(COMPANY_LOGO_PROP);
  return { success: true };
}

function getCompanyLogoUrl() {
  var id = PropertiesService.getScriptProperties().getProperty(COMPANY_LOGO_PROP);
  if (!id) return null;
  return 'https://drive.google.com/uc?export=view&id=' + id;
}
