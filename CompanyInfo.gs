var COMPANY_PROP_NAME = 'companyName';
// Stores an absolute logo URL directly - either a Vercel Blob URL (new web
// app upload path) or a Drive view URL (legacy CompanySettings.html dialog,
// which still uploads raw bytes to Drive below).
var COMPANY_LOGO_PROP = 'companyLogoUrl';
var DRIVE_LOGO_URL_RE = /^https:\/\/drive\.google\.com\/uc\?export=view&id=(.+)$/;

function getCompanyInfo() {
  var props = PropertiesService.getScriptProperties();
  return {
    name: props.getProperty(COMPANY_PROP_NAME) || 'United Ocean Trading and Services',
    logoUrl: props.getProperty(COMPANY_LOGO_PROP) || null
  };
}

function saveCompanyInfo(formData) {
  var props = PropertiesService.getScriptProperties();
  var name = formData.companyName && String(formData.companyName).trim();
  if (name) props.setProperty(COMPANY_PROP_NAME, name);

  if (typeof formData.logoUrl === 'string' && formData.logoUrl.trim()) {
    // New web app path: already uploaded to Vercel Blob, just store the URL.
    props.setProperty(COMPANY_LOGO_PROP, formData.logoUrl.trim());
  } else if (formData.logo && typeof formData.logo === 'object' && formData.logo.bytes) {
    // Legacy Sheets-dialog path: upload raw bytes to Drive.
    try {
      var decodedBytes = Utilities.base64Decode(formData.logo.bytes);
      var blob = Utilities.newBlob(decodedBytes, formData.logo.mimeType || 'image/png', 'logo');
      var file = DriveApp.createFile(blob);
      var oldUrl = props.getProperty(COMPANY_LOGO_PROP);
      var oldMatch = oldUrl && oldUrl.match(DRIVE_LOGO_URL_RE);
      if (oldMatch) {
        try { DriveApp.getFileById(oldMatch[1]).setTrashed(true); } catch (e) {}
      }
      props.setProperty(COMPANY_LOGO_PROP, 'https://drive.google.com/uc?export=view&id=' + file.getId());
    } catch (e) {
      return { success: false, error: 'Logo upload failed: ' + e.message };
    }
  }

  return { success: true, message: 'Company information saved.' };
}

function removeCompanyLogo() {
  var props = PropertiesService.getScriptProperties();
  var oldUrl = props.getProperty(COMPANY_LOGO_PROP);
  var oldMatch = oldUrl && oldUrl.match(DRIVE_LOGO_URL_RE);
  if (oldMatch) {
    try { DriveApp.getFileById(oldMatch[1]).setTrashed(true); } catch (e) {}
  }
  props.deleteProperty(COMPANY_LOGO_PROP);
  return { success: true };
}

function getCompanyLogoUrl() {
  return PropertiesService.getScriptProperties().getProperty(COMPANY_LOGO_PROP) || null;
}
