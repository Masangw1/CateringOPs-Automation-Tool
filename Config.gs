// =====================================================================
// CRYSTAL CATERERS & EVENTS — CONFIG & UTILITIES
// Config.gs — Business config loader, formatters, Drive helpers
// =====================================================================

// ==================== CONFIG SYSTEM ====================

// Cache so we only read the Config sheet once per script execution
var _configCache = null;

function _loadConfig() {
  if (_configCache) return _configCache;
  var sheet = SpreadsheetApp.getActive().getSheetByName('Config');
  var cache = {};
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var key = String(data[i][0] || '').trim();
      var val = data[i][1];
      if (key) cache[key] = val;
    }
  }
  _configCache = cache;
  return cache;
}

/**
 * Get a config value by key.
 * Returns defaultValue if key is missing or empty.
 */
function getConfig(key, defaultValue) {
  var cache = _loadConfig();
  if (Object.prototype.hasOwnProperty.call(cache, key)) {
    var v = cache[key];
    if (v !== '' && v !== null && v !== undefined) return v;
  }
  return (defaultValue !== undefined) ? defaultValue : '';
}

/**
 * Force-refresh the config cache.
 * Call this if you edit Config and want the change to apply immediately.
 */
function refreshConfig() {
  _configCache = null;
  return _loadConfig();
}

// ==================== FORMATTING HELPERS ====================

function formatMoney(value) {
  var num = Number(value) || 0;
  var parts = num.toFixed(2).split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return parts.join('.');
}

function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ==================== DRIVE HELPERS ====================

function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function getFileIdFromUrl(url) {
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

// ==================== LOGO ====================

function getLogoBase64() {
  var fileId = getConfig('LogoFileId', '1jYWFoWhmTa9nAcVRXMGto76D8W-Ha2_d');
  if (!fileId) {
    Logger.log('⚠️ No LogoFileId set in Config — logo will be skipped.');
    return '';
  }
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var base64 = Utilities.base64Encode(blob.getBytes());
    var mimeType = blob.getContentType();
    return 'data:' + mimeType + ';base64,' + base64;
  } catch (e) {
    Logger.log('❌ Logo error: ' + e.toString());
    return '';
  }
}

// ==================== EDITOR TEST WRAPPERS ====================

function RUN_testConfig() {
  Logger.log('BusinessName: ' + getConfig('BusinessName'));
  Logger.log('Tagline:      ' + getConfig('Tagline'));
  Logger.log('Address:      ' + getConfig('Address'));
  Logger.log('Phone1:       ' + getConfig('Phone1'));
  Logger.log('Email:        ' + getConfig('Email'));
  Logger.log('CompanyReg:   ' + getConfig('CompanyReg'));
  Logger.log('WHTCert:      ' + getConfig('WHTCert'));
  Logger.log('LogoFileId:   ' + getConfig('LogoFileId'));
  Logger.log('WebAppUrl:    ' + getConfig('WebAppUrl'));
  Logger.log('BankName:     ' + getConfig('BankName'));
  Logger.log('DefaultTax:   ' + getConfig('DefaultTaxRate'));
  Logger.log('MenuName:     ' + getConfig('MenuName'));
  Logger.log('---');
  Logger.log('Unknown key:  ' + getConfig('DoesNotExist', '(fallback)'));
}

function RUN_testEscape() {
  Logger.log(escapeHtml('Crystal Caterers & Events'));
  Logger.log(escapeHtml('A < B & C > D "quoted"'));
}
