// =====================================================================
// CRYSTAL CATERERS & EVENTS — CATERING TOOL
// Code.gs — Menu, sidebars, ensureColumns, clients, menu items,
//           number generation, quote save/update/get
// (Config, formatters, Drive helpers live in Config.gs)
// (PDF generators and admin tools live in their own modules)
// =====================================================================

// ==================== WEB APP ENTRY POINT ====================
function doGet(e) {
  var page = e.parameter ? e.parameter.page : '';
  var bizName = getConfig('BusinessName', 'Catering Tool');

  if (page === 'sign') {
    var dn = e.parameter.dn || '';
    var token = e.parameter.token || '';
    var template = HtmlService.createTemplateFromFile('DeliveryNoteSign');
    template.dn = dn;
    template.token = token;
    var html = template.evaluate();
    html.setTitle(bizName + ' — Sign Delivery Note')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    return html;
  }

  return HtmlService.createHtmlOutputFromFile('Shell')
    .setTitle(bizName)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getPageHTML(page) {
  var htmlFile;
  switch (page) {
    case 'quote':             htmlFile = 'QuoteForm'; break;
    case 'sendQuote':         htmlFile = 'SendQuoteForm'; break;
    case 'deliveryNote':      htmlFile = 'DeliveryNoteForm'; break;
    case 'sendDeliveryNote':  htmlFile = 'SendDeliveryNoteForm'; break;
    case 'viewDN':            htmlFile = 'ViewDeliveryNotes'; break;
    case 'convert':           htmlFile = 'ConvertForm'; break;
    case 'sendInvoice':       htmlFile = 'SendInvoiceForm'; break;
    case 'payment':           htmlFile = 'PaymentForm'; break;
    case 'statement':         htmlFile = 'StatementForm'; break;
    case 'documents':         htmlFile = 'DocumentsForm'; break;
    default:                  htmlFile = 'QuoteForm';
  }
  return HtmlService.createHtmlOutputFromFile(htmlFile).getContent();
}

// ==================== SPREADSHEET MENU ====================
function onOpen() {
  try {
    ensureColumns();
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('🧑‍🍳 Catering Tool')
      .addItem('📄 New Quote', 'showQuoteSidebar')
      .addItem('📧 Send Quote', 'showSendQuoteSidebar')
      .addSeparator()
      .addItem('📦 New Delivery Note', 'showDeliveryNoteSidebar')
      .addItem('📧 Send Delivery Note', 'showSendDeliveryNoteSidebar')
      .addItem('📋 View Delivery Notes', 'showViewDeliveryNotesSidebar')
      .addSeparator()
      .addItem('🔄 Convert Quote to Invoice', 'showConvertSidebar')
      .addItem('📧 Send Invoice', 'showSendInvoiceSidebar')
      .addItem('💳 Record Payment', 'showPaymentSidebar')
      .addItem('📄 Unpaid Statement', 'showStatementSidebar')
      .addSeparator()
      .addItem('📂 View Documents', 'showDocumentsSidebar')
      .addToUi();
  } catch (e) {
    Logger.log('onOpen error: ' + e.message);
  }
}

// ==================== SIDEBAR OPENERS ====================
function showQuoteSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('QuoteForm').setTitle('New Quotation').setWidth(400));
}
function showSendQuoteSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('SendQuoteForm').setTitle('Send Quote').setWidth(400));
}
function showDeliveryNoteSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('DeliveryNoteForm').setTitle('New Delivery Note').setWidth(400));
}
function showSendDeliveryNoteSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('SendDeliveryNoteForm').setTitle('Send Delivery Note').setWidth(400));
}
function showViewDeliveryNotesSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('ViewDeliveryNotes').setTitle('Delivery Notes').setWidth(450));
}
function showConvertSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('ConvertForm').setTitle('Convert to Invoice').setWidth(400));
}
function showSendInvoiceSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('SendInvoiceForm').setTitle('Send Invoice').setWidth(400));
}
function showPaymentSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('PaymentForm').setTitle('Record Payment').setWidth(400));
}
function showStatementSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('StatementForm').setTitle('Unpaid Statement').setWidth(400));
}
function showDocumentsSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('DocumentsForm').setTitle('Documents').setWidth(400));
}

// ==================== ENSURE COLUMNS ====================
function ensureColumns() {
  var ss = SpreadsheetApp.getActive();

  // ------ QUOTES ------
  var qSheet = ss.getSheetByName('Quotes');
  if (qSheet) {
    var headers = qSheet.getRange(1, 1, 1, qSheet.getLastColumn()).getValues()[0];

    if (headers.indexOf('FreshPriceAtQuote') === -1) {
      qSheet.insertColumnAfter(11);
      qSheet.getRange(1, 12).setValue('FreshPriceAtQuote');
      headers = qSheet.getRange(1, 1, 1, qSheet.getLastColumn()).getValues()[0];
    }
    if (headers.indexOf('ServiceDescription') === -1) {
      qSheet.insertColumnAfter(12);
      qSheet.getRange(1, 13).setValue('ServiceDescription');
      headers = qSheet.getRange(1, 1, 1, qSheet.getLastColumn()).getValues()[0];
    }
    if (headers.indexOf('ShowItemDescriptions') === -1) {
      qSheet.insertColumnAfter(13);
      qSheet.getRange(1, 14).setValue('ShowItemDescriptions');
      headers = qSheet.getRange(1, 1, 1, qSheet.getLastColumn()).getValues()[0];
    }
    if (headers.indexOf('IncludeInMenu') === -1) {
      qSheet.insertColumnAfter(14);
      qSheet.getRange(1, 15).setValue('IncludeInMenu');
      headers = qSheet.getRange(1, 1, 1, qSheet.getLastColumn()).getValues()[0];
    }
    if (headers.indexOf('MenuOverride') === -1) {
      qSheet.insertColumnAfter(15);
      qSheet.getRange(1, 16).setValue('MenuOverride');
      headers = qSheet.getRange(1, 1, 1, qSheet.getLastColumn()).getValues()[0];
    }
    if (headers.indexOf('ItemDescriptionOverride') === -1) {
      qSheet.insertColumnAfter(16);
      qSheet.getRange(1, 17).setValue('ItemDescriptionOverride');
      headers = qSheet.getRange(1, 1, 1, qSheet.getLastColumn()).getValues()[0];
    }
    if (headers.indexOf('FocalPersonEmail') === -1) {
      var lastCol = qSheet.getLastColumn();
      qSheet.getRange(1, lastCol + 1).setValue('FocalPersonEmail');
    }
    if (headers.indexOf('Days') === -1) {
      var lastCol = qSheet.getLastColumn();
      qSheet.getRange(1, lastCol + 1).setValue('Days');
      headers = qSheet.getRange(1, 1, 1, qSheet.getLastColumn()).getValues()[0];
    }
    if (headers.indexOf('ServiceFeeDays') === -1) {
      var lastCol = qSheet.getLastColumn();
      qSheet.getRange(1, lastCol + 1).setValue('ServiceFeeDays');
    }
  }

  // ------ INVOICES ------
  var iSheet = ss.getSheetByName('Invoices');
  if (iSheet) {
    var headers = iSheet.getRange(1, 1, 1, iSheet.getLastColumn()).getValues()[0];

    if (headers.indexOf('PaymentDate') === -1) {
      iSheet.insertColumnAfter(12);
      iSheet.getRange(1, 13).setValue('PaymentDate');
      headers = iSheet.getRange(1, 1, 1, iSheet.getLastColumn()).getValues()[0];
    }
    if (headers.indexOf('FreshPriceAtQuote') === -1) {
      var payDateIdx = headers.indexOf('PaymentDate');
      iSheet.insertColumnAfter(payDateIdx + 1);
      iSheet.getRange(1, payDateIdx + 2).setValue('FreshPriceAtQuote');
      headers = iSheet.getRange(1, 1, 1, iSheet.getLastColumn()).getValues()[0];
    }
    if (headers.indexOf('ServiceDescription') === -1) {
      var nextCol = iSheet.getLastColumn() + 1;
      iSheet.getRange(1, nextCol).setValue('ServiceDescription');
    }
    if (headers.indexOf('ShowFocalPerson') === -1) {
      var nextCol = iSheet.getLastColumn() + 1;
      iSheet.getRange(1, nextCol).setValue('ShowFocalPerson');
    }
    if (headers.indexOf('PaymentReference') === -1) {
      var nextCol = iSheet.getLastColumn() + 1;
      iSheet.getRange(1, nextCol).setValue('PaymentReference');
    }
    if (headers.indexOf('Days') === -1) {
      var lastCol = iSheet.getLastColumn();
      iSheet.getRange(1, lastCol + 1).setValue('Days');
      headers = iSheet.getRange(1, 1, 1, iSheet.getLastColumn()).getValues()[0];
    }
    if (headers.indexOf('ServiceFeeDays') === -1) {
      var lastCol = iSheet.getLastColumn();
      iSheet.getRange(1, lastCol + 1).setValue('ServiceFeeDays');
    }
  }

  // ------ DELIVERY NOTES ------
  var dnSheet = ss.getSheetByName('DeliveryNotes');
  if (dnSheet) {
    var dnHeaders = dnSheet.getRange(1, 1, 1, dnSheet.getLastColumn()).getValues()[0];
    if (dnHeaders.indexOf('DeliveryDay') === -1) {
      var lastCol = dnSheet.getLastColumn();
      dnSheet.getRange(1, lastCol + 1).setValue('DeliveryDay');
      dnHeaders = dnSheet.getRange(1, 1, 1, dnSheet.getLastColumn()).getValues()[0];
    }
    if (dnHeaders.indexOf('TotalDays') === -1) {
      var lastCol = dnSheet.getLastColumn();
      dnSheet.getRange(1, lastCol + 1).setValue('TotalDays');
    }
  }
}

// ==================== CLIENT HELPERS ====================
function getClients() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Clients');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1]) result.push({ id: data[i][0], name: data[i][1] });
  }
  return result;
}

function getClientName(clientId) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Clients');
  if (!sheet) return 'Unknown';
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === clientId) return data[i][1] || 'Unknown';
  }
  return 'Unknown';
}

function getClientEmail(clientId) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Clients');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === clientId) return data[i][3] || null;
  }
  return null;
}

function getClientAddress(clientId) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Clients');
  if (!sheet) return '';
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === clientId) return data[i][4] || '';
  }
  return '';
}

function addClient(clientData) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Clients');
  if (!sheet) throw new Error('Clients sheet not found.');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toLowerCase() === clientData.name.toLowerCase()) {
      throw new Error('Client "' + clientData.name + '" already exists.');
    }
  }
  var prefix = 'C';
  var year = new Date().getFullYear();
  var max = 0;
  for (var i = 1; i < data.length; i++) {
    var val = data[i][0];
    if (typeof val === 'string' && val.indexOf(prefix + '-' + year + '-') === 0) {
      var num = parseInt(val.replace(prefix + '-' + year + '-', ''), 10);
      if (num > max) max = num;
    }
  }
  var newId = prefix + '-' + year + '-' + String(max + 1).padStart(3, '0');
  sheet.appendRow([
    newId, clientData.name || '', clientData.contact || '',
    clientData.email || '', clientData.address || '', clientData.phone || ''
  ]);
  return { id: newId, name: clientData.name };
}

// ==================== MENU ITEM HELPERS ====================
function getAllMenuItems() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('MenuItems');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var id = data[i][0] ? data[i][0].toString().trim() : '';
    if (id) {
      result.push({
        id: id, name: data[i][1] || '',
        description: data[i][2] || '',
        price: Number(data[i][3]) || 0
      });
    }
  }
  return result;
}

function getMenuItemDetails(itemId) {
  if (!itemId) return null;
  var sheet = SpreadsheetApp.getActive().getSheetByName('MenuItems');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  var trimmedId = String(itemId).trim();
  for (var i = 1; i < data.length; i++) {
    var sheetId = data[i][0] ? data[i][0].toString().trim() : '';
    if (sheetId === trimmedId) {
      return {
        id: data[i][0], name: data[i][1] || trimmedId,
        description: data[i][2] || '',
        price: Number(data[i][3]) || 0
      };
    }
  }
  return null;
}

function getAllMenuPrices() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('MenuItems');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var priceMap = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) priceMap[data[i][0]] = data[i][3];
  }
  return priceMap;
}

function calculateSnackComboPrice(description) {
  if (!description) return 0;
  var snackNames = description.split(/[,;\n]/)
      .map(function(s) { return s.trim(); })
      .filter(function(s) { return s.length > 0; });
  if (snackNames.length === 0) return 0;
  var menuItems = getAllMenuItems();
  var nameMap = {};
  menuItems.forEach(function(item) {
    if (item.name) nameMap[item.name.toLowerCase()] = item.price;
  });
  var total = 0;
  snackNames.forEach(function(name) {
    var lower = name.toLowerCase();
    if (nameMap[lower] !== undefined) total += Number(nameMap[lower]);
  });
  return total;
}

// ==================== NUMBER GENERATION ====================
var USE_GAP_FILLING = true;

function getNextNumber(prefix, sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);
  var data = sheet.getDataRange().getValues();
  var year = new Date().getFullYear();
  var pattern = prefix + '-' + year + '-';

  if (USE_GAP_FILLING) {
    var numbers = [];
    for (var i = 1; i < data.length; i++) {
      var val = data[i][0];
      if (typeof val === 'string' && val.indexOf(pattern) === 0) {
        var num = parseInt(val.replace(pattern, ''), 10);
        if (!isNaN(num)) numbers.push(num);
      }
    }
    numbers.sort(function(a, b) { return a - b; });
    var missing = 1;
    for (var j = 0; j < numbers.length; j++) {
      if (numbers[j] === missing) missing++;
      else if (numbers[j] > missing) break;
    }
    return pattern + String(missing).padStart(3, '0');
  } else {
    var max = 0;
    for (var i = 1; i < data.length; i++) {
      var val = data[i][0];
      if (typeof val === 'string' && val.indexOf(pattern) === 0) {
        var num = parseInt(val.replace(pattern, ''), 10);
        if (num > max) max = num;
      }
    }
    return pattern + String(max + 1).padStart(3, '0');
  }
}

// ==================== QUOTE SAVE / UPDATE / GET ====================
function saveQuotation(formData) {
  var ss = SpreadsheetApp.getActive();
  var quoteSheet = ss.getSheetByName('Quotes');
  var menuSheet = ss.getSheetByName('MenuItems');
  if (!quoteSheet || !menuSheet) throw new Error('Required sheets not found.');

  var quoteNum = getNextNumber('Q', 'Quotes');
  var now = new Date();
  var defaultTax = 0.175;
  var includeMenu = formData.includeMenu || false;
  var menuOverride = formData.menuOverride || '';
  var focalPersonEmail = formData.focalPersonEmail || '';

  var menuData = menuSheet.getDataRange().getValues();
  var existingIds = {};
  for (var i = 1; i < menuData.length; i++) {
    if (menuData[i][0]) existingIds[menuData[i][0]] = i;
  }

  for (var i = 0; i < formData.items.length; i++) {
    var item = formData.items[i];
    var itemId = String(item.itemId || '').trim();
    if (itemId && !existingIds[itemId]) {
      menuSheet.appendRow([itemId, String(item.itemName || '').trim(), String(item.itemDescription || '').trim(), Number(item.unitPrice) || 0]);
      existingIds[itemId] = menuSheet.getLastRow() - 1;
    }
  }

  for (var i = 0; i < formData.items.length; i++) {
    var item = formData.items[i];
    var taxRate = formData.applyTax ? defaultTax : 0;
    if (item.tax !== undefined) taxRate = item.tax;
    var serviceFee = Number(formData.serviceFee) || 0;
    var userPrice = Number(item.unitPrice) || 0;
    var freshPrice = userPrice;

    if (userPrice === 0 && item.itemId && item.itemId.toUpperCase().indexOf('SN') === 0) {
      var calculated = calculateSnackComboPrice(item.itemDescription);
      if (calculated > 0) userPrice = calculated;
    }

    var rowIdx = existingIds[item.itemId];
    if (rowIdx !== undefined && rowIdx !== null && menuData[rowIdx]) {
      freshPrice = Number(menuData[rowIdx][3]) || userPrice;
    }

    quoteSheet.appendRow([
      quoteNum, now, formData.clientId, item.itemId, Number(item.qty),
      userPrice, Number(item.discount || 0), Number(taxRate), 'Draft',
      serviceFee, formData.focalPerson || '', includeMenu,
      formData.serviceDescription || '', freshPrice,
      formData.showItemDescriptions || false, item.includeInMenu || false,
      menuOverride, item.itemDescription || '', focalPersonEmail,
      Number(item.days || 1), Number(formData.serviceFeeDays || 1)
    ]);
  }

  var pdfUrl = generateQuotePDF(quoteNum, false, includeMenu);
  return { quoteNum: quoteNum, pdfUrl: pdfUrl, serviceFee: formData.serviceFee || 0 };
}

function updateQuotation(quoteNumber, formData) {
  var ss = SpreadsheetApp.getActive();
  var quoteSheet = ss.getSheetByName('Quotes');
  var menuSheet = ss.getSheetByName('MenuItems');
  if (!quoteSheet || !menuSheet) throw new Error('Required sheets not found.');

  var now = new Date();
  var defaultTax = 0.175;
  var includeMenu = formData.includeMenu || false;
  var menuOverride = formData.menuOverride || '';
  var focalPersonEmail = formData.focalPersonEmail || '';

  var data = quoteSheet.getDataRange().getValues();
  var existingStatus = '';
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === quoteNumber) { existingStatus = data[i][8] || ''; break; }
  }
  if (existingStatus === 'Invoiced') throw new Error('Cannot edit a quote that has been invoiced.');
  if (existingStatus === 'Superseded') throw new Error('This quote is Superseded. Create a new quote instead.');

  var menuData = menuSheet.getDataRange().getValues();
  var existingIds = {};
  for (var i = 1; i < menuData.length; i++) {
    if (menuData[i][0]) existingIds[menuData[i][0]] = i;
  }

  for (var i = 0; i < formData.items.length; i++) {
    var item = formData.items[i];
    var itemId = String(item.itemId || '').trim();
    if (itemId && !existingIds[itemId]) {
      menuSheet.appendRow([itemId, String(item.itemName || '').trim(), String(item.itemDescription || '').trim(), Number(item.unitPrice) || 0]);
      existingIds[itemId] = menuSheet.getLastRow() - 1;
    }
  }

  var rowsToDelete = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === quoteNumber) rowsToDelete.push(i + 1);
  }
  rowsToDelete.sort(function(a, b) { return b - a; });
  for (var i = 0; i < rowsToDelete.length; i++) quoteSheet.deleteRow(rowsToDelete[i]);

  for (var i = 0; i < formData.items.length; i++) {
    var item = formData.items[i];
    var taxRate = formData.applyTax ? defaultTax : 0;
    if (item.tax !== undefined) taxRate = item.tax;
    var serviceFee = Number(formData.serviceFee) || 0;
    var userPrice = Number(item.unitPrice) || 0;
    var freshPrice = userPrice;

    if (userPrice === 0 && item.itemId && item.itemId.toUpperCase().indexOf('SN') === 0) {
      var calculated = calculateSnackComboPrice(item.itemDescription);
      if (calculated > 0) userPrice = calculated;
    }

    var rowIdx = existingIds[item.itemId];
    if (rowIdx !== undefined && rowIdx !== null && menuData[rowIdx]) {
      freshPrice = Number(menuData[rowIdx][3]) || userPrice;
    }

    quoteSheet.appendRow([
      quoteNumber, now, formData.clientId, item.itemId, Number(item.qty),
      userPrice, Number(item.discount || 0), Number(taxRate), 'Edited',
      serviceFee, formData.focalPerson || '', includeMenu,
      formData.serviceDescription || '', freshPrice,
      formData.showItemDescriptions || false, item.includeInMenu || false,
      menuOverride, item.itemDescription || '', focalPersonEmail,
      Number(item.days || 1), Number(formData.serviceFeeDays || 1)
    ]);
  }

  var pdfUrl = generateQuotePDF(quoteNumber, false, includeMenu);
  return { quoteNum: quoteNumber, pdfUrl: pdfUrl, serviceFee: formData.serviceFee || 0 };
}

function getQuoteData(quoteNumber) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Quotes');
  if (!sheet) throw new Error('Quotes sheet not found.');
  var data = sheet.getDataRange().getValues();

  var rows = [];
  var clientId = '', focalPerson = '', focalPersonEmail = '', serviceFee = 0;
  var serviceFeeDays = 1;
  var includeMenu = false, status = '', serviceDescription = '';
  var showItemDescriptions = false, menuOverride = '';

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === quoteNumber) {
      if (rows.length === 0) {
        clientId = data[i][2];
        focalPerson = data[i][10] || '';
        serviceFee = data[i][9] || 0;
        serviceFeeDays = data[i][20] || 1;
        includeMenu = data[i][11] || false;
        status = data[i][8] || '';
        serviceDescription = data[i][12] || '';
        showItemDescriptions = data[i][14] || false;
        menuOverride = data[i][16] || '';
        focalPersonEmail = data[i][18] || '';
      }
      rows.push({
        itemId: data[i][3], qty: data[i][4], days: data[i][19] || 1,
        unitPrice: data[i][5], discount: data[i][6], tax: data[i][7],
        includeInMenu: data[i][15] || false, itemDescriptionOverride: data[i][17] || ''
      });
    }
  }

  if (rows.length === 0) throw new Error('Quote number not found.');

  var menuItems = getAllMenuItems();
  var menuMap = {};
  menuItems.forEach(function(item) {
    menuMap[item.id] = { name: item.name, description: item.description };
  });

  var items = rows.map(function(row) {
    var menu = menuMap[row.itemId] || { name: '', description: '' };
    return {
      itemId: row.itemId, itemName: menu.name, itemDescription: menu.description,
      qty: row.qty, days: row.days || 1, unitPrice: row.unitPrice,
      discount: row.discount, tax: row.tax, includeInMenu: row.includeInMenu,
      itemDescriptionOverride: row.itemDescriptionOverride
    };
  });

  return {
    quoteNumber: quoteNumber, clientId: clientId, focalPerson: focalPerson,
    focalPersonEmail: focalPersonEmail, serviceFee: serviceFee,
    serviceFeeDays: serviceFeeDays, includeMenu: includeMenu, status: status,
    serviceDescription: serviceDescription, showItemDescriptions: showItemDescriptions,
    menuOverride: menuOverride, items: items
  };
}

function getQuoteList(mode) {
  mode = mode || 'active';
  var sheet = SpreadsheetApp.getActive().getSheetByName('Quotes');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var quotes = [], seen = {};
  for (var i = 1; i < data.length; i++) {
    var num = data[i][0];
    if (!num || seen[num]) continue;
    var status = String(data[i][8] || '').trim();
    if (mode === 'active' && status === 'Superseded') continue;
    if (mode === 'superseded' && status !== 'Superseded') continue;
    seen[num] = true;
    quotes.push(num);
  }
  return quotes;
}

function getQuoteDetails(quoteNumber) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Quotes');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === quoteNumber) {
      return {
        quoteNumber: data[i][0], date: data[i][1], clientId: data[i][2],
        status: data[i][8], serviceFee: data[i][9] || 0,
        focalPerson: data[i][10] || '', includeMenu: data[i][11] || false,
        freshPriceAtQuote: data[i][12] || null
      };
    }
  }
  return null;
}

// ==================== QUOTE PDF LOCATION HELPERS ====================
function moveQuotePDFToInvoicedFolder(quoteNumber) {
  try {
    var sourceFolders = DriveApp.getFoldersByName('Catering PDFs');
    if (!sourceFolders.hasNext()) return;
    var sourceFolder = sourceFolders.next();
    var destFolders = DriveApp.getFoldersByName('Invoiced Quotes');
    var destFolder = destFolders.hasNext() ? destFolders.next() : DriveApp.createFolder('Invoiced Quotes');
    var files = sourceFolder.getFilesByName(quoteNumber + '.pdf');
    if (!files.hasNext()) return;
    files.next().moveTo(destFolder);
  } catch (e) {
    Logger.log('Error moving PDF for ' + quoteNumber + ': ' + e.toString());
  }
}

function findQuotePDF(quoteNumber) {
  var files = DriveApp.getFilesByName(quoteNumber + '.pdf');
  if (files.hasNext()) return files.next();
  var folders = ['Catering PDFs', 'Invoiced Quotes'];
  for (var f = 0; f < folders.length; f++) {
    var folderIter = DriveApp.getFoldersByName(folders[f]);
    if (folderIter.hasNext()) {
      var folder = folderIter.next();
      var fileIter = folder.getFilesByName(quoteNumber + '.pdf');
      if (fileIter.hasNext()) return fileIter.next();
    }
  }
  return null;
}

function getQuotePDFUrl(quoteNumber) {
  var file = findQuotePDF(quoteNumber);
  if (file) return file.getUrl();
  var details = getQuoteDetails(quoteNumber);
  var includeMenu = details ? details.includeMenu : false;
  generateQuotePDF(quoteNumber, false, includeMenu);
  file = findQuotePDF(quoteNumber);
  if (file) return file.getUrl();
  throw new Error('Failed to generate or locate PDF for ' + quoteNumber);
}

