// =====================================================================
// CRYSTAL CATERERS & EVENTS — INVOICES
// Invoices.gs — Conversion, PDF, totals, payment, send, lookups
// =====================================================================

// ==================== CONVERT SINGLE QUOTE → INVOICE ====================
function convertToInvoice(quoteNumber, useQuoteDate, showFocalPerson) {
  useQuoteDate = useQuoteDate || false;
  var showFocalPerson = showFocalPerson || false;
  var ss = SpreadsheetApp.getActive();
  var quoteSheet = ss.getSheetByName('Quotes');
  var invoiceSheet = ss.getSheetByName('Invoices');
  if (!quoteSheet || !invoiceSheet) throw new Error('Required sheets not found.');

  var invoiceNum = getNextNumber('INV', 'Invoices');
  var data = quoteSheet.getDataRange().getValues();
  var found = false;
  var serviceFee = 0, focalPerson = '', quoteDate = null;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === quoteNumber) {
      serviceFee = data[i][9] || 0;
      focalPerson = data[i][10] || '';
      quoteDate = data[i][1];
      break;
    }
  }

  var invoiceDate = (useQuoteDate && quoteDate) ? new Date(quoteDate) : new Date();
  var dueDate = new Date(invoiceDate);
  dueDate.setDate(dueDate.getDate() + 14);

  ensureColumns();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === quoteNumber) {
      found = true;
      invoiceSheet.appendRow([
        invoiceNum, quoteNumber, invoiceDate, dueDate,
        data[i][2], data[i][3], data[i][4], data[i][5],
        data[i][6], data[i][7], 'Unpaid', serviceFee, focalPerson, null,
        data[i][13] || null, data[i][12] || '', showFocalPerson,
        data[i][19] || 1, data[i][20] || 1
      ]);
    }
  }
  if (!found) throw new Error('Quote number not found.');

  for (var j = 1; j < data.length; j++) {
    if (data[j][0] === quoteNumber) {
      quoteSheet.getRange(j + 1, 9).setValue('Invoiced');
      break;
    }
  }

  var pdfUrl = generateInvoicePDF(invoiceNum, false);
  moveQuotePDFToInvoicedFolder(quoteNumber);

  var dateUsed = useQuoteDate ? 'Quote Date' : 'Current Date';
  var dateDisplay = useQuoteDate ? new Date(quoteDate).toDateString() : new Date().toDateString();

  return {
    invoiceNum: invoiceNum, pdfUrl: pdfUrl, serviceFee: serviceFee,
    dateUsed: dateUsed, invoiceDate: dateDisplay,
    message: 'Invoice ' + invoiceNum + ' created using ' + dateUsed + ': ' + dateDisplay
  };
}

// ==================== CONVERT MULTIPLE QUOTES → INVOICE ====================
function convertMultipleToInvoice(quoteNumbers, useQuoteDate, showFocalPerson) {
  if (!quoteNumbers || quoteNumbers.length === 0) throw new Error('No quotes selected.');
  var showFocalPerson = showFocalPerson || false;

  var ss = SpreadsheetApp.getActive();
  var quoteSheet = ss.getSheetByName('Quotes');
  var invoiceSheet = ss.getSheetByName('Invoices');
  if (!quoteSheet || !invoiceSheet) throw new Error('Required sheets not found.');

  var allQuoteData = [];
  var clientId = null;
  var firstQuoteDate = null;

  for (var qi = 0; qi < quoteNumbers.length; qi++) {
    var qNum = quoteNumbers[qi];
    var data = quoteSheet.getDataRange().getValues();
    var found = false;
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === qNum) {
        if (!found) {
          if (clientId === null) {
            clientId = data[i][2];
            firstQuoteDate = data[i][1];
          } else if (data[i][2] !== clientId) {
            throw new Error('All quotes must belong to the same client.');
          }
        }
        rows.push(data[i]);
        found = true;
      }
    }
    if (!found) throw new Error('Quote ' + qNum + ' not found.');
    allQuoteData.push({ quoteNum: qNum, rows: rows });
  }

  var invoiceDate = (useQuoteDate && firstQuoteDate) ? new Date(firstQuoteDate) : new Date();
  var dueDate = new Date(invoiceDate);
  dueDate.setDate(dueDate.getDate() + 14);

  var totalServiceFee = 0;
  var focalPerson = '';
  for (var qidx = 0; qidx < allQuoteData.length; qidx++) {
    var quoteRows = allQuoteData[qidx].rows;
    totalServiceFee += Number(quoteRows[0][9]) || 0;
    if (qidx === 0) focalPerson = quoteRows[0][10] || '';
  }

  var invoiceNum = getNextNumber('INV', 'Invoices');
  ensureColumns();
  var sourceQuotes = quoteNumbers.join(', ');

  for (var qi = 0; qi < allQuoteData.length; qi++) {
    var rows = allQuoteData[qi].rows;
    for (var ri = 0; ri < rows.length; ri++) {
      var row = rows[ri];
      invoiceSheet.appendRow([
        invoiceNum, sourceQuotes, invoiceDate, dueDate, clientId,
        row[3], row[4], row[5], row[6], row[7], 'Unpaid',
        totalServiceFee, focalPerson, null, row[13] || null, row[12] || '',
        showFocalPerson, row[19] || 1, row[20] || 1
      ]);
    }
  }

  var data2 = quoteSheet.getDataRange().getValues();
  for (var i = 1; i < data2.length; i++) {
    for (var qn = 0; qn < quoteNumbers.length; qn++) {
      if (data2[i][0] === quoteNumbers[qn]) {
        quoteSheet.getRange(i + 1, 9).setValue('Invoiced');
        break;
      }
    }
  }

  var pdfUrl = generateInvoicePDF(invoiceNum, false);
  for (var qn2 = 0; qn2 < quoteNumbers.length; qn2++) moveQuotePDFToInvoicedFolder(quoteNumbers[qn2]);

  var dateUsed = useQuoteDate ? 'Quote Date' : 'Current Date';
  var dateDisplay = useQuoteDate ? new Date(firstQuoteDate).toDateString() : new Date().toDateString();

  return {
    invoiceNum: invoiceNum, pdfUrl: pdfUrl, serviceFee: totalServiceFee,
    dateUsed: dateUsed, invoiceDate: dateDisplay,
    message: 'Invoice ' + invoiceNum + ' created from ' + quoteNumbers.length + ' quote(s).'
  };
}

// ==================== GENERATE INVOICE PDF ====================
function generateInvoicePDF(invoiceNum, showPaymentInfo) {
  showPaymentInfo = showPaymentInfo || false;
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Invoices');
  if (!sheet) throw new Error('Invoices sheet not found.');

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var freshPriceCol = headers.indexOf('FreshPriceAtQuote');
  if (freshPriceCol === -1) freshPriceCol = null;

  var items = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === invoiceNum) items.push(data[i]);
  }
  if (items.length === 0) throw new Error('No items found for this invoice.');

  var status = items[0][10] || 'Unpaid';
  var isPaid = (status === 'Paid');

  var priceMap = getAllMenuPrices();
  var clientId = items[0][4];
  var clientName = getClientName(clientId);
  var quoteNum = items[0][1] || '';
  var invoiceDate = items[0][2] || new Date();
  var dueDate = items[0][3] || new Date(new Date().setDate(new Date().getDate() + 14));
  var serviceFee = items[0][11] || 0;
  var colSFDays = headers.indexOf('ServiceFeeDays');
  var colDays   = headers.indexOf('Days');
  var colSvcDesc = headers.indexOf('ServiceDescription');
  var colFocal  = headers.indexOf('FocalPerson');
  var serviceFeeDays = (colSFDays !== -1 && items[0][colSFDays] !== '') ? Number(items[0][colSFDays]) : 1;
  if (!serviceFeeDays || isNaN(serviceFeeDays)) {
    serviceFeeDays = (colDays !== -1 && items[0][colDays] !== '') ? Number(items[0][colDays]) : 1;
  }
  if (!serviceFeeDays || isNaN(serviceFeeDays)) serviceFeeDays = 1;
  var serviceDescription = (colSvcDesc !== -1) ? (items[0][colSvcDesc] || '') : '';
  var focalPerson = (colFocal !== -1) ? (items[0][colFocal] || '') : '';

  var showFocalCol = headers.indexOf('ShowFocalPerson');
  var showFocalPerson = false;
  if (showFocalCol !== -1 && items[0][showFocalCol] === true) showFocalPerson = true;

  var subtotalExclTax = 0, totalTax = 0;
  var lineItems = [];

  for (var i = 0; i < items.length; i++) {
    var row = items[i];
    var itemId = row[5];
    var qty = row[6];
    var days = (colDays !== -1 && row[colDays] !== '') ? Number(row[colDays]) : 1;
    if (!days || isNaN(days)) days = 1;
    var storedPrice = row[7];
    var disc = row[8] || 0;
    var tax = row[9] || 0;

    var freshPriceAtQuote = null;
    if (freshPriceCol !== null) freshPriceAtQuote = row[freshPriceCol];

    var actualPrice = storedPrice;
    var isManualOverride = false;
    var freshPriceAtQuoteNum = null;

    if (freshPriceAtQuote !== null && freshPriceAtQuote !== undefined && freshPriceAtQuote !== '') {
      freshPriceAtQuoteNum = Number(freshPriceAtQuote);
      if (Math.abs(actualPrice - freshPriceAtQuoteNum) > 0.001) isManualOverride = true;
    } else {
      var currentFresh = priceMap[itemId];
      if (currentFresh !== undefined && currentFresh !== null && Math.abs(actualPrice - currentFresh) > 0.001) {
        isManualOverride = true;
        freshPriceAtQuoteNum = currentFresh;
      }
    }

    var itemSubtotalExclTax = (qty * days) * actualPrice * (1 - disc);
    subtotalExclTax += itemSubtotalExclTax;
    totalTax += itemSubtotalExclTax * tax;

    lineItems.push({
      itemId: itemId, qty: qty, days: days, actualPrice: actualPrice,
      disc: disc, tax: tax,
      itemSubtotal: (qty * days) * actualPrice * (1 - disc) * (1 + tax),
      itemSubtotalExclTax: itemSubtotalExclTax
    });
  }

  var foodLines = [], equipmentTotalExclTax = 0;
  for (var i = 0; i < lineItems.length; i++) {
    var idUpper = String(lineItems[i].itemId || '').toUpperCase();
    if (idUpper.indexOf('EQ') === 0) {
      equipmentTotalExclTax += lineItems[i].itemSubtotalExclTax;
    } else {
      foodLines.push(lineItems[i]);
    }
  }

  var foodSubtotalExclTax = subtotalExclTax - equipmentTotalExclTax;
  var totalServiceFee = serviceFee * serviceFeeDays;
  var grandTotal = subtotalExclTax + totalTax + totalServiceFee;
  var hasTax = totalTax > 0;
  var logoBase64 = getLogoBase64();

  var html = '<html><head><meta charset="UTF-8"><style>';
  html += 'body { font-family: "Georgia", serif; max-width: 750px; margin: 0 auto; padding: 30px 40px; color: #1e1e1e; }';
  html += '.letterhead { border-bottom: 4px solid #d4a373; padding-bottom: 15px; margin-bottom: 25px; display: flex; align-items: center; justify-content: space-between; }';
  html += '.brand-left { display: flex; align-items: center; gap: 20px; }';
  html += '.logo { max-height: 76px; width: auto; }';
  html += '.business-name { font-size: 27px; font-weight: bold; color: #2c3e50; letter-spacing: 1px; }';
  html += '.tagline { font-size: 13px; color: #7f8c8d; font-style: italic; letter-spacing: 0.5px; margin-top: 2px; }';
  html += '.business-details { font-size: 12px; color: #5d6d7e; line-height: 1.5; text-align: right; }';
  html += '.business-details span { display: block; }';
  html += '.doc-title { font-size: 31px; font-weight: bold; color: #e74c3c; text-align: center; letter-spacing: 4px; margin: 10px 0 5px 0; text-transform: uppercase; }';
  html += '.doc-sub { text-align: center; color: #7f8c8d; font-size: 13px; margin-bottom: 25px; }';
  html += '.invoice-info { display: flex; justify-content: space-between; background: #f8f5f0; padding: 15px 20px; border-radius: 6px; margin-bottom: 20px; }';
  html += '.invoice-info div { font-size: 13px; }';
  html += '.invoice-info strong { color: #2c3e50; }';
  html += '.client-box { background: #f8f5f0; padding: 15px 20px; border-radius: 6px; margin-bottom: 20px; border-left: 6px solid #e74c3c; }';
  html += '.client-box strong { color: #2c3e50; }';
  html += 'table { width: 100%; border-collapse: collapse; margin: 25px 0; font-size: 14px; }';
  html += 'th { background: #2c3e50; color: white; padding: 12px 10px; text-align: left; font-weight: 600; }';
  html += 'td { padding: 12px 10px; border-bottom: 1px solid #e5e8eb; }';
  html += 'tr:last-child td { border-bottom: none; }';
  html += '.totals { text-align: right; margin-top: 15px; font-size: 17px; }';
  html += '.totals .grand { font-size: 25px; font-weight: bold; color: #e74c3c; }';
  html += '.payment-info { margin-top: 20px; padding: 15px; background: #e8f5e9; border-radius: 6px; font-size: 13px; }';
  if (isPaid) {
    html += '.paid-stamp { position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%) rotate(-30deg); font-size: 120px; font-weight: bold; color: rgba(46, 204, 113, 0.25); pointer-events: none; z-index: 1000; text-transform: uppercase; border: 8px solid rgba(46, 204, 113, 0.25); padding: 20px 40px; border-radius: 20px; }';
  }
  html += '@page { margin: 40px 40px 45px 40px; }';
  html += '.pdf-footer { position: fixed; bottom: 0; left: 0; right: 0; text-align: center; font-size: 10px; color: #95a5a6; border-top: 1px solid #e5e8eb; padding: 3px 40px; background: white; line-height: 1.25; }';
  html += '.pdf-footer .line1 { font-size: 10px; color: #7f8c8d; }';
  html += '.pdf-footer .line2 { font-size: 9px; color: #95a5a6; margin-top: 2px; }';
  html += '</style></head><body>';

  if (isPaid) html += '<div class="paid-stamp">PAID</div>';

  html += '<div class="letterhead">';
  html += '  <div class="brand-left">';
  html += '    <img class="logo" src="' + logoBase64 + '" alt="Company Logo" />';
  html += '    <div>';
  html += '      <div class="business-name">' + escapeHtml(getConfig('BusinessName', 'CRYSTAL CATERERS AND EVENTS')).toUpperCase() + '</div>';
  html += '      <div class="tagline">' + escapeHtml(getConfig('Tagline', 'The home of your best Quality Foods & Services')) + '</div>';
  html += '    </div>';
  html += '  </div>';
  html += '  <div class="business-details">';
  html += '    <span>📍 ' + escapeHtml(getConfig('Address')) + '</span>';
  html += '    <span>📱 ' + escapeHtml(getConfig('Phone1')) + '</span>';
  html += '    <span>📱 ' + escapeHtml(getConfig('Phone2')) + '</span>';
  html += '    <span>✉️ ' + escapeHtml(getConfig('Email')) + '</span>';
  html += '  </div>';
  html += '</div>';

  html += '<div class="doc-title">INVOICE</div>';
  html += '<div class="doc-sub">' + invoiceNum + '  |  ' + (new Date(invoiceDate).toDateString()) + '</div>';

  html += '<div class="invoice-info">';
  html += '  <div><strong>Invoice #:</strong> ' + invoiceNum + '</div>';
  var quoteDisplay = quoteNum || '';
  if (quoteDisplay.indexOf(',') > -1) html += '  <div><strong>Quote(s) #:</strong> ' + quoteDisplay + '</div>';
  else html += '  <div><strong>Quote #:</strong> ' + quoteDisplay + '</div>';
  html += '  <div><strong>Due Date:</strong> ' + (new Date(dueDate).toDateString()) + '</div>';
  html += '</div>';

  html += '<div class="client-box">';
  html += '  <strong>Bill To:</strong> ' + clientName;
  if (showFocalPerson && focalPerson) html += '<br><strong>Focal Person:</strong> ' + focalPerson;
  if (serviceDescription) html += '<br><br><strong>Service Description:</strong><br>' + serviceDescription;
  html += '</div>';

  html += '<table>';
  html += '  <thead><tr><th>Item</th><th style="text-align:center;">Qty</th><th style="text-align:center;">Days</th><th style="text-align:right;">Unit Price</th><th style="text-align:right;">Disc.</th><th style="text-align:right;">Tax</th><th style="text-align:right;">Amount</th></tr></thead><tbody>';

  for (var idx = 0; idx < foodLines.length; idx++) {
    var li = foodLines[idx];
    var itemDetails = getMenuItemDetails(li.itemId);
    var itemName = li.itemId;
    if (itemDetails) itemName = itemDetails.name || li.itemId;
    var discPercent = li.disc * 100;
    var taxPercent = li.tax * 100;
    html += '    <tr>';
    html += '      <td><strong>' + li.itemId + '</strong>';
    html += '        <br><span style="font-size:12px;color:#2c3e50;font-weight:500;">' + itemName + '</span>';
    html += '      </td>';
    html += '      <td style="text-align:center;">' + li.qty + '</td>';
    html += '      <td style="text-align:center;">' + li.days + '</td>';
    html += '      <td style="text-align:right;">MK' + Number(li.actualPrice).toFixed(2) + '</td>';
    html += '      <td style="text-align:right;">' + discPercent + '%</td>';
    html += '      <td style="text-align:right;">' + taxPercent + '%</td>';
    html += '      <td style="text-align:right;">MK' + li.itemSubtotal.toFixed(2) + '</td>';
    html += '    </tr>';
  }
  html += '  </tbody></table>';

  html += '<div class="totals">';
  html += '  <div>Subtotal: MK' + formatMoney(foodSubtotalExclTax) + '</div>';
  if (equipmentTotalExclTax > 0) html += '  <div>AV Equipment Hire: MK' + formatMoney(equipmentTotalExclTax) + '</div>';
  if (hasTax) html += '  <div>VAT: MK' + formatMoney(totalTax) + '</div>';
  if (totalServiceFee > 0) {
    html += '  <div>Service Fee: MK' + formatMoney(totalServiceFee) + (serviceFeeDays > 1 ? ' (MK' + formatMoney(serviceFee) + ' x ' + serviceFeeDays + ' days)' : '') + '</div>';
  }
  html += '  <div style="border-top: 2px solid #e5e8eb; padding-top: 8px; margin-top: 8px;">';
  html += '    <span class="grand">TOTAL DUE: MK' + formatMoney(grandTotal) + '</span>';
  html += '  </div>';
  html += '  <div style="font-size:11px;color:#27ae60;margin-top:4px;">* Prices reflect current menu rates unless manually adjusted.</div>';
  html += '</div>';

  if (showPaymentInfo) {
    html += '<div class="payment-info">';
    html += '  <strong>Payment Instructions:</strong><br>';
    html += '  Bank: ' + escapeHtml(getConfig('BankName', '[Your Bank Name]')) + '<br>';
    html += '  Account Name: ' + escapeHtml(getConfig('BusinessName', 'Crystal Caterers and Events')) + '<br>';
    html += '  Account Number: ' + escapeHtml(getConfig('BankAccount', '[Your Account Number]')) + '<br>';
    html += '  Reference: ' + invoiceNum;
    html += '</div>';
  }

  html += '<div class="pdf-footer">';
  html += '  <div class="line1">Thank you for your business. Payment is due within 14 days.</div>';
  html += '  <div class="line2">Company Reg: ' + escapeHtml(getConfig('CompanyReg')) + ' &nbsp;|&nbsp; Withholding Tax Certificate No: ' + escapeHtml(getConfig('WHTCert')) + '</div>';
  html += '</div>';
  html += '</body></html>';

  var blob = Utilities.newBlob(html, 'text/html', invoiceNum + '.html');
  var pdf = blob.getAs('application/pdf');

  var folderIterator = DriveApp.getFoldersByName('Invoices PDFs');
  var folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder('Invoices PDFs');

  var existingFiles = folder.getFilesByName(invoiceNum + '.pdf');
  while (existingFiles.hasNext()) existingFiles.next().setTrashed(true);

  var file = folder.createFile(pdf).setName(invoiceNum + '.pdf');
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ==================== CALCULATE INVOICE TOTAL ====================
function calculateInvoiceTotal(invoiceNumber) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Invoices');
  if (!sheet) return 0;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var colInvoice = headers.indexOf('Invoice#');
  var colQty = headers.indexOf('Qty');
  var colPrice = headers.indexOf('UnitPrice');
  var colDisc = headers.indexOf('Discount%');
  var colTax = headers.indexOf('Tax%');
  var colServiceFee = headers.indexOf('ServiceFee');
  var colDays = headers.indexOf('Days');
  var colServiceFeeDays = headers.indexOf('ServiceFeeDays');
  var colFreshPrice = headers.indexOf('FreshPriceAtQuote');
  var colItemID = headers.indexOf('ItemID');

  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colInvoice]).trim() === invoiceNumber) items.push(data[i]);
  }
  if (items.length === 0) throw new Error('No items found for this invoice.');

  var priceMap = getAllMenuPrices();

  var serviceFee = Number(items[0][colServiceFee]) || 0;
  var sfDaysRaw = (colServiceFeeDays !== -1) ? items[0][colServiceFeeDays] : '';
  var daysRaw   = (colDays !== -1) ? items[0][colDays] : '';
  var serviceFeeDays = Number(sfDaysRaw);
  if (!serviceFeeDays || isNaN(serviceFeeDays)) serviceFeeDays = Number(daysRaw);
  if (!serviceFeeDays || isNaN(serviceFeeDays)) serviceFeeDays = 1;
  var totalServiceFee = serviceFee * serviceFeeDays;

  var subtotalExclTax = 0, totalTax = 0;
  for (var i = 0; i < items.length; i++) {
    var row = items[i];
    var itemId = row[colItemID];
    var qty = Number(row[colQty]) || 0;
    var days = (colDays !== -1 && row[colDays] !== '') ? Number(row[colDays]) : 1;
    var storedPrice = Number(row[colPrice]) || 0;
    var disc = Number(row[colDisc]) || 0;
    var tax = Number(row[colTax]) || 0;

    var freshPriceAtQuote = (colFreshPrice !== -1) ? row[colFreshPrice] : null;
    var actualPrice = storedPrice;
    if (freshPriceAtQuote !== null && freshPriceAtQuote !== undefined && freshPriceAtQuote !== '') {
      actualPrice = Number(freshPriceAtQuote) || storedPrice;
    } else {
      var currentFresh = priceMap[itemId];
      if (currentFresh !== undefined && currentFresh !== null) actualPrice = Number(currentFresh) || storedPrice;
    }
    if (isNaN(actualPrice)) actualPrice = 0;

    var lineSubtotal = qty * days * actualPrice * (1 - disc);
    subtotalExclTax += lineSubtotal;
    totalTax += lineSubtotal * tax;
  }

  var grandTotal = subtotalExclTax + totalTax + totalServiceFee;
  if (isNaN(grandTotal) || !isFinite(grandTotal)) grandTotal = 0;
  return grandTotal;
}

// ==================== LISTS & LOOKUPS ====================
function getInvoiceList() {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName('Invoices');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var invoices = [];
    for (var i = 1; i < data.length; i++) {
      var num = data[i][0];
      if (num && num.toString().trim() !== '' && invoices.indexOf(num) === -1) invoices.push(num);
    }
    return invoices;
  } catch (error) {
    console.error('getInvoiceList error:', error);
    return [];
  }
}

function getInvoicePDFUrl(invoiceNumber) {
  var folderIter = DriveApp.getFoldersByName('Invoices PDFs');
  if (folderIter.hasNext()) {
    var folder = folderIter.next();
    var files = folder.getFilesByName(invoiceNumber + '.pdf');
    if (files.hasNext()) return files.next().getUrl();
  }
  generateInvoicePDF(invoiceNumber, false);
  folderIter = DriveApp.getFoldersByName('Invoices PDFs');
  if (folderIter.hasNext()) {
    var folder = folderIter.next();
    var files = folder.getFilesByName(invoiceNumber + '.pdf');
    if (files.hasNext()) return files.next().getUrl();
  }
  throw new Error('Failed to generate or locate PDF for ' + invoiceNumber);
}

function getInvoicesListWithDetails() {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName('Invoices');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var result = [], seen = {};
    for (var i = 1; i < data.length; i++) {
      var num = data[i][0];
      if (num && !seen[num]) {
        seen[num] = true;
        var clientName = 'Unknown';
        try { clientName = getClientName(data[i][4]) || 'Unknown'; } catch (e) {}
        result.push({
          invoiceNumber: String(num),
          date: data[i][2] ? new Date(data[i][2]).toISOString() : null,
          clientId: data[i][4] !== undefined ? String(data[i][4]) : null,
          clientName: clientName,
          status: data[i][10] || 'Unpaid',
          quoteNumber: data[i][1] ? String(data[i][1]) : ''
        });
      }
    }
    return result;
  } catch (e) {
    Logger.log('CRITICAL error in getInvoicesListWithDetails: ' + e);
    return [];
  }
}

function getUnpaidInvoicesWithClient() {
  var ss = SpreadsheetApp.getActive();
  var clientMap = {};
  var cliSheet = ss.getSheetByName('Clients');
  if (cliSheet) {
    var cData = cliSheet.getDataRange().getValues();
    for (var i = 1; i < cData.length; i++) {
      var id = String(cData[i][0] || '').trim();
      if (id) clientMap[id] = cData[i][1] || 'Unknown';
    }
  }

  var invSheet = ss.getSheetByName('Invoices');
  if (!invSheet) {
    Logger.log('❌ Invoices sheet not found.');
    return [];
  }
  var data = invSheet.getDataRange().getValues();

  var result = [], seen = {};
  for (var i = 1; i < data.length; i++) {
    var invNum = String(data[i][0] || '').trim();
    var status = String(data[i][10] || '').trim();
    var clientId = String(data[i][4] || '').trim();
    if (!invNum || seen[invNum]) continue;
    if (status !== 'Unpaid') continue;
    seen[invNum] = true;
    result.push({
      invoiceNum: invNum, clientId: clientId,
      clientName: clientMap[clientId] || 'Unknown',
      date: data[i][2] ? new Date(data[i][2]).toISOString() : null
    });
  }
  result.sort(function (a, b) { return a.invoiceNum.localeCompare(b.invoiceNum); });

  Logger.log('=== getUnpaidInvoicesWithClient ===');
  Logger.log('Total invoices scanned: ' + (data.length - 1));
  Logger.log('Unpaid invoices found: ' + result.length);
  return result;
}

// ==================== MARK PAID ====================
function markInvoicePaid(invoiceNumber, paymentDate, paymentReference) {
  if (!invoiceNumber) throw new Error('Invoice number is required.');
  paymentReference = (paymentReference || '').toString().trim();

  var sheet = SpreadsheetApp.getActive().getSheetByName('Invoices');
  if (!sheet) throw new Error('Invoices sheet not found.');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var invCol = headers.indexOf('Invoice#');
  var statusCol = headers.indexOf('PaymentStatus');
  var dateCol = headers.indexOf('PaymentDate');
  var refCol = headers.indexOf('PaymentReference');

  if (dateCol === -1) {
    var lastCol = sheet.getLastColumn();
    sheet.getRange(1, lastCol + 1).setValue('PaymentDate');
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    dateCol = headers.indexOf('PaymentDate');
  }
  if (refCol === -1) {
    var lastCol2 = sheet.getLastColumn();
    sheet.getRange(1, lastCol2 + 1).setValue('PaymentReference');
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    refCol = headers.indexOf('PaymentReference');
  }
  if (invCol === -1 || statusCol === -1) throw new Error('Required columns not found.');

  var payDate = paymentDate ? new Date(paymentDate) : new Date();
  var rowsUpdated = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][invCol] === invoiceNumber) {
      sheet.getRange(i + 1, statusCol + 1).setValue('Paid');
      sheet.getRange(i + 1, dateCol + 1).setValue(payDate);
      if (refCol !== -1) sheet.getRange(i + 1, refCol + 1).setValue(paymentReference);
      rowsUpdated++;
    }
  }
  if (rowsUpdated === 0) throw new Error('Invoice number not found or already paid.');

  try { generateInvoicePDF(invoiceNumber, false); } catch (e) { Logger.log('PDF regen error: ' + e); }
  try { generatePaymentReceipt(invoiceNumber, payDate); } catch (e) { Logger.log('Receipt gen error: ' + e); }

  return 'Updated ' + rowsUpdated + ' line item(s) for invoice ' + invoiceNumber +
         ' and generated receipt.' + (paymentReference ? ' Reference: ' + paymentReference : '');
}

// ==================== SEND INVOICE PDF ====================
function sendInvoicePDF(invoiceNumber, emailAddress) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Invoices');
  var data = sheet.getDataRange().getValues();
  var found = false;
  var clientId = '', invoiceDate = '', dueDate = '', quoteRef = '';

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === invoiceNumber) {
      found = true;
      clientId = data[i][4];
      invoiceDate = data[i][2];
      dueDate = data[i][3];
      quoteRef = data[i][1] || '';
      break;
    }
  }
  if (!found) throw new Error('Invoice number not found.');

  var clientName = getClientName(clientId);
  var clientEmail = getClientEmail(clientId);

  var focalEmail = '';
  if (quoteRef) {
    var quoteSheet = ss.getSheetByName('Quotes');
    var quoteData = quoteSheet.getDataRange().getValues();
    var quoteCodes = quoteRef.split(',').map(function(s) { return s.trim(); });
    for (var q = 1; q < quoteData.length; q++) {
      if (quoteCodes.indexOf(quoteData[q][0]) > -1) {
        var fe = quoteData[q][18] || '';
        if (fe) { focalEmail = fe; break; }
      }
    }
  }

  var recipientEmail = (emailAddress && emailAddress.trim() !== '')
    ? emailAddress.trim()
    : (focalEmail || clientEmail);
  if (!recipientEmail) throw new Error('No email available.');

  var pdfUrl = generateInvoicePDF(invoiceNumber, false);
  var folderIterator = DriveApp.getFoldersByName('Invoices PDFs');
  if (!folderIterator.hasNext()) throw new Error('Invoices PDFs folder not found.');
  var folder = folderIterator.next();
  var files = folder.getFilesByName(invoiceNumber + '.pdf');
  if (!files.hasNext()) throw new Error('PDF file not found.');
  var pdfBlob = files.next().getBlob();

  var emailBody = 'Dear ' + clientName + ',\n\nPlease find attached your invoice ' + invoiceNumber + '.\n\n';
  emailBody += 'Invoice Date: ' + (new Date(invoiceDate).toDateString()) + '\n';
  emailBody += 'Due Date: ' + (new Date(dueDate).toDateString()) + '\n\n';
  emailBody += 'Please make payment by the due date.\n\nBest regards,\n' + getConfig('BusinessName', 'Crystal Caterers and Events');

  GmailApp.sendEmail(recipientEmail, 'Invoice ' + invoiceNumber, emailBody, { attachments: [pdfBlob] });

  return { clientName: clientName, clientEmail: clientEmail, usedEmail: recipientEmail, pdfUrl: pdfUrl };
}
