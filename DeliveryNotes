// =====================================================================
// CRYSTAL CATERERS & EVENTS — DELIVERY NOTES
// DeliveryNotes.gs — Creation, PDF, signing, sending, listing
// =====================================================================

// ==================== NUMBER GENERATION ====================
function getNextDeliveryNoteNumber() {
  return getNextNumber('DN', 'DeliveryNotes');
}

// ==================== EXISTING DAYS LOOKUP ====================
function getExistingDeliveryDays(quoteNumber) {
  if (!quoteNumber) return [];
  var sheet = SpreadsheetApp.getActive().getSheetByName('DeliveryNotes');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).trim(); });
  var colQ = headers.indexOf('Quote#');
  var colDay = headers.indexOf('DeliveryDay');
  if (colQ === -1 || colDay === -1) return [];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colQ] || '').trim() === String(quoteNumber).trim()) {
      var d = Number(data[i][colDay]) || 1;
      if (result.indexOf(d) === -1) result.push(d);
    }
  }
  return result;
}

// ==================== CREATE DELIVERY NOTE ====================
function createDeliveryNote(data) {
  var ss = SpreadsheetApp.getActive();
  var dnSheet = ss.getSheetByName('DeliveryNotes');
  if (!dnSheet) throw new Error('DeliveryNotes sheet not found.');

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error('Another delivery note is currently being created. Please wait a moment and try again.');
  }

  try {
    var quoteNum  = String(data.quoteNumber || '').trim();
    var dayNum    = Number(data.deliveryDay) || 1;
    var totalDays = Number(data.totalDays) || dayNum;

    if (dayNum < 1) throw new Error('Delivery day must be 1 or greater.');
    if (dayNum > totalDays) {
      throw new Error('Delivery day (' + dayNum + ') cannot exceed total days (' + totalDays + ').');
    }

    var headerRow = dnSheet.getRange(1, 1, 1, dnSheet.getLastColumn()).getValues()[0];
    var headers = headerRow.map(function(h){ return String(h).trim(); });
    var colDay   = headers.indexOf('DeliveryDay');
    var colTotal = headers.indexOf('TotalDays');
    if (colDay === -1 || colTotal === -1) {
      throw new Error('DeliveryDay/TotalDays columns missing. Reload the sheet (so onOpen runs ensureColumns) and try again.');
    }

    var existing = dnSheet.getDataRange().getValues();
    if (quoteNum) {
      for (var i = 1; i < existing.length; i++) {
        var rowQ   = String(existing[i][1] || '').trim();
        var rowDay = Number(existing[i][colDay]) || 1;
        if (rowQ === quoteNum && rowDay === dayNum) {
          throw new Error(
            'Quote ' + quoteNum + ' already has a delivery note for day ' + dayNum +
            ' (' + existing[i][0] + ').'
          );
        }
      }
    }

    var baseNum = getNextDeliveryNoteNumber();
    var dnNum   = (dayNum > 1) ? (baseNum + '/' + dayNum) : baseNum;

    var row = [];
    row[0]  = dnNum;
    row[1]  = quoteNum;
    row[2]  = data.invoiceNumber || '';
    row[3]  = data.clientId;
    row[4]  = new Date();
    row[5]  = 'Draft';
    row[6]  = '';
    row[7]  = '';
    row[8]  = JSON.stringify(data.items);
    row[9]  = data.total || 0;
    row[10] = '';
    row[11] = '';
    row[12] = '';
    row[13] = Number(data.serviceFee) || 0;
    row[14] = Number(data.serviceFeeDays) || 1;
    row[colDay]   = dayNum;
    row[colTotal] = totalDays;

    for (var c = 0; c < headers.length; c++) {
      if (row[c] === undefined) row[c] = '';
    }

    dnSheet.appendRow(row);

    var pdfUrl = generateDeliveryNotePDF(dnNum);
    return { dnNum: dnNum, pdfUrl: pdfUrl };

  } finally {
    lock.releaseLock();
  }
}

// ==================== GENERATE DELIVERY NOTE PDF ====================
function generateDeliveryNotePDF(dnNum, includeSignature) {
  includeSignature = includeSignature || false;
  var sheet = SpreadsheetApp.getActive().getSheetByName('DeliveryNotes');
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).trim(); });
  var row = null;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === dnNum) { row = data[i]; break; }
  }
  if (!row) throw new Error('Delivery note not found: ' + dnNum);

  var clientId = row[3];
  var date = row[4];
  var signatureUrl = row[6];
  var signedDate = row[7];
  var items = JSON.parse(row[8]);
  var sigType = row[11] || '';

  var clientName = getClientName(clientId);
  var logoBase64 = getLogoBase64();

  var html = '<html><head><meta charset="UTF-8"><style>';
  html += 'body { font-family: Georgia, serif; max-width: 750px; margin: 30px auto; padding: 20px; }';
  html += '.letterhead { border-bottom: 4px solid #d4a373; padding-bottom: 15px; margin-bottom: 25px; display: flex; align-items: center; justify-content: space-between; }';
  html += '.brand-left { display: flex; align-items: center; gap: 20px; }';
  html += '.logo { max-height: 76px; width: auto; }';
  html += '.business-name { font-size: 27px; font-weight: bold; color: #2c3e50; }';
  html += '.business-details { font-size: 12px; color: #5d6d7e; text-align: right; }';
  html += '.business-details span { display: block; }';
  html += '.doc-title { font-size: 31px; font-weight: bold; color: #2c3e50; text-align: center; letter-spacing: 4px; margin: 10px 0 5px 0; }';
  html += '.doc-sub { text-align: center; color: #7f8c8d; font-size: 13px; margin-bottom: 25px; }';
  html += '.client-box { background: #f8f5f0; padding: 15px; border-radius: 6px; margin-bottom: 20px; border-left: 6px solid #d4a373; }';
  html += 'table { width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 14px; }';
  html += 'th { background: #2c3e50; color: white; padding: 10px; text-align: left; }';
  html += 'td { padding: 10px; border-bottom: 1px solid #e5e8eb; }';
  html += '.signature-box { margin-top: 40px; border-top: 2px solid #ccc; padding-top: 20px; }';
  html += '.signature-box .row { display: flex; justify-content: space-between; margin-bottom: 15px; }';
  html += '.signature-box .label { font-weight: bold; }';
  html += '.signature-box .line { border-bottom: 1px solid #333; min-width: 150px; display: inline-block; }';
  html += '.signature-box .signed { color: #27ae60; font-weight: bold; }';
  html += '@page { margin: 40px 40px 45px 40px; }';
  html += '.pdf-footer { position: fixed; bottom: 0; left: 0; right: 0; text-align: center; font-size: 10px; color: #95a5a6; border-top: 1px solid #e5e8eb; padding: 3px 40px; background: white; line-height: 1.25; }';
  html += '.pdf-footer .line1 { font-size: 10px; color: #7f8c8d; }';
  html += '.pdf-footer .line2 { font-size: 9px; color: #95a5a6; margin-top: 2px; }';
  html += '</style></head><body>';

  html += '<div class="letterhead">';
  html += '  <div class="brand-left">';
  html += '    <img class="logo" src="' + logoBase64 + '" alt="Logo" />';
  html += '    <div><div class="business-name">' + escapeHtml(getConfig('BusinessName', 'CRYSTAL CATERERS AND EVENTS')).toUpperCase() + '</div></div>';
  html += '  </div>';
  html += '  <div class="business-details">';
  html += '    <span>📍 ' + escapeHtml(getConfig('Address')) + '</span>';
  html += '    <span>📱 ' + escapeHtml(getConfig('Phone1')) + '</span>';
  html += '    <span>✉️ ' + escapeHtml(getConfig('Email')) + '</span>';
  html += '  </div>';
  html += '</div>';

  html += '<div class="doc-title">DELIVERY NOTE</div>';
  var dateStr = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'dd MMM yyyy');
  var deliveryDay  = Number(row[headers.indexOf('DeliveryDay')]) || 1;
  var totalDays    = Number(row[headers.indexOf('TotalDays')]) || 1;
  var dayLabel     = (totalDays > 1) ? ('  |  Delivery ' + deliveryDay + ' of ' + totalDays) : '';
  html += '<div class="doc-sub">' + dnNum + '  |  ' + dateStr + dayLabel + '</div>';

  html += '<div class="client-box"><strong>Client:</strong> ' + clientName + '</div>';

  html += '<table><thead><tr><th>Item</th><th>Qty</th><th>Days</th><th>Price</th><th>Total</th></tr></thead><tbody>';
  items.forEach(function(item) {
    var days = item.days || 1;
    var lineTotal = item.qty * days * item.price;
    html += '<tr><td>' + item.itemId + ' - ' + (item.desc || '') + '</td><td>' + item.qty + '</td><td>' + days + '</td><td>MK' + formatMoney(item.price) + '</td><td>MK' + formatMoney(lineTotal) + '</td></tr>';
  });
  html += '</tbody></table>';

  html += '<div class="signature-box">';
  html += '  <div class="row"><span class="label">Received by:</span> <span class="line">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></div>';
  html += '  <div class="row"><span class="label">Signature:</span> ';
  if (includeSignature && signatureUrl) {
    html += '    <img src="' + signatureUrl + '" style="max-height:60px;border:1px solid #ccc;" />';
  } else {
    html += '<span class="line">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>';
  }
  html += '  </div>';
  if (includeSignature && sigType) {
    html += '  <div class="row"><span class="label">Confirmation Method:</span> <span style="font-size:12px;color:#555;">' + sigType + ' signature</span></div>';
  }
  html += '  <div class="row"><span class="label">Date:</span> ';
  if (includeSignature && signedDate) {
    var signedDateStr = Utilities.formatDate(new Date(signedDate), Session.getScriptTimeZone(), 'dd MMM yyyy');
    html += '<span class="signed">' + signedDateStr + '</span>';
  } else {
    html += '<span class="line">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>';
  }
  html += '  </div>';
  if (!includeSignature) {
    html += '  <div style="font-size:12px;color:#888;margin-top:10px;">This delivery note must be signed by the recipient.</div>';
  } else if (signatureUrl) {
    html += '  <div style="font-size:12px;color:#27ae60;margin-top:10px;">✓ Signed and confirmed on ' + Utilities.formatDate(new Date(signedDate), Session.getScriptTimeZone(), 'dd MMM yyyy') + (sigType ? ' via ' + sigType + ' signature' : '') + '</div>';
  }
  html += '</div>';

  html += '<div class="pdf-footer">';
  html += '  <div class="line1">Thank you for your business. We pride ourselves on fresh, locally-sourced ingredients.</div>';
  html += '  <div class="line2">Business Reg: ' + escapeHtml(getConfig('CompanyReg')) + ' &nbsp;|&nbsp; Withholding Tax Certificate No: ' + escapeHtml(getConfig('WHTCert')) + '</div>';
  html += '</div>';
  html += '</body></html>';

  var blob = Utilities.newBlob(html, 'text/html', dnNum + '.html');
  var pdf = blob.getAs('application/pdf');

  var folder = getOrCreateFolder('Delivery Notes PDFs');
  var existingFiles = folder.getFilesByName(dnNum + '.pdf');
  while (existingFiles.hasNext()) existingFiles.next().setTrashed(true);
  var file = folder.createFile(pdf).setName(dnNum + '.pdf');
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ==================== SEND FOR SIGNING ====================
function sendDeliveryNoteEmailWithEmailV2(dnNum, customEmail) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('DeliveryNotes');
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  var rowData = null;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === dnNum) { rowIndex = i; rowData = data[i]; break; }
  }
  if (rowIndex === -1) throw new Error('Delivery note not found.');

  var clientId = rowData[3];
  var clientEmail = (customEmail && customEmail.trim() !== '') ? customEmail.trim() : getClientEmail(clientId);
  if (!clientEmail) throw new Error('No email address available.');

  var token = Utilities.base64Encode(dnNum + '|' + new Date().getTime() + '|' + Math.random());
  sheet.getRange(rowIndex + 1, 11).setValue(token);

  var scriptUrl = getConfig('WebAppUrl', '');
  var signUrl = scriptUrl + '?page=sign&dn=' + encodeURIComponent(dnNum) + '&token=' + encodeURIComponent(token);
  Logger.log('SIGN URL: ' + signUrl);

  var folder = getOrCreateFolder('Delivery Notes PDFs');
  var files = folder.getFilesByName(dnNum + '.pdf');
  if (!files.hasNext()) {
    generateDeliveryNotePDF(dnNum, false);
    files = folder.getFilesByName(dnNum + '.pdf');
    if (!files.hasNext()) throw new Error('PDF could not be generated.');
  }
  var pdfBlob = files.next().getBlob();

  var clientName = getClientName(clientId);
  var emailBody = 'Dear ' + clientName + ',\n\n';
  emailBody += 'Please find attached the delivery note for your order.\n';
  emailBody += 'To confirm receipt, please click the link below and sign electronically:\n';
  emailBody += signUrl + '\n\n';
  emailBody += 'If you have any questions, please contact us.\n\nBest regards,\n' + getConfig('BusinessName', 'Crystal Caterers and Events');

  GmailApp.sendEmail(clientEmail, 'Delivery Note ' + dnNum + ' - Please Sign', emailBody, { attachments: [pdfBlob] });
  sheet.getRange(rowIndex + 1, 6).setValue('Sent');
  return 'Email sent to ' + clientEmail;
}

// ==================== SIGN PAGE DATA ====================
function getDeliveryNoteData(dnNum, token) {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName('DeliveryNotes');
    if (!sheet) {
      return JSON.stringify({ success: false, message: 'DeliveryNotes sheet not found.', data: null });
    }
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return JSON.stringify({ success: false, message: 'DeliveryNotes sheet is empty.', data: null });
    }

    dnNum = String(dnNum).trim();
    token = String(token).trim();

    for (var i = 1; i < data.length; i++) {
      var storedDn = data[i][0] ? String(data[i][0]).trim() : '';
      if (storedDn !== dnNum) continue;

      var found = false;
      for (var col = 0; col < data[i].length; col++) {
        var storedVal = data[i][col] ? String(data[i][col]).trim() : '';
        if (storedVal === token ||
            (storedVal.length > 0 && token.length > 0 &&
             (storedVal.indexOf(token) === 0 || token.indexOf(storedVal) === 0))) {
          found = true;
          break;
        }
      }
      if (!found) {
        return JSON.stringify({ success: false, message: 'Token mismatch for ' + dnNum, data: null });
      }

      var items = [];
      try {
        if (typeof data[i][8] === 'string') items = JSON.parse(data[i][8]);
        else if (Array.isArray(data[i][8])) items = data[i][8];
      } catch (e) {
        return JSON.stringify({ success: false, message: 'Items parse error: ' + e.message, data: null });
      }

      var safeItems = items.map(function(it) {
        return {
          itemId: String(it.itemId || ''), qty: Number(it.qty) || 0,
          price: Number(it.price) || 0, desc: String(it.desc || '')
        };
      });

      var dateStr = '';
      try {
        if (data[i][4] instanceof Date) dateStr = data[i][4].toISOString();
        else if (data[i][4]) dateStr = String(data[i][4]);
      } catch (e) { dateStr = ''; }

      var clientName = '';
      try { clientName = String(getClientName(data[i][3]) || ''); } catch (e) { clientName = ''; }

      return JSON.stringify({
        success: true, message: 'Success',
        data: {
          dnNum: String(data[i][0] || ''), quote: String(data[i][1] || ''),
          invoice: String(data[i][2] || ''), clientId: String(data[i][3] || ''),
          date: dateStr, status: String(data[i][5] || 'Draft'),
          items: safeItems, total: Number(data[i][9]) || 0,
          serviceFee: Number(data[i][13]) || 0, clientName: clientName
        }
      });
    }
    return JSON.stringify({ success: false, message: 'DN not found: ' + dnNum, data: null });
  } catch (e) {
    Logger.log('Exception: ' + e.message);
    return JSON.stringify({ success: false, message: 'Exception: ' + e.message, data: null });
  }
}

// ==================== SAVE SIGNATURE ====================
function saveSignature(dnNum, signatureDataUrl, token, signatureType, confirmationText) {
  signatureType = signatureType || 'Drawn';
  confirmationText = confirmationText || '';

  var sheet = SpreadsheetApp.getActive().getSheetByName('DeliveryNotes');
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  var rowData = null;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === dnNum && data[i][10] === token) {
      rowIndex = i;
      rowData = data[i];
      break;
    }
  }
  if (rowIndex === -1) throw new Error('Invalid delivery note or token.');

  var imageData = signatureDataUrl.split(',')[1];
  var blob = Utilities.newBlob(Utilities.base64Decode(imageData), 'image/png', 'signature.png');

  var folder = getOrCreateFolder('Delivery Signatures');
  var existingFiles = folder.getFilesByName(dnNum + '_signature.png');
  while (existingFiles.hasNext()) existingFiles.next().setTrashed(true);
  var file = folder.createFile(blob).setName(dnNum + '_signature.png');
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var signatureUrl = file.getUrl();

  var now = new Date();
  sheet.getRange(rowIndex + 1, 6).setValue('Signed');
  sheet.getRange(rowIndex + 1, 7).setValue(signatureUrl);
  sheet.getRange(rowIndex + 1, 8).setValue(now);
  sheet.getRange(rowIndex + 1, 12).setValue(signatureType);
  sheet.getRange(rowIndex + 1, 13).setValue(confirmationText);

  generateDeliveryNotePDF(dnNum, true);

  try {
    var clientId = rowData[3];
    var clientEmail = getClientEmail(clientId);
    var clientName = getClientName(clientId);
    if (clientEmail) {
      var pdfFolder = getOrCreateFolder('Delivery Notes PDFs');
      var pdfFiles = pdfFolder.getFilesByName(dnNum + '.pdf');
      var pdfBlob = pdfFiles.hasNext() ? pdfFiles.next().getBlob() : null;
      var emailBody = 'Dear ' + clientName + ',\n\n';
      emailBody += 'Thank you for confirming receipt of your order.\n\n';
      emailBody += 'Please find attached the signed delivery note ' + dnNum + ' for your records.\n';
      emailBody += 'Confirmation method: ' + signatureType + ' signature\n\n';
      emailBody += 'Best regards,\n' + getConfig('BusinessName', 'Crystal Caterers and Events');
      var options = {};
      if (pdfBlob) options.attachments = [pdfBlob];
      GmailApp.sendEmail(clientEmail, 'Signed Delivery Note ' + dnNum + ' – Confirmation', emailBody, options);
    }
  } catch (e) {
    Logger.log('Confirmation email error: ' + e.message);
  }

  return signatureUrl;
}

// ==================== LIST / RESET / PDF URL ====================
function getDeliveryNotesList() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('DeliveryNotes');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var dateVal = data[i][4], signedVal = data[i][7];
    var dateStr = '', signedDateStr = '';
    try {
      if (dateVal instanceof Date) dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'dd MMM yyyy');
      else if (dateVal) dateStr = String(dateVal);
    } catch (e) { dateStr = ''; }
    try {
      if (signedVal instanceof Date) signedDateStr = Utilities.formatDate(signedVal, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
      else if (signedVal) signedDateStr = String(signedVal);
    } catch (e) { signedDateStr = ''; }

    var clientName = '';
    try { clientName = getClientName(data[i][3]) || 'Unknown'; } catch (e) { clientName = 'Unknown'; }

    result.push({
      dnNum: String(data[i][0]), quote: String(data[i][1] || ''),
      invoice: String(data[i][2] || ''), clientId: String(data[i][3] || ''),
      dateStr: dateStr, status: String(data[i][5] || 'Draft'),
      signatureUrl: String(data[i][6] || ''), signedDateStr: signedDateStr,
      total: Number(data[i][9]) || 0, serviceFee: Number(data[i][13]) || 0,
      clientName: clientName
    });
  }
  result.reverse();
  return result;
}

function getDeliveryNotePDFUrl(dnNum) {
  generateDeliveryNotePDF(dnNum, true);
  var folder = getOrCreateFolder('Delivery Notes PDFs');
  var files = folder.getFilesByName(dnNum + '.pdf');
  if (!files.hasNext()) throw new Error('PDF not found for ' + dnNum);
  return files.next().getUrl();
}

function resetDeliveryNoteStatus(dnNum) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('DeliveryNotes');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === dnNum) {
      sheet.getRange(i + 1, 6).setValue('Sent');
      sheet.getRange(i + 1, 7).setValue('');
      sheet.getRange(i + 1, 8).setValue('');
      sheet.getRange(i + 1, 12).setValue('');
      sheet.getRange(i + 1, 13).setValue('');
      try { generateDeliveryNotePDF(dnNum, false); } catch (e) {}
      return 'Reset ' + dnNum + ' to Sent';
    }
  }
  throw new Error('Delivery note not found: ' + dnNum);
}

// ==================== SENDING LISTS & RESEND ====================
function getDeliveryNotesForSending() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('DeliveryNotes');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var status = String(data[i][5] || 'Draft');
    if (status === 'Signed') continue;

    var clientName = '';
    try { clientName = getClientName(data[i][3]) || 'Unknown'; } catch (e) {}

    var dateStr = '';
    try {
      if (data[i][4] instanceof Date) dateStr = Utilities.formatDate(data[i][4], Session.getScriptTimeZone(), 'dd MMM yyyy');
    } catch (e) {}

    result.push({
      dnNum: String(data[i][0]), clientId: String(data[i][3] || ''),
      clientName: clientName, dateStr: dateStr, status: status,
      total: Number(data[i][9]) || 0, serviceFee: Number(data[i][13]) || 0
    });
  }
  result.reverse();
  return result;
}

function resendDeliveryNoteEmail(dnNum, customEmail) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('DeliveryNotes');
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  var rowData = null;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === dnNum) { rowIndex = i; rowData = data[i]; break; }
  }
  if (rowIndex === -1) throw new Error('Delivery note not found.');

  var status = String(rowData[5] || 'Draft');
  if (status === 'Signed') throw new Error('This delivery note is already signed and cannot be resent.');

  var clientId = rowData[3];
  var clientEmail = (customEmail && customEmail.trim() !== '') ? customEmail.trim() : getClientEmail(clientId);
  if (!clientEmail) throw new Error('No email address available.');

  var token = rowData[10];
  if (!token) {
    token = Utilities.base64Encode(dnNum + '|' + new Date().getTime() + '|' + Math.random());
    sheet.getRange(rowIndex + 1, 11).setValue(token);
  }

  var scriptUrl = getConfig('WebAppUrl', '');
  var signUrl = scriptUrl + '?page=sign&dn=' + encodeURIComponent(dnNum) + '&token=' + encodeURIComponent(token);
  Logger.log('RESEND URL: ' + signUrl);

  var folder = getOrCreateFolder('Delivery Notes PDFs');
  var files = folder.getFilesByName(dnNum + '.pdf');
  if (!files.hasNext()) {
    generateDeliveryNotePDF(dnNum, false);
    files = folder.getFilesByName(dnNum + '.pdf');
    if (!files.hasNext()) throw new Error('PDF could not be generated.');
  }
  var pdfBlob = files.next().getBlob();

  var clientName = getClientName(clientId);
  var emailBody = 'Dear ' + clientName + ',\n\n';
  emailBody += 'Please find attached the delivery note for your order.\n';
  emailBody += 'To confirm receipt, please click the link below and sign electronically:\n';
  emailBody += signUrl + '\n\n';
  emailBody += 'Best regards,\n' + getConfig('BusinessName', 'Crystal Caterers and Events');

  GmailApp.sendEmail(clientEmail, 'Delivery Note ' + dnNum + ' - Please Sign', emailBody, { attachments: [pdfBlob] });
  sheet.getRange(rowIndex + 1, 6).setValue('Sent');
  return 'Email sent to ' + clientEmail;
}
