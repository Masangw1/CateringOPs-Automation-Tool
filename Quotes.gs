// =====================================================================
// CRYSTAL CATERERS & EVENTS — QUOTES (PDF & SEND)
// Quotes.gs — Quote PDF generation, sending, date/number updates,
//             quote lists and delivery-day lookups
// =====================================================================

// ==================== QUOTE PDF GENERATION ====================
function generateQuotePDF(quoteNum, sendEmail, includeMenu) {
  sendEmail = sendEmail || false;
  includeMenu = (includeMenu !== undefined) ? includeMenu : false;

  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('Quotes');
    if (!sheet) throw new Error('Quotes sheet not found.');

    var data = sheet.getDataRange().getValues();
    var items = [];
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === quoteNum) items.push(data[i]);
    }
    if (items.length === 0) throw new Error('No items found for quote: ' + quoteNum);

    var rawDate = items[0][1];
    var quoteDate;
    if (rawDate instanceof Date) quoteDate = rawDate;
    else if (typeof rawDate === 'string') quoteDate = new Date(rawDate);
    else quoteDate = new Date();

    var priceMap = getAllMenuPrices();
    var clientId = items[0][2];
    var clientName = getClientName(clientId);
    var focalPerson = items[0][10] || '';
    var serviceFee = items[0][9] || 0;
    var serviceFeeDays = items[0][20] || 1;
    var serviceDescription = items[0][12] || '';
    var showItemDescriptions = items[0][14] || false;
    var actualIncludeMenu = items[0][11] || false;
    var menuOverride = items[0][16] || '';

    var subtotalExclTax = 0, totalTax = 0;
    var lineItems = [];

    for (var i = 0; i < items.length; i++) {
      var row = items[i];
      var itemId = row[3];
      var qty = row[4];
      var days = row[19] || 1;
      var storedPrice = row[5];
      var freshPriceAtQuote = row[12] || null;
      var disc = row[6] || 0;
      var tax = row[7] || 0;

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
        disc: disc, tax: tax, isManualOverride: isManualOverride,
        freshPriceAtQuote: freshPriceAtQuoteNum,
        itemSubtotal: (qty * days) * actualPrice * (1 - disc) * (1 + tax),
        itemSubtotalExclTax: itemSubtotalExclTax,
        itemDescriptionOverride: row[17] || ''
      });
    }

    var foodLines = [];
    var equipmentLines = [];
    var equipmentTotalExclTax = 0;
    for (var i = 0; i < lineItems.length; i++) {
      var idUpper = String(lineItems[i].itemId || '').toUpperCase();
      if (idUpper.indexOf('EQ') === 0) {
        equipmentLines.push(lineItems[i]);
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
    html += 'body { font-family: "Georgia", serif; max-width: 750px; margin: 0 auto; padding: 30px 40px 80px 40px; color: #1e1e1e; }';
    html += '.letterhead { border-bottom: 4px solid #d4a373; padding-bottom: 15px; margin-bottom: 25px; display: flex; align-items: center; justify-content: space-between; }';
    html += '.brand-left { display: flex; align-items: center; gap: 20px; }';
    html += '.logo { max-height: 76px; width: auto; }';
    html += '.business-name { font-size: 27px; font-weight: bold; color: #2c3e50; letter-spacing: 1px; }';
    html += '.tagline { font-size: 13px; color: #7f8c8d; font-style: italic; letter-spacing: 0.5px; margin-top: 2px; }';
    html += '.business-details { font-size: 12px; color: #5d6d7e; line-height: 1.5; text-align: right; }';
    html += '.business-details span { display: block; }';
    html += '.doc-title { font-size: 31px; font-weight: bold; color: #d4a373; text-align: center; letter-spacing: 4px; margin: 10px 0 5px 0; text-transform: uppercase; }';
    html += '.doc-sub { text-align: center; color: #7f8c8d; font-size: 13px; margin-bottom: 25px; }';
    html += '.client-box { background: #f8f5f0; padding: 15px 20px; border-radius: 6px; margin-bottom: 20px; border-left: 6px solid #d4a373; }';
    html += '.client-box strong { color: #2c3e50; }';
    html += 'table { width: 100%; border-collapse: collapse; margin: 25px 0; font-size: 14px; }';
    html += 'th { background: #2c3e50; color: white; padding: 12px 10px; text-align: left; font-weight: 600; }';
    html += 'td { padding: 12px 10px; border-bottom: 1px solid #e5e8eb; }';
    html += 'tr:last-child td { border-bottom: none; }';
    html += '.totals { text-align: right; margin-top: 15px; font-size: 17px; }';
    html += '.totals .grand { font-size: 25px; font-weight: bold; color: #d4a373; }';
    html += '.menu-section { margin: 25px 0 15px 0; page-break-inside: avoid; break-inside: avoid; }';
    html += '.menu-title { font-weight: bold; font-size: 15px; color: #2c3e50; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 2px; }';
    html += '.item-description { font-size: 11px; color: #7f8c8d; font-style: italic; display: block; margin-top: 2px; }';
    html += '@page { margin: 40px 40px 45px 40px; }';
    html += '.pdf-footer { position: fixed; bottom: 0; left: 0; right: 0; text-align: center; font-size: 10px; color: #95a5a6; border-top: 1px solid #e5e8eb; padding: 3px 40px; background: white; line-height: 1.25; }';
    html += '.pdf-footer .line1 { font-size: 10px; color: #7f8c8d; }';
    html += '.pdf-footer .line2 { font-size: 9px; color: #95a5a6; margin-top: 2px; }';
    html += '</style></head><body>';

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

    var formattedDate = Utilities.formatDate(quoteDate, Session.getScriptTimeZone(), 'dd MMM yyyy');
    html += '<div class="doc-title">Quotation</div>';
    html += '<div class="doc-sub">' + quoteNum + '  |  ' + formattedDate + '</div>';

    html += '<div class="client-box">';
    html += '  <strong>Prepared for:</strong> ' + clientName;
    if (focalPerson) html += '<br><strong>Focal Person:</strong> ' + focalPerson;
    if (serviceDescription) html += '<br><br><strong>Service Description:</strong><br>' + serviceDescription;
    html += '</div>';

    html += '<table>';
    html += '  <thead><tr>';
    html += '    <th>Item</th>';
    html += '    <th style="text-align:center;">Qty</th>';
    html += '    <th style="text-align:center;">Days</th>';
    html += '    <th style="text-align:right;">Unit Price</th>';
    html += '    <th style="text-align:right;">Disc.</th>';
    html += '    <th style="text-align:right;">Tax</th>';
    html += '    <th style="text-align:right;">Subtotal</th>';
    html += '  </tr></thead>';
    html += '  <tbody>';

    for (var idx = 0; idx < foodLines.length; idx++) {
      var li = foodLines[idx];
      var itemDetails = getMenuItemDetails(li.itemId);
      var itemName = li.itemId;
      var itemDescription = '';
      var itemDescriptionOverride = li.itemDescriptionOverride;
      if (itemDescriptionOverride) {
        itemDescription = itemDescriptionOverride;
        if (itemDetails) itemName = itemDetails.name || li.itemId;
      } else if (itemDetails) {
        itemName = itemDetails.name || li.itemId;
        itemDescription = itemDetails.description || '';
      }
      var discPercent = li.disc * 100;
      var taxPercent = li.tax * 100;
      html += '    <tr>';
      html += '      <td><strong>' + li.itemId + '</strong>';
      html += '        <br><span style="font-size:12px;color:#2c3e50;font-weight:500;">' + itemName + '</span>';
      if (showItemDescriptions && itemDescription) {
        html += '        <br><span class="item-description">' + itemDescription + '</span>';
      }
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
    if (equipmentTotalExclTax > 0) {
      html += '  <div>AV Equipment Hire: MK' + formatMoney(equipmentTotalExclTax) + '</div>';
    }
    if (hasTax) html += '  <div>VAT: MK' + formatMoney(totalTax) + '</div>';
    if (totalServiceFee > 0) {
      html += '  <div>Service Fee: MK' + formatMoney(totalServiceFee) + (serviceFeeDays > 1 ? ' (MK' + formatMoney(serviceFee) + ' x ' + serviceFeeDays + ' days)' : '') + '</div>';
    }
    html += '  <div style="border-top: 2px solid #e5e8eb; padding-top: 8px; margin-top: 8px;">';
    html += '    <span class="grand">Total: MK' + formatMoney(grandTotal) + '</span>';
    html += '  </div>';
    html += '  <div style="font-size:13px;color:#7f8c8d;margin-top:4px;">Valid for 7 days. Payment due within 14 days of invoice.</div>';
    html += '  <div style="font-size:11px;color:#27ae60;margin-top:4px;">* Prices reflect current menu rates unless manually adjusted.</div>';
    html += '</div>';

    var useMenu = includeMenu || actualIncludeMenu;
    if (useMenu) {
      var menuGroups = [];
      if (menuOverride) {
        var overrideBlocks = String(menuOverride).split(/\n\s*\n/);
        var bi = 0;
        while (bi < overrideBlocks.length) {
          var block = overrideBlocks[bi];
          var lines = block.split('\n').map(function(s) { return s.trim(); }).filter(function(s) { return s.length > 0; });
          if (lines.length === 0) { bi++; continue; }
          if (lines.length === 1 && bi + 1 < overrideBlocks.length) {
            var nextLines = overrideBlocks[bi + 1].split('\n').map(function(s) { return s.trim(); }).filter(function(s) { return s.length > 0; });
            if (nextLines.length > 1) {
              menuGroups.push({ name: lines[0], items: nextLines });
              bi += 2;
              continue;
            }
          }
          menuGroups.push({ name: lines[0], items: lines.slice(1) });
          bi++;
        }
      } else {
        for (var j = 0; j < items.length; j++) {
          var itemId = items[j][3];
          var includeInMenu = items[j][15] || false;
          if (!includeInMenu) continue;
          var itemDetails = getMenuItemDetails(itemId);
          var itemName = itemDetails ? itemDetails.name : itemId;
          var desc = itemDetails ? itemDetails.description : '';
          if (!desc) {
            var overrideDesc = items[j][17] || '';
            if (overrideDesc) desc = overrideDesc;
            else desc = itemName;
          }
          var parts = desc.split(/[,;\n]/).map(function(s) { return s.trim(); });
          while (parts.length > 0 && parts[parts.length - 1] === '') parts.pop();
          while (parts.length > 0 && parts[0] === '') parts.shift();
          var uniqueParts = [];
          var seen = {};
          for (var k = 0; k < parts.length; k++) {
            var lower = parts[k].toLowerCase();
            if (!seen[lower]) { seen[lower] = true; uniqueParts.push(parts[k]); }
          }
          if (uniqueParts.length > 0) menuGroups.push({ name: itemName, items: uniqueParts });
          else menuGroups.push({ name: itemName, items: [desc] });
        }
      }

      var mergedGroups = [];
      var groupMap = {};
      menuGroups.forEach(function(group) {
        var key = group.name.toLowerCase().trim();
        if (!groupMap[key]) {
          groupMap[key] = { name: group.name, items: [] };
          mergedGroups.push(groupMap[key]);
        }
        group.items.forEach(function(item) {
          if (groupMap[key].items.indexOf(item) === -1) groupMap[key].items.push(item);
        });
      });
      menuGroups = mergedGroups;

      if (menuGroups.length > 0) {
        html += '<div class="menu-section" style="page-break-inside: avoid; break-inside: avoid;">';
        html += '  <div class="menu-title">Menu</div>';

        var columns = [[], [], []];
        var numCols = Math.min(3, menuGroups.length);
        for (var i = 0; i < numCols; i++) columns[i].push(menuGroups[i]);
        var colIdx = 2;
        for (var i = 3; i < menuGroups.length; i++) {
          columns[colIdx].push(menuGroups[i]);
          colIdx--;
          if (colIdx < 0) colIdx = 2;
        }

        var activeCols = Math.min(3, menuGroups.length);
        var colWidth = Math.floor(100 / activeCols) + '%';

        html += '  <table style="width:100%; border-collapse:collapse; margin:0 0 18px 0; page-break-inside: avoid; break-inside: avoid;">';
        html += '    <tr>';
        for (var c = 0; c < activeCols; c++) {
          html += '      <td style="width:' + colWidth + '; padding:0 15px 0 0; border:none; vertical-align:top; font-size:12px; color:#5d6d7e; line-height:1.5; page-break-inside: avoid;">';
          var colGroups = columns[c];
          for (var g = 0; g < colGroups.length; g++) {
            if (g > 0) html += '<br><br>';
            html += '<span style="font-weight:bold; font-size:13px; color:#2c3e50;">' + colGroups[g].name + '</span>';
            for (var h = 0; h < colGroups[g].items.length; h++) {
              html += '<br>' + colGroups[g].items[h];
            }
          }
          html += '</td>';
        }
        html += '    </tr>';
        html += '  </table>';
        html += '</div>';
      }
    }

    html += '<div class="pdf-footer" style="border-top: 1px solid #e5e8eb; padding-top: 5px; margin-top: 5px;">';
    html += '  <div class="line1">Thank you for your business. We pride ourselves on fresh, locally-sourced ingredients.</div>';
    html += '  <div class="line2">Company Reg: ' + escapeHtml(getConfig('CompanyReg')) + ' &nbsp;|&nbsp; Withholding Tax Certificate No: ' + escapeHtml(getConfig('WHTCert')) + '</div>';
    html += '</div>';
    html += '</body></html>';

    var blob = Utilities.newBlob(html, 'text/html', quoteNum + '.html');
    var pdf = blob.getAs('application/pdf');

    var folderIterator = DriveApp.getFoldersByName('Catering PDFs');
    var folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder('Catering PDFs');

    var existingFiles = folder.getFilesByName(quoteNum + '.pdf');
    while (existingFiles.hasNext()) existingFiles.next().setTrashed(true);

    var file = folder.createFile(pdf).setName(quoteNum + '.pdf');
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    Logger.log('ERROR in generateQuotePDF: ' + e.toString());
    throw e;
  }
}

// ==================== SEND QUOTE ====================
function sendQuotePDF(quoteNumber) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Quotes');
  if (!sheet) throw new Error('Quotes sheet not found.');

  var data = sheet.getDataRange().getValues();
  var found = false, clientId = '', includeMenu = false, focalEmail = '';
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === quoteNumber) {
      found = true;
      clientId = data[i][2];
      includeMenu = data[i][11] || false;
      focalEmail = data[i][18] || '';
      break;
    }
  }
  if (!found) throw new Error('Quote number not found.');

  var clientName = getClientName(clientId);
  var clientEmail = getClientEmail(clientId);
  var recipientEmail = focalEmail || clientEmail;
  if (!recipientEmail) throw new Error('No email found for this client or focal person.');

  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(recipientEmail)) {
    throw new Error('Invalid email address: "' + recipientEmail + '". Please update the Clients sheet or Focal Person Email.');
  }

  var pdfUrl = generateQuotePDF(quoteNumber, true, includeMenu);
  return { clientName: clientName, clientEmail: recipientEmail, pdfUrl: pdfUrl };
}

// ==================== QUOTE LIST & LISTS-WITH-DETAILS ====================
function getQuotesListWithDetails() {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName('Quotes');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var result = [];
    var seen = {};
    for (var i = 1; i < data.length; i++) {
      var num = data[i][0];
      if (num && !seen[num]) {
        seen[num] = true;
        var clientName = 'Unknown';
        try { clientName = getClientName(data[i][2]) || 'Unknown'; } catch (e) {}
        result.push({
          quoteNumber: String(num),
          date: data[i][1] ? new Date(data[i][1]).toISOString() : null,
          clientId: data[i][2] !== undefined ? String(data[i][2]) : null,
          clientName: clientName,
          status: data[i][8] || 'Draft',
          serviceFee: Number(data[i][9]) || 0,
          focalPerson: data[i][10] || ''
        });
      }
    }
    return result;
  } catch (e) {
    Logger.log('CRITICAL error in getQuotesListWithDetails: ' + e);
    return [];
  }
}

function getQuotesWithoutDeliveryNotes() {
  var ss = SpreadsheetApp.getActive();
  var qSheet = ss.getSheetByName('Quotes');
  if (!qSheet) return [];
  var qData = qSheet.getDataRange().getValues();
  var qHeaders = qData[0].map(function(h){ return String(h).trim(); });
  var qColQ = qHeaders.indexOf('Quote#');
  if (qColQ === -1) qColQ = 0;
  var qColStatus = qHeaders.indexOf('Status');
  var qColSFD = qHeaders.indexOf('ServiceFeeDays');
  var qColDays = qHeaders.indexOf('Days');

  var dnSheet = ss.getSheetByName('DeliveryNotes');
  var dnHeaders = dnSheet ? dnSheet.getDataRange().getValues()[0].map(function(h){ return String(h).trim(); }) : [];
  var dnColQ = dnHeaders.indexOf('Quote#');
  if (dnColQ === -1) dnColQ = 1;

  var dnCount = {};
  if (dnSheet) {
    var dnData = dnSheet.getDataRange().getValues();
    for (var i = 1; i < dnData.length; i++) {
      var qn = String(dnData[i][dnColQ] || '').trim();
      if (qn) dnCount[qn] = (dnCount[qn] || 0) + 1;
    }
  }

  var result = [], seen = {};
  for (var i = 1; i < qData.length; i++) {
    var num = qData[i][qColQ];
    if (!num) continue;
    var numStr = String(num).trim();
    if (seen[numStr]) continue;
    var status = (qColStatus !== -1) ? String(qData[i][qColStatus] || '').trim() : '';
    if (status === 'Superseded') continue;

    var sfd = (qColSFD !== -1) ? Number(qData[i][qColSFD]) || 0 : 0;
    var d   = (qColDays !== -1) ? Number(qData[i][qColDays]) || 0 : 0;
    var totalDays = (sfd > 0) ? sfd : (d > 0 ? d : 1);

    var done = dnCount[numStr] || 0;
    if (done < totalDays) {
      seen[numStr] = true;
      result.push(numStr);
    }
  }
  return result;
}

// ==================== QUOTE DATE / NUMBER UPDATES ====================
function updateQuoteDateAndPDF(quoteNumber, newDate) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Quotes');
  var data = sheet.getDataRange().getValues();

  var rowsUpdated = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === quoteNumber) {
      sheet.getRange(i + 1, 2).setValue(newDate);
      rowsUpdated++;
    }
  }
  if (rowsUpdated === 0) throw new Error('Quote not found: ' + quoteNumber);

  SpreadsheetApp.flush();

  var updatedData = sheet.getDataRange().getValues();
  var includeMenu = false;
  for (var i = 1; i < updatedData.length; i++) {
    if (updatedData[i][0] === quoteNumber) {
      includeMenu = updatedData[i][11] || false;
      break;
    }
  }
  var pdfUrl = generateQuotePDF(quoteNumber, false, includeMenu);
  Logger.log('PDF updated for ' + quoteNumber + ' with date: ' + newDate);
  return pdfUrl;
}

function updateQuoteDate(quoteNumber, newDate) {
  quoteNumber = quoteNumber || 'Q-2026-022';
  newDate = newDate || '2026-09-8';

  if (!quoteNumber) throw new Error('Quote number is required.');

  var dateObj;
  if (newDate instanceof Date) dateObj = newDate;
  else dateObj = new Date(newDate);
  if (isNaN(dateObj.getTime())) throw new Error('Invalid date: ' + newDate);

  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Quotes');
  if (!sheet) throw new Error('Quotes sheet not found.');

  var data = sheet.getDataRange().getValues();
  var rowsUpdated = 0;
  var includeMenu = false;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === quoteNumber) {
      sheet.getRange(i + 1, 2).setValue(dateObj);
      rowsUpdated++;
      if (rowsUpdated === 1) includeMenu = data[i][11] || false;
    }
  }

  if (rowsUpdated === 0) throw new Error('Quote not found: ' + quoteNumber);

  Logger.log('Updated ' + rowsUpdated + ' row(s) for ' + quoteNumber + ' to ' + dateObj);
  var pdfUrl = generateQuotePDF(quoteNumber, false, includeMenu);
  Logger.log('PDF regenerated: ' + pdfUrl);
  return pdfUrl;
}

function updateQuoteNumber(oldNumber, newNumber) {
  var ss = SpreadsheetApp.getActive();
  var quoteSheet = ss.getSheetByName('Quotes');
  var invoiceSheet = ss.getSheetByName('Invoices');

  if (!oldNumber || !newNumber) throw new Error('Both old and new quote numbers are required.');
  if (oldNumber === newNumber) throw new Error('New number must be different from old number.');

  var quoteData = quoteSheet.getDataRange().getValues();
  var foundOld = false;
  for (var i = 1; i < quoteData.length; i++) {
    if (quoteData[i][0] === oldNumber) { foundOld = true; break; }
  }
  if (!foundOld) throw new Error('Quote ' + oldNumber + ' not found.');

  for (var i = 1; i < quoteData.length; i++) {
    if (quoteData[i][0] === newNumber) {
      throw new Error('Quote ' + newNumber + ' already exists. Choose a different number.');
    }
  }

  var rowsUpdated = 0;
  for (var i = 1; i < quoteData.length; i++) {
    if (quoteData[i][0] === oldNumber) {
      quoteSheet.getRange(i + 1, 1).setValue(newNumber);
      rowsUpdated++;
    }
  }

  if (invoiceSheet) {
    var invoiceData = invoiceSheet.getDataRange().getValues();
    for (var i = 1; i < invoiceData.length; i++) {
      var cellValue = invoiceData[i][1];
      if (cellValue && typeof cellValue === 'string') {
        var parts = cellValue.split(',').map(function(s) { return s.trim(); });
        var changed = false;
        for (var j = 0; j < parts.length; j++) {
          if (parts[j] === oldNumber) { parts[j] = newNumber; changed = true; }
        }
        if (changed) invoiceSheet.getRange(i + 1, 2).setValue(parts.join(', '));
      }
    }
  }

  var folders = ['Catering PDFs', 'Invoiced Quotes'];
  for (var f = 0; f < folders.length; f++) {
    var folderIter = DriveApp.getFoldersByName(folders[f]);
    if (folderIter.hasNext()) {
      var folder = folderIter.next();
      var files = folder.getFilesByName(oldNumber + '.pdf');
      while (files.hasNext()) files.next().setName(newNumber + '.pdf');
    }
  }

  var updatedData = quoteSheet.getDataRange().getValues();
  var includeMenu = false;
  for (var i = 1; i < updatedData.length; i++) {
    if (updatedData[i][0] === newNumber) {
      includeMenu = updatedData[i][11] || false;
      break;
    }
  }
  var newPdfUrl = generateQuotePDF(newNumber, false, includeMenu);

  return {
    message: 'Successfully updated quote from ' + oldNumber + ' to ' + newNumber,
    pdfUrl: newPdfUrl, rowsUpdated: rowsUpdated
  };
}

// ==================== QUOTE DELIVERY DAYS (helper for DN form) ====================
function getQuoteDeliveryDays(quoteNumber) {
  if (!quoteNumber) return 1;
  var sheet = SpreadsheetApp.getActive().getSheetByName('Quotes');
  if (!sheet) return 1;
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).trim(); });
  var colQ = headers.indexOf('Quote#');
  if (colQ === -1) colQ = 0;
  var colSFD = headers.indexOf('ServiceFeeDays');
  var colD = headers.indexOf('Days');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colQ]).trim() === String(quoteNumber).trim()) {
      var sfd = (colSFD !== -1) ? Number(data[i][colSFD]) : 0;
      var d   = (colD !== -1) ? Number(data[i][colD]) : 0;
      if (sfd > 0) return sfd;
      if (d > 0) return d;
      return 1;
    }
  }
  return 1;
}

// ==================== EDITOR TEST WRAPPER ====================
function RUN_testQuotePDF() {
  // ⬇️ Change the quote number to whichever one you're testing
  return generateQuotePDF('Q-2026-056', false, true);
}

function RUN_testQuotePDF_withLog() {
  try {
    var url = generateQuotePDF('Q-2026-056', false, true);
    Logger.log('✅ PDF generated: ' + url);
  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
  }
}
