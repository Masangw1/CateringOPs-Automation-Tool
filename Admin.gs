// =====================================================================
// CRYSTAL CATERERS & EVENTS — ADMIN TOOLS
// Admin.gs — Health check, cleanup, duplicate detection, blank rows,
//            supersede, delete-invoices, dashboard stats, test wrappers
// =====================================================================

// ==================== HELPERS ====================
function admin_sheetRows(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return null;
  return { sheet: sheet, data: sheet.getDataRange().getValues(), lastRow: sheet.getLastRow() };
}

function admin_isBlankRow(row) {
  for (var i = 0; i < row.length; i++) {
    if (row[i] !== '' && row[i] !== null && row[i] !== undefined) return false;
  }
  return true;
}

// =====================================================================
// HEALTH CHECK
// =====================================================================

function admin_runHealthCheck() {
  Logger.log('=========================================');
  Logger.log('  HEALTH CHECK REPORT');
  Logger.log('  Generated: ' + new Date().toISOString());
  Logger.log('=========================================\n');

  var summary = { total: 0 };
  summary.total += admin_hc_paidInvoicesNoDate();
  summary.total += admin_hc_invoicedQuotesNoInvoice();
  summary.total += admin_hc_signedDNsNoSignature();
  summary.total += admin_hc_brokenItemRefs();
  summary.total += admin_hc_brokenClientRefs();
  summary.total += admin_hc_duplicateClients();
  summary.total += admin_hc_trailingEmptyRows();

  Logger.log('\n=========================================');
  Logger.log('  SUMMARY: ' + summary.total + ' issue(s) found');
  Logger.log('=========================================');
  return summary;
}

function admin_hc_paidInvoicesNoDate() {
  Logger.log('--- 1. Paid invoices with no PaymentDate ---');
  var r = admin_sheetRows('Invoices');
  if (!r) { Logger.log('  (sheet missing)\n'); return 0; }
  var issues = [];
  for (var i = 1; i < r.data.length; i++) {
    var status = String(r.data[i][10] || '').trim();
    var payDate = r.data[i][13];
    if (status === 'Paid' && (!payDate || payDate === '')) {
      issues.push('  • ' + r.data[i][0] + ' (row ' + (i + 1) + ') — client ' + r.data[i][4]);
    }
  }
  if (issues.length === 0) Logger.log('  ✅ None');
  else issues.forEach(function (s) { Logger.log(s); });
  Logger.log('');
  return issues.length;
}

function admin_hc_invoicedQuotesNoInvoice() {
  Logger.log('--- 2. Invoiced quotes with no matching invoice ---');
  var qr = admin_sheetRows('Quotes');
  var ir = admin_sheetRows('Invoices');
  if (!qr || !ir) { Logger.log('  (sheets missing)\n'); return 0; }

  var invoicedQuoteSet = {};
  for (var i = 1; i < ir.data.length; i++) {
    var ref = ir.data[i][1];
    if (ref) {
      String(ref).split(',').forEach(function (q) {
        invoicedQuoteSet[q.trim().toUpperCase()] = true;
      });
    }
  }

  var issues = [], seen = {};
  for (var i = 1; i < qr.data.length; i++) {
    var qn = String(qr.data[i][0] || '').trim().toUpperCase();
    var status = String(qr.data[i][8] || '').trim();
    if (qn && status === 'Invoiced' && !seen[qn] && !invoicedQuoteSet[qn]) {
      seen[qn] = true;
      issues.push('  • ' + qn + ' — client ' + qr.data[i][2]);
    }
  }
  if (issues.length === 0) Logger.log('  ✅ None');
  else issues.forEach(function (s) { Logger.log(s); });
  Logger.log('');
  return issues.length;
}

function admin_hc_signedDNsNoSignature() {
  Logger.log('--- 3. Signed DNs with no signature image ---');
  var r = admin_sheetRows('DeliveryNotes');
  if (!r) { Logger.log('  (sheet missing)\n'); return 0; }
  var issues = [];
  for (var i = 1; i < r.data.length; i++) {
    var dn = String(r.data[i][0] || '').trim();
    var status = String(r.data[i][5] || '').trim();
    var sig = r.data[i][6];
    if (dn && status === 'Signed' && (!sig || sig === '')) {
      issues.push('  • ' + dn + ' (row ' + (i + 1) + ')');
    }
  }
  if (issues.length === 0) Logger.log('  ✅ None');
  else issues.forEach(function (s) { Logger.log(s); });
  Logger.log('');
  return issues.length;
}

function admin_hc_brokenItemRefs() {
  Logger.log('--- 4. Quote items with missing MenuItems reference ---');

  var KNOWN_LEGACY_IDS = [
    'MF01', 'AF01', 'SN1', 'SN2', 'SN3', 'SN4', 'SN5',
    'RF001', 'RF 03', 'RF 04', 'TFF1', 'TFF2', 'JC01', 'B04', 'B06',
    'AF02', 'AF03', 'MF03', 'BL03', 'LB003', 'LB004', 'BW04', 'L01'
  ];
  var legacySet = {};
  KNOWN_LEGACY_IDS.forEach(function (id) { legacySet[id.toUpperCase()] = true; });

  var qr = admin_sheetRows('Quotes');
  var mr = admin_sheetRows('MenuItems');
  if (!qr || !mr) { Logger.log('  (sheets missing)\n'); return 0; }

  var menuSet = {};
  for (var i = 1; i < mr.data.length; i++) {
    var id = String(mr.data[i][0] || '').trim().toUpperCase();
    if (id) menuSet[id] = true;
  }

  var issues = [], seen = {};
  for (var i = 1; i < qr.data.length; i++) {
    var rawId = String(qr.data[i][3] || '').trim();
    var itemId = rawId.toUpperCase();
    var qn = String(qr.data[i][0] || '').trim();
    if (itemId && !menuSet[itemId] && !legacySet[itemId]) {
      var key = qn + '|' + itemId;
      if (!seen[key]) {
        seen[key] = true;
        issues.push('  • ' + qn + ' uses "' + rawId + '" (not in MenuItems)');
      }
    }
  }
  if (issues.length === 0) Logger.log('  ✅ None');
  else issues.forEach(function (s) { Logger.log(s); });
  Logger.log('');
  return issues.length;
}

function admin_hc_brokenClientRefs() {
  Logger.log('--- 5. Invoices with unknown ClientID ---');
  var ir = admin_sheetRows('Invoices');
  var cr = admin_sheetRows('Clients');
  if (!ir || !cr) { Logger.log('  (sheets missing)\n'); return 0; }

  var clientSet = {};
  for (var i = 1; i < cr.data.length; i++) {
    var id = String(cr.data[i][0] || '').trim().toUpperCase();
    if (id) clientSet[id] = true;
  }

  var issues = [], seen = {};
  for (var i = 1; i < ir.data.length; i++) {
    var inv = String(ir.data[i][0] || '').trim();
    var cid = String(ir.data[i][4] || '').trim().toUpperCase();
    if (inv && cid && !clientSet[cid] && !seen[inv]) {
      seen[inv] = true;
      issues.push('  • ' + inv + ' → client "' + ir.data[i][4] + '" (missing in Clients)');
    }
  }
  if (issues.length === 0) Logger.log('  ✅ None');
  else issues.forEach(function (s) { Logger.log(s); });
  Logger.log('');
  return issues.length;
}

function admin_hc_duplicateClients() {
  Logger.log('--- 6. Duplicate client names ---');
  var r = admin_sheetRows('Clients');
  if (!r) { Logger.log('  (sheet missing)\n'); return 0; }

  var byName = {};
  for (var i = 1; i < r.data.length; i++) {
    var name = String(r.data[i][1] || '').trim();
    if (!name) continue;
    var key = name.toLowerCase();
    if (!byName[key]) byName[key] = [];
    byName[key].push({ id: r.data[i][0], row: i + 1 });
  }

  var issues = [];
  for (var k in byName) {
    if (byName[k].length > 1) {
      var entries = byName[k].map(function (x) { return x.id + ' (row ' + x.row + ')'; });
      issues.push('  • "' + k + '" → ' + entries.join(' + '));
    }
  }
  if (issues.length === 0) Logger.log('  ✅ None');
  else issues.forEach(function (s) { Logger.log(s); });
  Logger.log('');
  return issues.length;
}

function admin_hc_trailingEmptyRows() {
  Logger.log('--- 7. Trailing empty rows in sheets ---');
  var sheets = ['Clients', 'MenuItems', 'Quotes', 'DeliveryNotes', 'Invoices'];
  var issues = [];
  sheets.forEach(function (name) {
    var r = admin_sheetRows(name);
    if (!r) return;
    var blank = 0;
    for (var i = r.data.length - 1; i >= 1; i--) {
      if (admin_isBlankRow(r.data[i])) blank++;
      else break;
    }
    if (blank > 0) issues.push('  • ' + name + ': ' + blank + ' trailing blank row(s)');
  });
  if (issues.length === 0) Logger.log('  ✅ None');
  else issues.forEach(function (s) { Logger.log(s); });
  Logger.log('');
  return issues.length;
}

// =====================================================================
// DUPLICATE DN ROWS
// =====================================================================

function admin_previewDuplicateDNRows() {
  var r = admin_sheetRows('DeliveryNotes');
  if (!r) return { sameDn: [], sameQuote: [], totalIssues: 0 };

  var byDn = {};
  for (var i = 1; i < r.data.length; i++) {
    var dn = String(r.data[i][0] || '').trim();
    if (!dn) continue;
    if (!byDn[dn]) byDn[dn] = [];
    byDn[dn].push({ row: i + 1, data: r.data[i] });
  }
  var sameDn = [];
  for (var dnKey in byDn) {
    if (byDn[dnKey].length > 1) {
      sameDn.push({ dn: dnKey, rows: byDn[dnKey].map(function (x) { return x.row; }) });
    }
  }

  var byQuote = {};
  for (var i = 1; i < r.data.length; i++) {
    var qn = String(r.data[i][1] || '').trim();
    var dn = String(r.data[i][0] || '').trim();
    if (!qn || !dn) continue;
    if (!byQuote[qn]) byQuote[qn] = {};
    byQuote[qn][dn] = true;
  }
  var sameQuote = [];
  for (var qnKey in byQuote) {
    var dns = Object.keys(byQuote[qnKey]);
    if (dns.length > 1) sameQuote.push({ quote: qnKey, dns: dns });
  }

  Logger.log('=== DUPLICATE DN ROW REPORT ===\n');
  Logger.log('Flavor 1 — same DeliveryNote# on multiple rows:');
  if (sameDn.length === 0) Logger.log('  ✅ None');
  sameDn.forEach(function (x) {
    Logger.log('  • ' + x.dn + ' appears at row(s): ' + x.rows.join(', '));
  });
  Logger.log('');
  Logger.log('Flavor 2 — same Quote# on multiple different DNs:');
  if (sameQuote.length === 0) Logger.log('  ✅ None');
  sameQuote.forEach(function (x) {
    Logger.log('  • ' + x.quote + ' → ' + x.dns.join(', '));
  });
  Logger.log('');

  return { sameDn: sameDn, sameQuote: sameQuote, totalIssues: sameDn.length + sameQuote.length };
}

function admin_fixDuplicateDNRows(confirmed) {
  var preview = admin_previewDuplicateDNRows();
  if (preview.totalIssues === 0) {
    Logger.log('Nothing to fix.');
    return { fixed: 0 };
  }
  if (confirmed !== true) {
    Logger.log('⚠️ Preview only. To actually fix, run: admin_fixDuplicateDNRows(true)');
    return { fixed: 0, preview: preview };
  }

  var r = admin_sheetRows('DeliveryNotes');
  var sheet = r.sheet;
  var rowsToDelete = [];

  preview.sameDn.forEach(function (group) {
    var rowInfos = group.rows.map(function (rowNum) {
      return { rowNum: rowNum, date: r.data[rowNum - 1][4] };
    });
    rowInfos.sort(function (a, b) {
      return (new Date(b.date).getTime() || 0) - (new Date(a.date).getTime() || 0);
    });
    rowInfos.slice(1).forEach(function (x) { rowsToDelete.push(x.rowNum); });
  });

  preview.sameQuote.forEach(function (group) {
    var dnInfos = group.dns.map(function (dn) {
      var rowNum = -1, date = null;
      for (var i = 1; i < r.data.length; i++) {
        if (String(r.data[i][0]).trim() === dn) {
          rowNum = i + 1;
          date = r.data[i][4];
          break;
        }
      }
      return { dn: dn, rowNum: rowNum, date: date };
    }).filter(function (x) { return x.rowNum > 0; });

    dnInfos.sort(function (a, b) {
      return (new Date(b.date).getTime() || 0) - (new Date(a.date).getTime() || 0);
    });

    dnInfos.slice(1).forEach(function (keepMe) {
      for (var i = r.data.length - 1; i >= 1; i--) {
        if (String(r.data[i][0]).trim() === keepMe.dn) rowsToDelete.push(i + 1);
      }
    });
  });

  rowsToDelete = rowsToDelete.filter(function (v, i, a) { return a.indexOf(v) === i; });
  rowsToDelete.sort(function (a, b) { return b - a; });
  rowsToDelete.forEach(function (rowNum) {
    sheet.deleteRow(rowNum);
    Logger.log('  Deleted row ' + rowNum);
  });

  Logger.log('\nFixed ' + rowsToDelete.length + ' row(s).');
  return { fixed: rowsToDelete.length };
}

// =====================================================================
// BLANK ROWS
// =====================================================================

function admin_findAllBlankRows() {
  var sheets = ['Clients', 'MenuItems', 'Quotes', 'DeliveryNotes', 'Invoices'];
  var report = [];
  sheets.forEach(function (name) {
    var r = admin_sheetRows(name);
    if (!r || r.data.length <= 1) return;
    var blanks = [];
    for (var i = 1; i < r.data.length; i++) {
      if (admin_isBlankRow(r.data[i])) blanks.push(i + 1);
    }
    if (blanks.length > 0) report.push({ sheet: name, rows: blanks });
  });

  Logger.log('=== BLANK ROW REPORT ===\n');
  if (report.length === 0) {
    Logger.log('  ✅ No blank rows found anywhere.\n');
    return { total: 0, sheets: [] };
  }
  var totalBlanks = 0;
  report.forEach(function (x) {
    Logger.log('[' + x.sheet + '] ' + x.rows.length + ' blank row(s): ' + x.rows.join(', '));
    totalBlanks += x.rows.length;
  });
  Logger.log('\nTOTAL blank rows: ' + totalBlanks);
  return { total: totalBlanks, sheets: report };
}

function admin_removeAllBlankRows(confirmed) {
  var report = admin_findAllBlankRows();
  if (report.total === 0) {
    Logger.log('Nothing to remove.');
    return { removed: 0 };
  }
  if (confirmed !== true) {
    Logger.log('⚠️ Preview only. To remove, run: admin_removeAllBlankRows(true)');
    return { removed: 0, preview: report.total };
  }

  var ss = SpreadsheetApp.getActive();
  var removed = 0;
  report.sheets.forEach(function (x) {
    var sheet = ss.getSheetByName(x.sheet);
    if (!sheet) return;
    var rows = x.rows.slice().sort(function (a, b) { return b - a; });
    rows.forEach(function (rowNum) { sheet.deleteRow(rowNum); removed++; });
    Logger.log('  Removed ' + rows.length + ' blank row(s) from ' + x.sheet);
  });
  Logger.log('\n✅ Removed ' + removed + ' blank row(s) total.');
  return { removed: removed };
}

// =====================================================================
// MARK AS SUPERSEDED
// =====================================================================

function admin_markSuperseded(quoteNumbers, confirmed) {
  if (!quoteNumbers || quoteNumbers.length === 0) {
    Logger.log('No quotes passed.');
    return { updated: 0 };
  }

  var qr = admin_sheetRows('Quotes');
  var sheet = qr.sheet;
  var rowsToUpdate = [];
  var alreadyDone = [];

  for (var i = 1; i < qr.data.length; i++) {
    var qn = String(qr.data[i][0] || '').trim().toUpperCase();
    if (quoteNumbers.indexOf(qn) === -1) continue;
    var status = String(qr.data[i][8] || '').trim();
    if (status === 'Superseded') { alreadyDone.push(qn); continue; }
    rowsToUpdate.push({ rowNum: i + 1, qn: qn, from: status });
  }

  var seen = {};
  rowsToUpdate = rowsToUpdate.filter(function (x) {
    if (seen[x.rowNum]) return false;
    seen[x.rowNum] = true;
    return true;
  });

  Logger.log('=== MARK AS SUPERSEDED ===');
  Logger.log('Will change ' + rowsToUpdate.length + ' row(s).');
  if (alreadyDone.length > 0) Logger.log('Already Superseded: ' + alreadyDone.join(', '));

  if (!confirmed) {
    Logger.log('⚠️ Preview only. To apply, re-run with true.');
    return { updated: 0, preview: rowsToUpdate.length };
  }

  rowsToUpdate.forEach(function (x) {
    sheet.getRange(x.rowNum, 9).setValue('Superseded');
  });
  Logger.log('✅ Marked ' + rowsToUpdate.length + ' row(s) as Superseded.');
  return { updated: rowsToUpdate.length };
}

// =====================================================================
// DELETE INVOICES
// =====================================================================

function admin_deleteInvoices(invoiceNumbers, confirmed) {
  if (!invoiceNumbers || invoiceNumbers.length === 0) {
    Logger.log('No invoices passed.');
    return { deleted: 0 };
  }
  var ir = admin_sheetRows('Invoices');
  if (!ir) { Logger.log('❌ Invoices sheet not found.'); return { deleted: 0 }; }
  var sheet = ir.sheet;

  var dn = admin_sheetRows('DeliveryNotes');
  var referencedInDNs = {};
  if (dn) {
    for (var i = 1; i < dn.data.length; i++) {
      var dnInv = String(dn.data[i][2] || '').trim().toUpperCase();
      if (dnInv) {
        dnInv.split(',').forEach(function (x) { referencedInDNs[x.trim().toUpperCase()] = true; });
      }
    }
  }

  var rowsToDelete = [];
  for (var i = 1; i < ir.data.length; i++) {
    var inv = String(ir.data[i][0] || '').trim().toUpperCase();
    if (invoiceNumbers.indexOf(inv) > -1) {
      rowsToDelete.push({
        rowNum: i + 1, inv: inv, client: ir.data[i][4], status: ir.data[i][10],
        referenced: referencedInDNs[inv] || false
      });
    }
  }

  Logger.log('=== DELETE INVOICES ===');
  Logger.log('Will delete ' + rowsToDelete.length + ' row(s):');
  rowsToDelete.forEach(function (x) {
    Logger.log('  • Row ' + x.rowNum + ' | ' + x.inv + ' | client ' + x.client +
               (x.referenced ? '  ⚠️ REFERENCED IN DN' : ''));
  });

  if (!confirmed) {
    Logger.log('⚠️ Preview only. To apply, run admin_deleteInvoices([...], true)');
    return { deleted: 0, preview: rowsToDelete.length };
  }

  var seen = {};
  rowsToDelete = rowsToDelete.filter(function (x) {
    if (seen[x.rowNum]) return false;
    seen[x.rowNum] = true;
    return true;
  }).sort(function (a, b) { return b.rowNum - a.rowNum; });

  rowsToDelete.forEach(function (x) { sheet.deleteRow(x.rowNum); });
  Logger.log('✅ Deleted ' + rowsToDelete.length + ' row(s).');
  return { deleted: rowsToDelete.length };
}

// =====================================================================
// PDF CLEANUP (duplicates / orphans)
// =====================================================================

var ADMIN_CLEANUP_FOLDERS = [
  { name: 'Catering PDFs',       kind: 'quote',   sheet: 'Quotes' },
  { name: 'Invoiced Quotes',     kind: 'quote',   sheet: 'Quotes' },
  { name: 'Invoices PDFs',       kind: 'invoice', sheet: 'Invoices' },
  { name: 'Delivery Notes PDFs', kind: 'dn',      sheet: 'DeliveryNotes' }
];

var ADMIN_CLEANUP_REGEX = {
  quote:   /^Q-(\d{4})-(\d{3})\.pdf$/i,
  invoice: /^INV-(\d{4})-(\d{3})\.pdf$/i,
  dn:      /^DN-(\d{4})-(\d{3})\.pdf$/i
};

function admin_getIdSet(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var ids = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) ids[String(data[i][0]).trim().toUpperCase()] = true;
  }
  return ids;
}

function admin_fileInfo(file) {
  return {
    id: file.getId(), name: file.getName(),
    lastUpdated: file.getLastUpdated().toISOString(), url: file.getUrl()
  };
}

function admin_stripExtension(name) {
  return name.replace(/\.pdf$/i, '').toUpperCase();
}

function admin_previewDuplicatePDFs() {
  var report = { generatedAt: new Date().toISOString(), groups: [], totalDuplicatesFound: 0, totalFilesToTrash: 0 };

  for (var i = 0; i < ADMIN_CLEANUP_FOLDERS.length; i++) {
    var cfg = ADMIN_CLEANUP_FOLDERS[i];
    var iter = DriveApp.getFoldersByName(cfg.name);
    if (!iter.hasNext()) continue;
    var folder = iter.next();
    var regex = ADMIN_CLEANUP_REGEX[cfg.kind];

    var byName = {};
    var files = folder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      var nm = f.getName();
      if (!regex.test(nm)) continue;
      var key = nm.toUpperCase();
      if (!byName[key]) byName[key] = [];
      byName[key].push(f);
    }

    for (var k in byName) {
      var arr = byName[k];
      if (arr.length < 2) continue;
      arr.sort(function (a, b) { return b.getLastUpdated().getTime() - a.getLastUpdated().getTime(); });
      var keep = arr[0];
      var trash = arr.slice(1);
      report.groups.push({
        folder: cfg.name, filename: keep.getName(),
        keep: admin_fileInfo(keep), trash: trash.map(admin_fileInfo)
      });
      report.totalDuplicatesFound++;
      report.totalFilesToTrash += trash.length;
    }
  }

  Logger.log('=== DUPLICATE PDF REPORT ===');
  Logger.log('Duplicate groups:  ' + report.totalDuplicatesFound);
  Logger.log('Files to trash:    ' + report.totalFilesToTrash);
  report.groups.forEach(function (g) {
    Logger.log('[' + g.folder + '] ' + g.filename);
    g.trash.forEach(function (t) { Logger.log('  TRASH: ' + t.url); });
  });
  return report;
}

function admin_previewOrphanPDFs() {
  var report = { generatedAt: new Date().toISOString(), orphans: [], totalOrphans: 0 };
  var idCache = {};

  for (var i = 0; i < ADMIN_CLEANUP_FOLDERS.length; i++) {
    var cfg = ADMIN_CLEANUP_FOLDERS[i];
    var iter = DriveApp.getFoldersByName(cfg.name);
    if (!iter.hasNext()) continue;
    var folder = iter.next();
    var regex = ADMIN_CLEANUP_REGEX[cfg.kind];

    if (!idCache[cfg.sheet]) idCache[cfg.sheet] = admin_getIdSet(cfg.sheet);
    var validIds = idCache[cfg.sheet];

    var files = folder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      var nm = f.getName();
      if (!regex.test(nm)) continue;
      var id = admin_stripExtension(nm);
      if (!validIds[id]) {
        report.orphans.push({
          folder: cfg.name, name: nm, expectedId: id,
          fileId: f.getId(), lastUpdated: f.getLastUpdated().toISOString(), url: f.getUrl()
        });
        report.totalOrphans++;
      }
    }
  }

  Logger.log('=== ORPHAN PDF REPORT ===');
  Logger.log('Orphans found: ' + report.totalOrphans);
  report.orphans.forEach(function (o) {
    Logger.log('[' + o.folder + '] ' + o.name + ' → ' + o.url);
  });
  return report;
}

function admin_runCleanupPreview() {
  Logger.log('=== ADMIN CLEANUP PREVIEW ===');
  var dup  = admin_previewDuplicatePDFs();
  var orph = admin_previewOrphanPDFs();
  Logger.log('Duplicate files that would be trashed: ' + dup.totalFilesToTrash);
  Logger.log('Orphan files that would be trashed:    ' + orph.totalOrphans);
  return { duplicates: dup, orphans: orph };
}

function admin_trashFiles(fileIds) {
  var trashed = 0, errors = [];
  fileIds.forEach(function (id) {
    try {
      DriveApp.getFileById(id).setTrashed(true);
      trashed++;
    } catch (e) {
      errors.push(id + ': ' + e.message);
    }
  });
  return { trashed: trashed, errors: errors };
}

function admin_trashDuplicatePDFs(confirmed) {
  var preview = admin_previewDuplicatePDFs();
  if (preview.totalFilesToTrash === 0) { Logger.log('Nothing to trash.'); return { trashed: 0 }; }

  var proceed = (confirmed === true);
  if (!proceed) {
    try {
      var ui = SpreadsheetApp.getUi();
      var resp = ui.alert('⚠️ Trash Duplicate PDFs',
        'About to trash ' + preview.totalFilesToTrash + ' duplicate file(s).\nProceed?',
        ui.ButtonSet.YES_NO);
      proceed = (resp === ui.Button.YES);
    } catch (e) {
      return { trashed: 0, message: 'Not confirmed. Pass true.' };
    }
  }
  if (!proceed) return { trashed: 0, message: 'Cancelled.' };

  var idsToTrash = [];
  preview.groups.forEach(function (g) {
    g.trash.forEach(function (t) { idsToTrash.push(t.id); });
  });
  var result = admin_trashFiles(idsToTrash);
  Logger.log('Trashed ' + result.trashed + ' duplicate file(s).');
  return { trashed: result.trashed, errors: result.errors };
}

function admin_trashOrphanPDFs(confirmed) {
  var preview = admin_previewOrphanPDFs();
  if (preview.totalOrphans === 0) { Logger.log('Nothing to trash.'); return { trashed: 0 }; }

  var proceed = (confirmed === true);
  if (!proceed) {
    try {
      var ui = SpreadsheetApp.getUi();
      var resp = ui.alert('⚠️ Trash Orphan PDFs',
        'About to trash ' + preview.totalOrphans + ' orphan file(s).\nProceed?',
        ui.ButtonSet.YES_NO);
      proceed = (resp === ui.Button.YES);
    } catch (e) {
      return { trashed: 0, message: 'Not confirmed. Pass true.' };
    }
  }
  if (!proceed) return { trashed: 0, message: 'Cancelled.' };

  var idsToTrash = preview.orphans.map(function (o) { return o.fileId; });
  var result = admin_trashFiles(idsToTrash);
  Logger.log('Trashed ' + result.trashed + ' orphan file(s).');
  return { trashed: result.trashed, errors: result.errors };
}

function admin_runCleanupFull(confirmed) {
  var dup  = admin_previewDuplicatePDFs();
  var orph = admin_previewOrphanPDFs();
  var totalToTrash = dup.totalFilesToTrash + orph.totalOrphans;
  if (totalToTrash === 0) { Logger.log('Nothing to clean.'); return { duplicates: 0, orphans: 0 }; }

  var proceed = (confirmed === true);
  if (!proceed) {
    try {
      var ui = SpreadsheetApp.getUi();
      var resp = ui.alert('⚠️ Full Admin Cleanup',
        'About to trash ' + totalToTrash + ' file(s).\nProceed?', ui.ButtonSet.YES_NO);
      proceed = (resp === ui.Button.YES);
    } catch (e) {
      return { duplicates: 0, orphans: 0, message: 'Not confirmed. Pass true.' };
    }
  }
  if (!proceed) return { duplicates: 0, orphans: 0, message: 'Cancelled.' };

  var dupIds = [];
  dup.groups.forEach(function (g) { g.trash.forEach(function (t) { dupIds.push(t.id); }); });
  var dupResult = admin_trashFiles(dupIds);
  var orphIds = orph.orphans.map(function (o) { return o.fileId; });
  var orphResult = admin_trashFiles(orphIds);

  return {
    duplicates: dupResult.trashed, orphans: orphResult.trashed,
    errors: dupResult.errors.concat(orphResult.errors)
  };
}

// =====================================================================
// RESET DELIVERY NOTES
// =====================================================================

function admin_resetAllDeliveryNotes() {
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }

  if (ui) {
    var confirm = ui.alert('⚠️ Reset All Delivery Notes',
      'This deletes all DeliveryNotes rows and trashes their PDFs/signatures.\nAre you sure?',
      ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) { Logger.log('Cancelled.'); return; }
  }

  var sheet = SpreadsheetApp.getActive().getSheetByName('DeliveryNotes');
  if (sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
      Logger.log('Deleted ' + (lastRow - 1) + ' rows from DeliveryNotes.');
    }
  }
  admin_trashFolderContents('Delivery Notes PDFs');
  admin_trashFolderContents('Delivery Signatures');
  Logger.log('All delivery notes reset.');
  if (ui) ui.alert('✅ Delivery notes reset complete.');
}

function admin_trashFolderContents(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) { Logger.log('Folder not found: ' + folderName); return; }
  var folder = folders.next();
  var files = folder.getFiles();
  var count = 0;
  while (files.hasNext()) { files.next().setTrashed(true); count++; }
  Logger.log('Trashed ' + count + ' file(s) from "' + folderName + '".');
}

// =====================================================================
// DASHBOARD STATS
// =====================================================================

function getDashboardStats() {
  var ss = SpreadsheetApp.getActive();
  var stats = {
    quotes: { draft: 0, invoiced: 0, superseded: 0, total: 0 },
    invoices: { unpaid: 0, paid: 0, total: 0 },
    deliveryNotes: { draft: 0, sent: 0, signed: 0, total: 0 },
    clients: 0
  };

  var qSheet = ss.getSheetByName('Quotes');
  if (qSheet) {
    var qData = qSheet.getDataRange().getValues();
    var seenQ = {};
    for (var i = 1; i < qData.length; i++) {
      var num = qData[i][0];
      if (num && !seenQ[num]) {
        seenQ[num] = true;
        var status = String(qData[i][8] || 'Draft');
        if (status === 'Invoiced') stats.quotes.invoiced++;
        else if (status === 'Superseded') stats.quotes.superseded++;
        else stats.quotes.draft++;
      }
    }
    stats.quotes.total = stats.quotes.draft + stats.quotes.invoiced + stats.quotes.superseded;
  }

  var iSheet = ss.getSheetByName('Invoices');
  if (iSheet) {
    var iData = iSheet.getDataRange().getValues();
    var seenI = {};
    for (var i = 1; i < iData.length; i++) {
      var num = iData[i][0];
      if (num && !seenI[num]) {
        seenI[num] = true;
        var status = String(iData[i][10] || 'Unpaid');
        if (status === 'Paid') stats.invoices.paid++;
        else stats.invoices.unpaid++;
      }
    }
    stats.invoices.total = stats.invoices.paid + stats.invoices.unpaid;
  }

  var dSheet = ss.getSheetByName('DeliveryNotes');
  if (dSheet) {
    var dData = dSheet.getDataRange().getValues();
    for (var i = 1; i < dData.length; i++) {
      if (!dData[i][0]) continue;
      var status = String(dData[i][5] || 'Draft');
      if (status === 'Signed') stats.deliveryNotes.signed++;
      else if (status === 'Sent') stats.deliveryNotes.sent++;
      else stats.deliveryNotes.draft++;
    }
    stats.deliveryNotes.total = stats.deliveryNotes.draft + stats.deliveryNotes.sent + stats.deliveryNotes.signed;
  }

  var cSheet = ss.getSheetByName('Clients');
  if (cSheet) {
    var cData = cSheet.getDataRange().getValues();
    for (var i = 1; i < cData.length; i++) {
      if (cData[i][0] && cData[i][1]) stats.clients++;
    }
  }
  return stats;
}

// =====================================================================
// EDITOR WRAPPERS (live, keep these)
// =====================================================================

function RUN_healthCheck() { return admin_runHealthCheck(); }
function RUN_previewDNDuplicates() { return admin_previewDuplicateDNRows(); }
function RUN_fixDNDuplicates() { return admin_fixDuplicateDNRows(true); }
function RUN_previewAllBlankRows() { return admin_findAllBlankRows(); }
function RUN_removeAllBlankRows_PREVIEW() { return admin_removeAllBlankRows(false); }
function RUN_removeAllBlankRows_APPLY() { return admin_removeAllBlankRows(true); }
function RUN_trashDuplicates() { return admin_trashDuplicatePDFs(true); }
function RUN_trashOrphans() { return admin_trashOrphanPDFs(true); }
function RUN_fullCleanup() { return admin_runCleanupFull(true); }

function testServerFunctions() {
  try { Logger.log('✅ Clients: ' + getClients().length); } catch (e) { Logger.log('❌ ' + e.message); }
  try { Logger.log('✅ Quotes: ' + getQuoteList().length); } catch (e) { Logger.log('❌ ' + e.message); }
  try { Logger.log('✅ Menu items: ' + getAllMenuItems().length); } catch (e) { Logger.log('❌ ' + e.message); }
}
