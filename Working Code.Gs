/* ==========================================================================
 * Existing form-photo logic (unchanged)
 * ========================================================================== */

/**
 * Run this ONCE from the Apps Script editor (select setupFormTrigger,
 * click Run, authorize when prompted) to wire this up to "On form submit".
 * Safe to re-run any time — it won't create duplicate triggers.
 */
function setupFormTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onMemberFormSubmit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onMemberFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  return 'Trigger installed.';
}

function onMemberFormSubmit(e) {
  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });

  var photoColIndex = headers.indexOf('Member Photo') + 1;
  var pathColIndex = headers.indexOf('PhotoPath') + 1;
  if (photoColIndex === 0 || pathColIndex === 0) return; // headers missing, nothing to do

  var rawValue = sheet.getRange(row, photoColIndex).getValue();
  if (!rawValue) return;

  // Forms can return multiple comma-separated URLs if multi-file upload is allowed.
  var urls = String(rawValue).split(',').map(function (u) { return u.trim(); }).filter(Boolean);
  var paths = urls.map(deriveImagePath).filter(Boolean);

  sheet.getRange(row, pathColIndex).setValue(paths.join(', '));
}

/** Turns a Drive "open?id=..." URL into "/Images/<actual filename>.jpg" without moving the file. */
function deriveImagePath(url) {
  var match = url.match(/[-\w]{25,}/); // Drive file IDs are long alphanumeric strings
  if (!match) return '';
  var fileId = match[0];
  try {
    var file = DriveApp.getFileById(fileId);
    return '/Member Registration Form (File responses)/Member Photo (File responses)/' + file.getName();
  } catch (err) {
    return ''; // file inaccessible or ID couldn't be resolved
  }
}

/* ==========================================================================
 * InvoicePath -> InvoiceUrl logic (for AppSheet-inserted rows)
 * ========================================================================== */

/**
 * Run this ONCE from the Apps Script editor to install an "On change" trigger.
 *
 * IMPORTANT: AppSheet inserts/edits rows via the Sheets API, which does NOT
 * fire onEdit (simple or installable). Only an installable onChange trigger
 * fires for API-driven edits — so this uses onChange instead of onEdit.
 *
 * Safe to re-run any time — it won't create duplicate triggers.
 */
function setupInvoiceUrlTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onInvoiceSheetChange') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onInvoiceSheetChange')
    .forSpreadsheet(ss)
    .onChange()
    .create();
  return 'Invoice URL trigger installed.';
}

/**
 * onChange doesn't tell us which sheet/row changed, so we sweep every sheet
 * in the spreadsheet and fill in any blank InvoiceUrl cells that have a
 * matching InvoicePath. Cheap enough for typical contribution-sheet sizes.
 */
function onInvoiceSheetChange(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function (sheet) {
    fillMissingInvoiceUrls(sheet);
  });
}

/**
 * Scans a sheet for rows where InvoicePath is set but InvoiceUrl is blank,
 * resolves the file in Drive, and writes its view URL.
 * Can also be run manually, or wired to a time-based trigger as a safety net.
 */
function fillMissingInvoiceUrls(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSheet();
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h).trim(); });

  var pathCol = headers.indexOf('InvoicePath') + 1;
  var urlCol = headers.indexOf('InvoiceUrl') + 1;
  if (pathCol === 0 || urlCol === 0) return; // this sheet doesn't have these columns

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var numRows = lastRow - 1;
  var paths = sheet.getRange(2, pathCol, numRows, 1).getValues();
  var urls = sheet.getRange(2, urlCol, numRows, 1).getValues();

  var folderCache = {}; // avoid re-walking the same folder path repeatedly
  var updates = [];
  var hasUpdates = false;

  for (var i = 0; i < numRows; i++) {
    var path = String(paths[i][0] || '').trim();
    var existingUrl = String(urls[i][0] || '').trim();

    if (path && !existingUrl) {
      var fileUrl = getFileUrlFromPath(path, folderCache);
      updates.push([fileUrl || '']);
      if (fileUrl) hasUpdates = true;
    } else {
      updates.push([existingUrl]);
    }
  }

  if (hasUpdates) {
    sheet.getRange(2, urlCol, numRows, 1).setValues(updates);
  }
}

/**
 * InvoicePath only stores the path relative to the actual parent folder
 * (e.g. "/Contribution Invoices/Some File_abc123.pdf"), but the real
 * location in Drive is:
 *   My Drive > Keishamthong Apunba Chaokhat Marup > Contribution Invoices
 * So we prepend that base path before walking folders. Update this array
 * if the folder is ever renamed or moved.
 */
var INVOICE_BASE_FOLDER_PATH = ['Keishamthong Apunba Chaokhat Marup'];

/**
 * Resolves a path like "/Contribution Invoices/Some File_abc123.pdf"
 * to that specific file's shareable view URL (not the folder), e.g.
 * "https://drive.google.com/file/d/<fileId>/view?usp=sharing"
 *
 * Assumes the folder lives under My Drive. If it's in a Shared Drive
 * instead, swap DriveApp.getRootFolder() in getFolderByPath() for the
 * Shared Drive's root (e.g. DriveApp.getFolderById('SHARED_DRIVE_ROOT_ID')).
 */
function getFileUrlFromPath(path, folderCache) {
  folderCache = folderCache || {};
  var relativeParts = String(path).split('/').filter(Boolean); // drop leading/trailing slashes
  if (relativeParts.length === 0) return '';

  var fileName = relativeParts.pop();
  var folderParts = INVOICE_BASE_FOLDER_PATH.concat(relativeParts);
  var folder = getFolderByPath(folderParts, folderCache);
  if (!folder) return '';

  var files = folder.getFilesByName(fileName);
  if (!files.hasNext()) return '';

  var file = files.next();
  // File-level "view" link — opens the PDF itself, not the containing folder.
  return 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=sharing';
}

/** Walks a list of folder names starting from My Drive root, with caching. */
function getFolderByPath(folderNames, cache) {
  var key = folderNames.join('/');
  if (cache[key]) return cache[key];

  var folder = DriveApp.getRootFolder();
  for (var i = 0; i < folderNames.length; i++) {
    var subFolders = folder.getFoldersByName(folderNames[i]);
    if (!subFolders.hasNext()) return null; // folder segment not found
    folder = subFolders.next();
  }
  cache[key] = folder;
  return folder;
}

/* ==========================================================================
 * PDF Report Generation (AppScriptReports table)
 * ========================================================================== */

var REPORTS_SHEET_NAME = 'AppScriptReports';
var MONTHLY_REPORT_SOURCE_SHEET = 'MonthlyContributionReport';
var YEARLY_REPORT_SOURCE_SHEET = 'YearlyContributionReport';
var ONE_TIME_REPORT_SOURCE_SHEET = 'OneTimeContributionReport';
var CONTRIBUTION_CATEGORY_SHEET_NAME = 'Contribution Category';
var ORG_NAME = 'Keishamthong Apunba Chaokhat Marup (KACM)';
var ORG_ADDRESS = 'Keishamthong Ahanthem Leikai, Imphal-West, Manipur - 795001';
var MONTH_HEADERS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// A4 Landscape (Monthly report), in points (1mm = 2.83465pt). Long side = width when landscape.
var PAGE_WIDTH = 841.89;
var PAGE_HEIGHT = 595.28;

// A4 Portrait (Yearly / One Time reports) — same A4 dimensions, just swapped.
var PAGE_WIDTH_PORTRAIT = 595.28;
var PAGE_HEIGHT_PORTRAIT = 841.89;

var PAGE_MARGIN = 28.35; // ~1cm
var MEMBERS_PER_PAGE = 18;         // Monthly report
var YEARLY_MEMBERS_PER_PAGE = 25;  // Yearly report
var ONE_TIME_ROWS_PER_PAGE = 25;   // One Time report
var MEMBER_FORM_SHEET_NAME = 'Form Responses 1';
var MEMBER_ROWS_PER_PAGE = 35;     // Member list pages
var PHOTO_GRID_COLUMNS = 4;        // Member photo collage — configurable
var PHOTO_GRID_ROWS = 8;           // Member photo collage — configurable
var PHOTO_GRID_GAP = 6;            // points, gap between photo cells
var STATEMENT_SOURCE_SHEET = 'AccountStatement';
var STATEMENT_ROWS_FIRST_PAGE = 35;   // page 1 has less room (summary table takes space)
var STATEMENT_ROWS_OTHER_PAGES = 45;  // subsequent pages have more room (no summary table)

/**
 * Run this ONCE from the Apps Script editor to install the trigger that
 * watches AppScriptReports for new/edited rows.
 *
 * IMPORTANT: If AppSheet inserts these rows via the Sheets API (same as the
 * Reports table before it), onEdit will NOT fire — only an installable
 * onChange trigger catches API-driven edits. That's why this uses onChange,
 * same reasoning as the InvoiceUrl trigger above.
 *
 * Safe to re-run any time — it won't create duplicate triggers.
 */
function setupReportGenerationTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onReportsSheetChange') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onReportsSheetChange')
    .forSpreadsheet(ss)
    .onChange()
    .create();
  return 'Report generation trigger installed.';
}

/**
 * Optional safety net: wire this function to a time-driven trigger
 * (Triggers UI -> Add Trigger -> processPendingReportsSafetyNet -> Time-driven
 * -> every 5 minutes) in case onChange ever misses an edit.
 */
function processPendingReportsSafetyNet() {
  processPendingReports();
}

function onReportsSheetChange(e) {
  processPendingReports();
}

/** Scans AppScriptReports for IsReady=FALSE rows and generates whatever it currently knows how to. */
function processPendingReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(REPORTS_SHEET_NAME);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });

  var col = {};
  headers.forEach(function (h, i) { col[h] = i + 1; });

  var required = ['ReportID', 'Report Type', 'Year', 'Collection Frequency', 'Generated DateTime', 'ReportFilePath', 'IsReady'];
  for (var r = 0; r < required.length; r++) {
    if (!col[required[r]]) {
      throw new Error('AppScriptReports is missing expected column: ' + required[r]);
    }
  }

  var numRows = lastRow - 1;
  var data = sheet.getRange(2, 1, numRows, sheet.getLastColumn()).getValues();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var isReady = row[col['IsReady'] - 1];
    if (isReady === true || String(isReady).toUpperCase() === 'TRUE') continue; // already done

    var reportId = row[col['ReportID'] - 1];
    var reportType = row[col['Report Type'] - 1];
    var year = row[col['Year'] - 1];
    var frequency = String(row[col['Collection Frequency'] - 1] || '').trim();
    var generatedDateTime = row[col['Generated DateTime'] - 1];
    var filePath = row[col['ReportFilePath'] - 1];

    if (!reportId || !filePath) continue; // incomplete row, skip for now

    try {
      var success = false;

      if (reportType === 'Statement') {
        // "Statement" isn't a real Category and isn't tied to a Collection
        // Frequency the way category-based reports are, so this is checked
        // first, ahead of the frequency-based dispatch below.
        success = generateStatementReport(year, filePath);
      } else if (reportType === 'Member') {
        // Same reasoning as "Statement" — not category/frequency-based.
        success = generateMemberReport(filePath);
      } else if (frequency === 'Monthly') {
        success = generateMonthlyReport(reportType, year, generatedDateTime, filePath);
      } else if (frequency === 'Yearly') {
        success = generateYearlyReport(reportType, year, filePath);
      } else if (frequency === 'One Time') {
        // Only handle this here if reportType is a real Category (registered
        // in Contribution Category with Collection Frequency = "One Time").
        // Other "One Time" rows are reserved for ad-hoc report types (e.g. a
        // future Transaction Report) that aren't category-based and will get
        // their own dedicated handler later — those are deliberately left
        // untouched (IsReady stays FALSE) rather than guessed at here.
        if (isRegisteredOneTimeCategory(reportType)) {
          success = generateOneTimeReport(reportType, filePath);
        } else {
          Logger.log('Skipping AppScriptReports row ' + reportId + ' ("' + reportType +
            '"): One Time but not a registered category — assuming this is an ad-hoc report type to be handled separately.');
        }
      }

      if (success) {
        sheet.getRange(2 + i, col['IsReady']).setValue(true);
      }
    } catch (err) {
      Logger.log('Failed to generate report ' + reportId + ': ' + err);
      // Row is left with IsReady = FALSE so it gets retried on the next run.
    }
  }
}

/**
 * Converts "FY-26" -> "2026" to match the plain-year values used in
 * MonthlyContributionReport / YearlyContributionReport. If Year already
 * looks like a plain year (e.g. "2026"), it's returned unchanged.
 * ADJUST THIS if your fiscal-year naming convention is different
 * (e.g. FY-26 meaning Apr 2025-Mar 2026 instead of calendar 2026).
 */
function fiscalYearToCalendarYear(fyValue) {
  var match = String(fyValue).match(/(\d{2})$/);
  if (!match) return String(fyValue).trim();
  return '20' + match[1];
}

/** True if `category` exists in Contribution Category with Collection Frequency = "One Time". */
function isRegisteredOneTimeCategory(category) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONTRIBUTION_CATEGORY_SHEET_NAME);
  if (!sheet) return false;

  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  if (idx['Category Name'] === undefined || idx['Collection Frequency'] === undefined) return false;

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[idx['Category Name']]).trim() === String(category).trim() &&
        String(row[idx['Collection Frequency']]).trim() === 'One Time') {
      return true;
    }
  }
  return false;
}

/* -------------------------------------------------------------------------
 * Monthly report
 * ---------------------------------------------------------------------- */

function generateMonthlyReport(category, yearRaw, generatedDateTime, filePath) {
  var year = fiscalYearToCalendarYear(yearRaw);
  var rows = getMonthlyReportRows(category, year);

  if (rows.length === 0) {
    Logger.log('No MonthlyContributionReport rows found for Category=' + category + ', Year=' + year + ' (generating report anyway).');
  }

  var doc = buildMonthlyReportDoc(category, year, generatedDateTime, rows);
  var docFile = DriveApp.getFileById(doc.getId());
  var pdfBlob = docFile.getAs('application/pdf');
  pdfBlob.setName(getFileNameFromPath(filePath));

  saveBlobToPath(pdfBlob, filePath);

  docFile.setTrashed(true); // clean up the temp Google Doc, keep only the PDF

  return true;
}

/** Pulls Member + Jan..Dec + Total for the given Category/Year from MonthlyContributionReport. */
function getMonthlyReportRows(category, year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MONTHLY_REPORT_SOURCE_SHEET);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  var results = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[idx['Category']]) === String(category) && String(row[idx['Year']]) === String(year)) {
      results.push({
        member: row[idx['Member']],
        months: MONTH_HEADERS.map(function (m) { return Number(row[idx[m]]) || 0; }),
        total: Number(row[idx['Total']]) || 0
      });
    }
  }
  return results;
}

/**
 * Builds the multi-page A4-landscape Monthly report:
 *  - Org header + info block repeat on every page
 *  - Member table (with SlNo/Member/Mobile, header row) repeats on every page, 18 members per page
 *  - Grand-total row only appears on the final page
 *  - "Page X of Y" centered under each page's table
 *  - Footer disclaimer appears once, after the last page's content
 */
function buildMonthlyReportDoc(category, year, generatedDateTime, rows) {
  var doc = DocumentApp.create('TEMP_Monthly_' + category + '_' + year + '_' + new Date().getTime());
  var body = doc.getBody();

  body.setPageWidth(PAGE_WIDTH).setPageHeight(PAGE_HEIGHT);
  body.setMarginTop(PAGE_MARGIN).setMarginBottom(PAGE_MARGIN)
      .setMarginLeft(PAGE_MARGIN).setMarginRight(PAGE_MARGIN);

  rows.sort(function (a, b) { return String(a.member).localeCompare(String(b.member)); });

  var totalPages = Math.max(1, Math.ceil(rows.length / MEMBERS_PER_PAGE));

  var totals = MONTH_HEADERS.map(function () { return 0; });
  var grandTotal = 0;
  rows.forEach(function (r) {
    r.months.forEach(function (v, i) { totals[i] += v; });
    grandTotal += r.total;
  });

  for (var p = 0; p < totalPages; p++) {
    if (p > 0) body.appendPageBreak();

    appendReportHeaderBlock(body, category, year, generatedDateTime, 'Monthly', PAGE_WIDTH - (PAGE_MARGIN * 2));

    var chunk = rows.slice(p * MEMBERS_PER_PAGE, (p + 1) * MEMBERS_PER_PAGE);
    var isLastPage = (p === totalPages - 1);

    appendMonthlyMemberTable(body, chunk, isLastPage ? totals : null, isLastPage ? grandTotal : null, p * MEMBERS_PER_PAGE);

    body.appendParagraph('');
    var pageLabel = body.appendParagraph('Page ' + (p + 1) + ' of ' + totalPages);
    pageLabel.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    pageLabel.editAsText().setFontSize(9).setForegroundColor('#888888');
  }

  var footer = body.appendParagraph('This is a system-generated report and does not require a physical or digital signature.');
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footer.editAsText().setItalic(true).setFontSize(9).setForegroundColor('#888888');

  doc.saveAndClose();
  return doc;
}

/**
 * The repeating top section shared by all report types: org name/address + info table.
 * frequencyLabel is shown in the "Collection Frequency" cell ('Monthly'/'Yearly'/'One Time').
 * year is optional — pass '' (empty string) for reports with no year (e.g. One Time);
 * the Report Name then omits the "FOR <year>" suffix.
 */
function appendReportHeaderBlock(body, category, year, generatedDateTime, frequencyLabel, contentWidth) {
  var title = body.appendParagraph(ORG_NAME);
  title.setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.LEFT).setFontFamily("Special Gothic Condensed One");

  var address = body.appendParagraph(ORG_ADDRESS);
  address.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  address.editAsText().setFontSize(10).setForegroundColor('#555555').setFontFamily("Arimo");

  body.appendParagraph('');

  var reportName = String(category).toUpperCase() + (year ? (' REPORT FOR ' + year) : ' REPORT');
  var genDateStr = Utilities.formatDate(new Date(generatedDateTime), Session.getScriptTimeZone(), "d MMMM yyyy hh:mm a");

  var infoTable = body.appendTable([
    ['Report Name', reportName, 'Collection Frequency', frequencyLabel],
    ['Category Name', category, 'Generated DateTime', genDateStr]
  ]);
  styleInfoTable(infoTable);
  setInfoTableColumnWidths(infoTable, contentWidth);
  // No trailing blank paragraph here — the table is appended right after
  // this, so it sits close to the info table instead of floating down.
}

/**
 * Col1 (labels) = 150, Col3 (labels) = 150, Col2/Col4 (values) split the
 * remaining width evenly: (contentWidth - col1 - col3) / 2 each.
 */
function setInfoTableColumnWidths(table, contentWidth) {
  var col1 = 100;
  var col3 = 130;
  var col4 = 130;
  var col2 = (contentWidth - col1 - col3 - col4);

  table.setColumnWidth(0, col1);
  table.setColumnWidth(1, col2);
  table.setColumnWidth(2, col3);
  table.setColumnWidth(3, col4);
}

/** Splits "AHANTHEM JODHA: 9089680302" into {name: "AHANTHEM JODHA", mobile: "9089680302"}. */
function splitMemberAndMobile(memberString) {
  var parts = String(memberString).split(':');
  var mobile = parts.length > 1 ? parts.pop().trim() : '';
  var name = parts.join(':').trim();
  return { name: name, mobile: mobile };
}

/**
 * Builds one page's member table: SlNo, Member, Mobile, Jan..Dec, Total.
 * totals/grandTotal are null on all but the last page (no grand-total row shown there).
 * startIndex is the running SlNo offset so numbering continues correctly across pages.
 */
function appendMonthlyMemberTable(body, chunkRows, totals, grandTotal, startIndex) {
  var tableData = [['SlNo', 'Member', 'Mobile'].concat(MONTH_HEADERS).concat(['Total'])];

  chunkRows.forEach(function (r, i) {
    var slNo = startIndex + i + 1;
    var parsed = splitMemberAndMobile(r.member);
    var line = [String(slNo), parsed.name.toUpperCase(), parsed.mobile].concat(r.months.map(String)).concat([String(r.total)]);
    tableData.push(line);
  });

  var hasTotalRow = !!totals;
  if (hasTotalRow) {
    tableData.push(['', 'TOTAL', ''].concat(totals.map(String)).concat([String(grandTotal)]));
  }

  var table = body.appendTable(tableData);
  styleMonthlyTable(table, hasTotalRow);
  setMonthlyTableColumnWidths(table);
  return table;
}

/** SlNo fixed narrow, Member ~25% of table width, Mobile fixed for 10 digits, remaining split evenly across Jan-Dec + Total. */
function setMonthlyTableColumnWidths(table) {
  var contentWidth = PAGE_WIDTH - (PAGE_MARGIN * 2);
  var slNoWidth = 32;
  var mobileWidth = 60; // fits a 10-digit phone number comfortably at 9pt
  var memberWidth = contentWidth * 0.25;
  var otherColsCount = MONTH_HEADERS.length + 1; // 12 months + Total
  var otherColWidth = (contentWidth - slNoWidth - memberWidth - mobileWidth) / otherColsCount;

  table.setColumnWidth(0, slNoWidth);
  table.setColumnWidth(1, memberWidth);
  table.setColumnWidth(2, mobileWidth);
  for (var c = 3; c < 3 + otherColsCount; c++) {
    table.setColumnWidth(c, otherColWidth);
  }
}

function styleMonthlyTable(table, hasTotalRow) {
  var headerRow = table.getRow(0);
  for (var c = 0; c < headerRow.getNumCells(); c++) {
    var cell = headerRow.getCell(c);
    cell.setBackgroundColor('#f1f1f1');
    cell.editAsText().setBold(true).setFontSize(9);
    if (c !== 1) setCellAlignment(cell, DocumentApp.HorizontalAlignment.CENTER); // everything except Member header
  }

  if (hasTotalRow) {
    var totalRow = table.getRow(table.getNumRows() - 1);
    for (var c2 = 0; c2 < totalRow.getNumCells(); c2++) {
      var totalCell = totalRow.getCell(c2);
      totalCell.setBackgroundColor('#f1f1f1'); // same shading as the header row
      totalCell.editAsText().setBold(true).setFontSize(9);
      if (c2 !== 1) setCellAlignment(totalCell, DocumentApp.HorizontalAlignment.CENTER);
    }
  }

  for (var r = 0; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    for (var c3 = 0; c3 < row.getNumCells(); c3++) {
      var dataCell = row.getCell(c3);
      dataCell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(3).setPaddingRight(3);
      if (r > 0 && !(hasTotalRow && r === table.getNumRows() - 1)) {
        dataCell.editAsText().setFontSize(9);
        if (c3 !== 1) setCellAlignment(dataCell, DocumentApp.HorizontalAlignment.CENTER); // center SlNo/Mobile/numbers, left-align Member
      }
    }
  }
}

/* -------------------------------------------------------------------------
 * Yearly report
 * ---------------------------------------------------------------------- */

function generateYearlyReport(category, yearRaw, filePath) {
  var year = fiscalYearToCalendarYear(yearRaw);
  var rows = getYearlyReportRows(category, year);

  if (rows.length === 0) {
    Logger.log('No YearlyContributionReport rows found for Category=' + category + ', Year=' + year + ' (generating report anyway).');
  }

  // Per spec, "Generated DateTime" for Yearly reports is the actual moment of
  // generation, not the AppScriptReports row's timestamp.
  var generatedNow = new Date();

  var doc = buildYearlyReportDoc(category, year, generatedNow, rows);
  var docFile = DriveApp.getFileById(doc.getId());
  var pdfBlob = docFile.getAs('application/pdf');
  pdfBlob.setName(getFileNameFromPath(filePath));

  saveBlobToPath(pdfBlob, filePath);

  docFile.setTrashed(true); // clean up the temp Google Doc, keep only the PDF

  return true;
}

/** Pulls SlNo/Member/Year/Transaction Date/Amount/TransactionBy for the given Category/Year. */
function getYearlyReportRows(category, year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(YEARLY_REPORT_SOURCE_SHEET);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  var results = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[idx['Category']]) === String(category) && String(row[idx['Year']]) === String(year)) {
      results.push({
        member: row[idx['Member']],
        year: row[idx['Year']],
        transactionDate: row[idx['Transaction Date']],
        amount: Number(row[idx['Amount']]) || 0,
        transactionBy: row[idx['TransactionBy']]
      });
    }
  }
  return results;
}

/**
 * Builds the multi-page A4-portrait Yearly report:
 *  - Org header + info block repeat on every page
 *  - Transaction table (with SlNo, header row) repeats on every page, 25 rows per page
 *  - Grand-total row (SUM of Amount) only appears on the final page
 *  - "Page X of Y" centered under each page's table
 *  - Footer disclaimer appears once, after the last page's content
 */
function buildYearlyReportDoc(category, year, generatedNow, rows) {
  var doc = DocumentApp.create('TEMP_Yearly_' + category + '_' + year + '_' + new Date().getTime());
  var body = doc.getBody();

  body.setPageWidth(PAGE_WIDTH_PORTRAIT).setPageHeight(PAGE_HEIGHT_PORTRAIT);
  body.setMarginTop(PAGE_MARGIN).setMarginBottom(PAGE_MARGIN)
      .setMarginLeft(PAGE_MARGIN).setMarginRight(PAGE_MARGIN);

  rows.sort(function (a, b) { return String(a.member).localeCompare(String(b.member)); });

  var totalPages = Math.max(1, Math.ceil(rows.length / YEARLY_MEMBERS_PER_PAGE));

  var grandTotal = 0;
  rows.forEach(function (r) { grandTotal += r.amount; });

  for (var p = 0; p < totalPages; p++) {
    if (p > 0) body.appendPageBreak();

    appendReportHeaderBlock(body, category, year, generatedNow, 'Yearly', PAGE_WIDTH_PORTRAIT - (PAGE_MARGIN * 2));

    var chunk = rows.slice(p * YEARLY_MEMBERS_PER_PAGE, (p + 1) * YEARLY_MEMBERS_PER_PAGE);
    var isLastPage = (p === totalPages - 1);

    appendYearlyTransactionTable(body, chunk, isLastPage ? grandTotal : null, p * YEARLY_MEMBERS_PER_PAGE);

    body.appendParagraph('');
    var pageLabel = body.appendParagraph('Page ' + (p + 1) + ' of ' + totalPages);
    pageLabel.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    pageLabel.editAsText().setFontSize(9).setForegroundColor('#888888');
  }

  var footer = body.appendParagraph('This is a system-generated report and does not require a physical or digital signature.');
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footer.editAsText().setItalic(true).setFontSize(9).setForegroundColor('#888888');

  doc.saveAndClose();
  return doc;
}

/**
 * Builds one page's transaction table: SlNo, Member, Mobile, Year, Transaction Date, Amount, TransactionBy.
 * grandTotal is null on all but the last page (no TOTAL row shown there).
 * startIndex is the running SlNo offset so numbering continues correctly across pages.
 */
function appendYearlyTransactionTable(body, chunkRows, grandTotal, startIndex) {
  var tableData = [['SlNo', 'Member', 'Mobile', 'Year', 'Transaction Date', 'Amount', 'TransactionBy']];

  chunkRows.forEach(function (r, i) {
    var slNo = startIndex + i + 1;
    var parsed = splitMemberAndMobile(r.member);
    var line = [
      String(slNo),
      parsed.name.toUpperCase(),
      parsed.mobile,
      String(r.year || ''),
      formatYearlyDate(r.transactionDate),
      formatIndianNumber(r.amount),
      r.transactionBy ? String(r.transactionBy) : ''
    ];
    tableData.push(line);
  });

  var hasTotalRow = (grandTotal !== null && grandTotal !== undefined);
  if (hasTotalRow) {
    tableData.push(['', 'TOTAL', '', '', '', formatIndianNumber(grandTotal), '']);
  }

  var table = body.appendTable(tableData);
  styleYearlyTable(table, hasTotalRow);
  setYearlyTableColumnWidths(table);
  return table;
}

/** Formats a Transaction Date cell value (Date object or date-like string) as "18 Jul 2026". */
function formatYearlyDate(value) {
  if (!value) return '';
  var d = (value instanceof Date) ? value : new Date(value);
  if (isNaN(d.getTime())) return String(value);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'd MMM yyyy');
}

/**
 * Formats a number with Indian-style thousands separators, e.g.
 * 45450 -> "45,450", 120000 -> "1,20,000", 1000000 -> "10,00,000".
 */
function formatIndianNumber(num) {
  num = Math.round(Number(num) || 0);
  var isNegative = num < 0;
  num = Math.abs(num);
  var numStr = String(num);

  if (numStr.length <= 3) {
    return (isNegative ? '-' : '') + numStr;
  }

  var lastThree = numStr.substring(numStr.length - 3);
  var otherNumbers = numStr.substring(0, numStr.length - 3);
  otherNumbers = otherNumbers.replace(/\B(?=(\d{2})+(?!\d))/g, ',');

  return (isNegative ? '-' : '') + otherNumbers + ',' + lastThree;
}

/** SlNo/Mobile/Year/Transaction Date/TransactionBy centered, Amount right-aligned; Member left-aligned. */
function setYearlyTableColumnWidths(table) {
  var contentWidth = PAGE_WIDTH_PORTRAIT - (PAGE_MARGIN * 2);
  var slNoWidth = 35;
  var mobileWidth = 60; // fits a 10-digit phone number comfortably at 9pt
  var yearWidth = 40;
  var dateWidth = 90;
  var amountWidth = 60;
  var memberWidth = 180;
  var transactionByWidth = contentWidth - (slNoWidth + mobileWidth + yearWidth + dateWidth + amountWidth + memberWidth);

  table.setColumnWidth(0, slNoWidth);
  table.setColumnWidth(1, memberWidth);
  table.setColumnWidth(2, mobileWidth);
  table.setColumnWidth(3, yearWidth);
  table.setColumnWidth(4, dateWidth);
  table.setColumnWidth(5, amountWidth);
  table.setColumnWidth(6, transactionByWidth);
}

function styleYearlyTable(table, hasTotalRow) {
  var CENTERED_COLUMNS = [0, 2, 3, 4, 6]; // SlNo, Mobile, Year, Transaction Date, TransactionBy
  var RIGHT_ALIGNED_COLUMNS = [5]; // Amount

  var headerRow = table.getRow(0);
  for (var c = 0; c < headerRow.getNumCells(); c++) {
    var cell = headerRow.getCell(c);
    cell.setBackgroundColor('#f1f1f1'); // same shading as the Monthly report's header row
    cell.editAsText().setBold(true).setFontSize(9);
    if (CENTERED_COLUMNS.indexOf(c) !== -1) setCellAlignment(cell, DocumentApp.HorizontalAlignment.CENTER);
    if (RIGHT_ALIGNED_COLUMNS.indexOf(c) !== -1) setCellAlignment(cell, DocumentApp.HorizontalAlignment.RIGHT);
  }

  if (hasTotalRow) {
    var totalRow = table.getRow(table.getNumRows() - 1);
    for (var c2 = 0; c2 < totalRow.getNumCells(); c2++) {
      var totalCell = totalRow.getCell(c2);
      totalCell.setBackgroundColor('#f1f1f1'); // same shading as the header row
      totalCell.editAsText().setBold(true).setFontSize(9);
      if (CENTERED_COLUMNS.indexOf(c2) !== -1) setCellAlignment(totalCell, DocumentApp.HorizontalAlignment.CENTER);
      if (RIGHT_ALIGNED_COLUMNS.indexOf(c2) !== -1) setCellAlignment(totalCell, DocumentApp.HorizontalAlignment.RIGHT);
    }
  }

  for (var r = 1; r < table.getNumRows(); r++) {
    if (hasTotalRow && r === table.getNumRows() - 1) continue; // already styled above
    var row = table.getRow(r);
    for (var c3 = 0; c3 < row.getNumCells(); c3++) {
      var dataCell = row.getCell(c3);
      dataCell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(3).setPaddingRight(3);
      dataCell.editAsText().setFontSize(9);
      if (CENTERED_COLUMNS.indexOf(c3) !== -1) setCellAlignment(dataCell, DocumentApp.HorizontalAlignment.CENTER);
      if (RIGHT_ALIGNED_COLUMNS.indexOf(c3) !== -1) setCellAlignment(dataCell, DocumentApp.HorizontalAlignment.RIGHT);
    }
  }
}

/* -------------------------------------------------------------------------
 * One Time report
 * (Same layout/columns as Yearly, minus the Year column. Reads from the
 * OneTimeContributionReport sheet — built the same way as YearlyContributionReport,
 * via its own OneTimeViewHelper sheet — rather than from Contributions directly.)
 * ---------------------------------------------------------------------- */

function generateOneTimeReport(category, filePath) {
  var rows = getOneTimeReportRows(category);

  if (rows.length === 0) {
    Logger.log('No OneTimeContributionReport rows found for Category=' + category + ' (generating report anyway).');
  }

  // Same convention as Yearly: "Generated DateTime" is the actual moment of
  // generation, not any timestamp stored elsewhere.
  var generatedNow = new Date();

  var doc = buildOneTimeReportDoc(category, generatedNow, rows);
  var docFile = DriveApp.getFileById(doc.getId());
  var pdfBlob = docFile.getAs('application/pdf');
  pdfBlob.setName(getFileNameFromPath(filePath));

  saveBlobToPath(pdfBlob, filePath);

  docFile.setTrashed(true); // clean up the temp Google Doc, keep only the PDF

  return true;
}

/** Pulls SlNo/Member/Transaction Date/Amount/TransactionBy for the given Category from OneTimeContributionReport. */
function getOneTimeReportRows(category) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ONE_TIME_REPORT_SOURCE_SHEET);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  var results = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[idx['Category']]) === String(category)) {
      results.push({
        member: row[idx['Member']],
        transactionDate: row[idx['Transaction Date']],
        amount: Number(row[idx['Amount']]) || 0,
        transactionBy: row[idx['TransactionBy']]
      });
    }
  }
  return results;
}

/**
 * Builds the multi-page A4-portrait One Time report:
 *  - Org header + info block repeat on every page (no year shown, since
 *    One Time contributions aren't year-scoped)
 *  - Transaction table (with SlNo, header row) repeats on every page, 25 rows per page
 *  - Grand-total row (SUM of Amount) only appears on the final page
 *  - "Page X of Y" centered under each page's table
 *  - Footer disclaimer appears once, after the last page's content
 */
function buildOneTimeReportDoc(category, generatedNow, rows) {
  var doc = DocumentApp.create('TEMP_OneTime_' + category + '_' + new Date().getTime());
  var body = doc.getBody();

  body.setPageWidth(PAGE_WIDTH_PORTRAIT).setPageHeight(PAGE_HEIGHT_PORTRAIT);
  body.setMarginTop(PAGE_MARGIN).setMarginBottom(PAGE_MARGIN)
      .setMarginLeft(PAGE_MARGIN).setMarginRight(PAGE_MARGIN);

  rows.sort(function (a, b) { return String(a.member).localeCompare(String(b.member)); });

  var totalPages = Math.max(1, Math.ceil(rows.length / ONE_TIME_ROWS_PER_PAGE));

  var grandTotal = 0;
  rows.forEach(function (r) { grandTotal += r.amount; });

  for (var p = 0; p < totalPages; p++) {
    if (p > 0) body.appendPageBreak();

    appendReportHeaderBlock(body, category, '', generatedNow, 'One Time', PAGE_WIDTH_PORTRAIT - (PAGE_MARGIN * 2));

    var chunk = rows.slice(p * ONE_TIME_ROWS_PER_PAGE, (p + 1) * ONE_TIME_ROWS_PER_PAGE);
    var isLastPage = (p === totalPages - 1);

    appendOneTimeTransactionTable(body, chunk, isLastPage ? grandTotal : null, p * ONE_TIME_ROWS_PER_PAGE);

    body.appendParagraph('');
    var pageLabel = body.appendParagraph('Page ' + (p + 1) + ' of ' + totalPages);
    pageLabel.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    pageLabel.editAsText().setFontSize(9).setForegroundColor('#888888');
  }

  var footer = body.appendParagraph('This is a system-generated report and does not require a physical or digital signature.');
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footer.editAsText().setItalic(true).setFontSize(9).setForegroundColor('#888888');

  doc.saveAndClose();
  return doc;
}

/**
 * Builds one page's transaction table: SlNo, Member, Mobile, Transaction Date, Amount, TransactionBy.
 * Identical to the Yearly report's table, just without the Year column.
 * grandTotal is null on all but the last page (no TOTAL row shown there).
 * startIndex is the running SlNo offset so numbering continues correctly across pages.
 */
function appendOneTimeTransactionTable(body, chunkRows, grandTotal, startIndex) {
  var tableData = [['SlNo', 'Member', 'Mobile', 'Transaction Date', 'Amount', 'TransactionBy']];

  chunkRows.forEach(function (r, i) {
    var slNo = startIndex + i + 1;
    var parsed = splitMemberAndMobile(r.member);
    var line = [
      String(slNo),
      parsed.name.toUpperCase(),
      parsed.mobile,
      formatYearlyDate(r.transactionDate),
      formatIndianNumber(r.amount),
      r.transactionBy ? String(r.transactionBy) : ''
    ];
    tableData.push(line);
  });

  var hasTotalRow = (grandTotal !== null && grandTotal !== undefined);
  if (hasTotalRow) {
    tableData.push(['', 'TOTAL', '', '', formatIndianNumber(grandTotal), '']);
  }

  var table = body.appendTable(tableData);
  styleOneTimeTable(table, hasTotalRow);
  setOneTimeTableColumnWidths(table);
  return table;
}

/** SlNo/Mobile/Transaction Date/TransactionBy centered, Amount right-aligned; Member left-aligned. */
function setOneTimeTableColumnWidths(table) {
  var contentWidth = PAGE_WIDTH_PORTRAIT - (PAGE_MARGIN * 2);
  var slNoWidth = 35;
  var mobileWidth = 60; // fits a 10-digit phone number comfortably at 9pt
  var dateWidth = 90;
  var amountWidth = 60;
  var memberWidth = 180;
  var transactionByWidth = contentWidth - (slNoWidth + mobileWidth + dateWidth + amountWidth + memberWidth);

  table.setColumnWidth(0, slNoWidth);
  table.setColumnWidth(1, memberWidth);
  table.setColumnWidth(2, mobileWidth);
  table.setColumnWidth(3, dateWidth);
  table.setColumnWidth(4, amountWidth);
  table.setColumnWidth(5, transactionByWidth);
}

function styleOneTimeTable(table, hasTotalRow) {
  var CENTERED_COLUMNS = [0, 2, 3, 5]; // SlNo, Mobile, Transaction Date, TransactionBy
  var RIGHT_ALIGNED_COLUMNS = [4]; // Amount

  var headerRow = table.getRow(0);
  for (var c = 0; c < headerRow.getNumCells(); c++) {
    var cell = headerRow.getCell(c);
    cell.setBackgroundColor('#f1f1f1'); // same shading as the other reports' header row
    cell.editAsText().setBold(true).setFontSize(9);
    if (CENTERED_COLUMNS.indexOf(c) !== -1) setCellAlignment(cell, DocumentApp.HorizontalAlignment.CENTER);
    if (RIGHT_ALIGNED_COLUMNS.indexOf(c) !== -1) setCellAlignment(cell, DocumentApp.HorizontalAlignment.RIGHT);
  }

  if (hasTotalRow) {
    var totalRow = table.getRow(table.getNumRows() - 1);
    for (var c2 = 0; c2 < totalRow.getNumCells(); c2++) {
      var totalCell = totalRow.getCell(c2);
      totalCell.setBackgroundColor('#f1f1f1'); // same shading as the header row
      totalCell.editAsText().setBold(true).setFontSize(9);
      if (CENTERED_COLUMNS.indexOf(c2) !== -1) setCellAlignment(totalCell, DocumentApp.HorizontalAlignment.CENTER);
      if (RIGHT_ALIGNED_COLUMNS.indexOf(c2) !== -1) setCellAlignment(totalCell, DocumentApp.HorizontalAlignment.RIGHT);
    }
  }

  for (var r = 1; r < table.getNumRows(); r++) {
    if (hasTotalRow && r === table.getNumRows() - 1) continue; // already styled above
    var row = table.getRow(r);
    for (var c3 = 0; c3 < row.getNumCells(); c3++) {
      var dataCell = row.getCell(c3);
      dataCell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(3).setPaddingRight(3);
      dataCell.editAsText().setFontSize(9);
      if (CENTERED_COLUMNS.indexOf(c3) !== -1) setCellAlignment(dataCell, DocumentApp.HorizontalAlignment.CENTER);
      if (RIGHT_ALIGNED_COLUMNS.indexOf(c3) !== -1) setCellAlignment(dataCell, DocumentApp.HorizontalAlignment.RIGHT);
    }
  }
}

/* -------------------------------------------------------------------------
 * Account Statement report
 * (Report Type = "Statement" in AppScriptReports — not a real Category, so
 * it's dispatched before the frequency-based checks. Pulls from
 * AccountStatement, filtered to the report's Year.)
 * ---------------------------------------------------------------------- */

function generateStatementReport(yearRaw, filePath) {
  var year = fiscalYearToCalendarYear(yearRaw);
  var rows = getStatementRows(year);

  if (rows.length === 0) {
    Logger.log('No AccountStatement rows found for Year=' + year + ' (generating report anyway).');
  }

  // Same convention as Yearly/One Time: "Generated Date Time" is the actual
  // moment of generation.
  var generatedNow = new Date();

  var doc = buildStatementReportDoc(year, generatedNow, rows);
  var docFile = DriveApp.getFileById(doc.getId());
  var pdfBlob = docFile.getAs('application/pdf');
  pdfBlob.setName(getFileNameFromPath(filePath));

  saveBlobToPath(pdfBlob, filePath);

  docFile.setTrashed(true); // clean up the temp Google Doc, keep only the PDF

  return true;
}

/** Pulls all AccountStatement rows whose Transaction Date falls in the given calendar year. */
function getStatementRows(year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(STATEMENT_SOURCE_SHEET);
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });

  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  var results = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var dateVal = row[idx['Transaction Date']];
    if (!dateVal) continue;

    var d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(d.getTime())) continue;
    if (String(d.getFullYear()) !== String(year)) continue;

    results.push({
      date: d,
      description: row[idx['Transaction Description']],
      referenceId: row[idx['ReferenceID']],
      transactionBy: row[idx['TransactionBy']],
      withdrawal: Number(row[idx['Withdrawal']]) || 0,
      deposit: Number(row[idx['Deposit']]) || 0,
      balance: Number(row[idx['Balance']]) || 0
    });
  }

  // Defensive re-sort by date, in case the sheet's own SORT formula ever changes.
  results.sort(function (a, b) { return a.date - b.date; });
  return results;
}

/** Sums Withdrawal/Deposit overall, plus broken out by Cash and UPI (TransactionBy). */
function computeStatementSummary(rows) {
  var cashWithdrawal = 0, cashDeposit = 0;
  var upiWithdrawal = 0, upiDeposit = 0;
  var totalWithdrawal = 0, totalDeposit = 0;

  rows.forEach(function (r) {
    var by = String(r.transactionBy || '').trim();
    totalWithdrawal += r.withdrawal;
    totalDeposit += r.deposit;
    if (by === 'Cash') {
      cashWithdrawal += r.withdrawal;
      cashDeposit += r.deposit;
    } else if (by === 'UPI') {
      upiWithdrawal += r.withdrawal;
      upiDeposit += r.deposit;
    }
  });

  return {
    cashWithdrawal: cashWithdrawal, cashDeposit: cashDeposit,
    upiWithdrawal: upiWithdrawal, upiDeposit: upiDeposit,
    totalWithdrawal: totalWithdrawal, totalDeposit: totalDeposit
  };
}

/**
 * Builds the multi-page A4-portrait Account Statement report:
 *  - Org header + "Account Statement for FY <year>" / Generated Date Time line repeat on every page
 *  - Cash/UPI/TOTAL summary table appears once, on page 1 only
 *  - Transaction table (with header row) repeats on every page, 25 rows per page
 *  - TOTAL row (SUM of Withdrawal/Deposit across the whole year) only on the final page
 *  - "Page X of Y" centered under each page's table
 *  - Footer disclaimer appears once, after the last page's content
 */
function buildStatementReportDoc(year, generatedNow, rows) {
  var doc = DocumentApp.create('TEMP_Statement_' + year + '_' + new Date().getTime());
  var body = doc.getBody();

  body.setPageWidth(PAGE_WIDTH_PORTRAIT).setPageHeight(PAGE_HEIGHT_PORTRAIT);
  body.setMarginTop(PAGE_MARGIN).setMarginBottom(PAGE_MARGIN)
      .setMarginLeft(PAGE_MARGIN).setMarginRight(PAGE_MARGIN);

  var contentWidth = PAGE_WIDTH_PORTRAIT - (PAGE_MARGIN * 2);
  var chunks = computeStatementPageChunks(rows);
  var totalPages = chunks.length;
  var summary = computeStatementSummary(rows);

  for (var p = 0; p < totalPages; p++) {
    if (p > 0) body.appendPageBreak();

    appendStatementHeaderBlock(body, year, generatedNow, contentWidth);

    if (p === 0) {
      appendStatementSummaryTable(body, summary, contentWidth);
      body.appendParagraph('');
    }

    var chunk = chunks[p];
    var isLastPage = (p === totalPages - 1);
    var pageTotals = isLastPage ? { withdrawal: summary.totalWithdrawal, deposit: summary.totalDeposit } : null;

    appendStatementTransactionTable(body, chunk, pageTotals);

    body.appendParagraph('');
    var pageLabel = body.appendParagraph('Page ' + (p + 1) + ' of ' + totalPages);
    pageLabel.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    pageLabel.editAsText().setFontSize(9).setForegroundColor('#888888');
  }

  var footer = body.appendParagraph('This is a system-generated report and does not require a physical or digital signature.');
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footer.editAsText().setItalic(true).setFontSize(9).setForegroundColor('#888888');

  doc.saveAndClose();
  return doc;
}

/**
 * Splits rows into pages using STATEMENT_ROWS_FIRST_PAGE for page 1 and
 * STATEMENT_ROWS_OTHER_PAGES for every page after that. Always returns at
 * least one (possibly empty) chunk, so a report with zero rows still
 * renders a single page.
 */
function computeStatementPageChunks(rows) {
  var chunks = [];
  var idx = 0;
  var isFirst = true;

  while (idx < rows.length) {
    var pageSize = isFirst ? STATEMENT_ROWS_FIRST_PAGE : STATEMENT_ROWS_OTHER_PAGES;
    chunks.push(rows.slice(idx, idx + pageSize));
    idx += pageSize;
    isFirst = false;
  }

  if (chunks.length === 0) chunks.push([]);
  return chunks;
}

/** The repeating top section: org name/address, then "Account Statement for FY <year>" left / Generated Date Time right. */
function appendStatementHeaderBlock(body, year, generatedNow, contentWidth) {
  var title = body.appendParagraph(ORG_NAME);
  title.setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.LEFT).setFontFamily("Special Gothic Condensed One");

  var address = body.appendParagraph(ORG_ADDRESS);
  address.setAlignment(DocumentApp.HorizontalAlignment.LEFT).setFontFamily("Arimo");
  address.editAsText().setFontSize(10).setForegroundColor('#555555');

  body.appendParagraph('');

  var genDateStr = Utilities.formatDate(new Date(generatedNow), Session.getScriptTimeZone(), "d MMMM yyyy hh:mm a");

  var titleRow = body.appendTable([
    ['Account Statement for FY ' + year, 'Generated Date Time : ' + genDateStr]
  ]);
  titleRow.setBorderWidth(0); // borderless — this is a layout row, not a data table

  var leftCell = titleRow.getRow(0).getCell(0);
  var rightCell = titleRow.getRow(0).getCell(1);
  leftCell.editAsText().setBold(true);
  setCellAlignment(leftCell, DocumentApp.HorizontalAlignment.LEFT);
  setCellAlignment(rightCell, DocumentApp.HorizontalAlignment.RIGHT);
  leftCell.setPaddingTop(2).setPaddingBottom(6).setPaddingLeft(0).setPaddingRight(0);
  rightCell.setPaddingTop(2).setPaddingBottom(6).setPaddingLeft(0).setPaddingRight(0);

  titleRow.setColumnWidth(0, contentWidth * 0.6);
  titleRow.setColumnWidth(1, contentWidth * 0.4);
}

/** Page-1-only Cash/UPI/TOTAL breakdown table. */
function appendStatementSummaryTable(body, summary, contentWidth) {
  var tableData = [
    ['TransactionBy', 'Withdrawal', 'Deposit'],
    ['Cash', formatIndianNumber(summary.cashWithdrawal), formatIndianNumber(summary.cashDeposit)],
    ['UPI', formatIndianNumber(summary.upiWithdrawal), formatIndianNumber(summary.upiDeposit)],
    ['TOTAL', formatIndianNumber(summary.totalWithdrawal), formatIndianNumber(summary.totalDeposit)]
  ];

  var table = body.appendTable(tableData);

  var col1 = contentWidth * 0.4;
  var col2 = (contentWidth - col1) / 2;
  table.setColumnWidth(0, col1);
  table.setColumnWidth(1, col2);
  table.setColumnWidth(2, col2);

  var lastRowIndex = table.getNumRows() - 1;
  for (var r = 0; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    var isHeaderOrTotal = (r === 0 || r === lastRowIndex);
    for (var c = 0; c < row.getNumCells(); c++) {
      var cell = row.getCell(c);
      cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(3).setPaddingRight(3);
      cell.editAsText().setFontSize(9);
      if (isHeaderOrTotal) {
        cell.setBackgroundColor('#f1f1f1');
        cell.editAsText().setBold(true);
      }
      if (c !== 0) setCellAlignment(cell, DocumentApp.HorizontalAlignment.RIGHT);
    }
  }
}

/**
 * Builds one page's transaction table: Date, Transaction Description,
 * Mode, Withdrawal, Deposit, Balance. (Ref ID removed — its width was
 * merged into Description.)
 * totals (Withdrawal/Deposit sums for the whole year) is null on all but the
 * last page — that's when the TOTAL row is appended.
 */
function appendStatementTransactionTable(body, chunkRows, totals) {
  var tableData = [['Date', 'Transaction Description', 'Mode', 'Withdrawal', 'Deposit', 'Balance']];

  chunkRows.forEach(function (r) {
    tableData.push([
      formatYearlyDate(r.date),
      stripMobileFromDescription(r.description),
      r.transactionBy ? String(r.transactionBy) : '',
      r.withdrawal ? formatIndianNumber(r.withdrawal) : '',
      r.deposit ? formatIndianNumber(r.deposit) : '',
      formatIndianNumber(r.balance)
    ]);
  });

  var hasTotalRow = !!totals;
  if (hasTotalRow) {
    tableData.push(['', 'TOTAL', '', formatIndianNumber(totals.withdrawal), formatIndianNumber(totals.deposit), '']);
  }

  var table = body.appendTable(tableData);
  styleStatementTable(table, hasTotalRow);
  setStatementTableColumnWidths(table);
  return table;
}

/**
 * Strips trailing "Member: Mobile" style mobile numbers out of a description,
 * e.g. "ShingLup Contribution - Pukhrambam Dorendrajit Singh: 9731399725"
 * -> "ShingLup Contribution - Pukhrambam Dorendrajit Singh".
 */
function stripMobileFromDescription(desc) {
  if (!desc) return '';
  var cleaned = String(desc).replace(/:\s?\d{6,}/g, '');
  cleaned = cleaned.replace(/\s{2,}/g, ' ').trim();
  return cleaned;
}

/** Date/Mode centered; Withdrawal/Deposit/Balance right-aligned; Description left-aligned. */
function setStatementTableColumnWidths(table) {
  var contentWidth = PAGE_WIDTH_PORTRAIT - (PAGE_MARGIN * 2);
  var dateWidth = 50;
  var transactionByWidth = 35;
  var withdrawalWidth = 55;
  var depositWidth = 55;
  var balanceWidth = 60;
  var descriptionWidth = contentWidth - (dateWidth + transactionByWidth + withdrawalWidth + depositWidth + balanceWidth);

  table.setColumnWidth(0, dateWidth);
  table.setColumnWidth(1, descriptionWidth);
  table.setColumnWidth(2, transactionByWidth);
  table.setColumnWidth(3, withdrawalWidth);
  table.setColumnWidth(4, depositWidth);
  table.setColumnWidth(5, balanceWidth);
}

function styleStatementTable(table, hasTotalRow) {
  var CENTERED_COLUMNS = [0, 2]; // Date, Mode
  var RIGHT_ALIGNED_COLUMNS = [3, 4, 5]; // Withdrawal, Deposit, Balance
  var FONT_SIZE = 8; // smaller than other reports' 9pt, to fit more rows per page

  var headerRow = table.getRow(0);
  for (var c = 0; c < headerRow.getNumCells(); c++) {
    var cell = headerRow.getCell(c);
    cell.setBackgroundColor('#f1f1f1');
    cell.editAsText().setBold(true).setFontSize(FONT_SIZE);
    if (CENTERED_COLUMNS.indexOf(c) !== -1) setCellAlignment(cell, DocumentApp.HorizontalAlignment.CENTER);
    if (RIGHT_ALIGNED_COLUMNS.indexOf(c) !== -1) setCellAlignment(cell, DocumentApp.HorizontalAlignment.RIGHT);
  }

  if (hasTotalRow) {
    var totalRow = table.getRow(table.getNumRows() - 1);
    for (var c2 = 0; c2 < totalRow.getNumCells(); c2++) {
      var totalCell = totalRow.getCell(c2);
      totalCell.setBackgroundColor('#f1f1f1'); // same shading as the header row
      totalCell.editAsText().setBold(true).setFontSize(FONT_SIZE);
      if (CENTERED_COLUMNS.indexOf(c2) !== -1) setCellAlignment(totalCell, DocumentApp.HorizontalAlignment.CENTER);
      if (RIGHT_ALIGNED_COLUMNS.indexOf(c2) !== -1) setCellAlignment(totalCell, DocumentApp.HorizontalAlignment.RIGHT);
    }
  }

  for (var r = 1; r < table.getNumRows(); r++) {
    if (hasTotalRow && r === table.getNumRows() - 1) continue; // already styled above
    var row = table.getRow(r);
    for (var c3 = 0; c3 < row.getNumCells(); c3++) {
      var dataCell = row.getCell(c3);
      dataCell.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(3).setPaddingRight(3);
      dataCell.editAsText().setFontSize(FONT_SIZE);
      if (CENTERED_COLUMNS.indexOf(c3) !== -1) setCellAlignment(dataCell, DocumentApp.HorizontalAlignment.CENTER);
      if (RIGHT_ALIGNED_COLUMNS.indexOf(c3) !== -1) setCellAlignment(dataCell, DocumentApp.HorizontalAlignment.RIGHT);
    }
  }
}

/* -------------------------------------------------------------------------
 * Member report
 * (Report Type = "Member" — not a real Category, dispatched before the
 * frequency-based checks, same as "Statement".)
 *
 * IMPORTANT: This report is built with SlidesApp instead of DocumentApp.
 * Google Docs' image API only supports resize/stretch, not true crop-to-fill
 * — SlidesApp.Image.replace(source, true) is the documented way to get real
 * "cover crop" behavior with uniform boxes regardless of source image size.
 * A Slides deck still exports to a single combined PDF (same as Docs), so
 * the final output is still one PDF file, just built via a different engine
 * for this one report type.
 * ---------------------------------------------------------------------- */

function generateMemberReport(filePath) {
  var rows = getMemberReportRows();

  if (rows.length === 0) {
    Logger.log('No rows found in Form Responses 1 for Member report (generating anyway).');
  }

  var presentation = buildMemberReportPresentation(rows);
  var presoId = presentation.getId();
  var file = DriveApp.getFileById(presoId);
  var pdfBlob = file.getAs('application/pdf');
  pdfBlob.setName(getFileNameFromPath(filePath));

  saveBlobToPath(pdfBlob, filePath);

  file.setTrashed(true); // clean up the temp Slides deck, keep only the PDF

  return true;
}

/**
 * Finds the first header in `headers` matching any of `candidates` (in order).
 * Returns -1 if none match. Used for columns whose exact header text isn't
 * confirmed (Gender / Email Address / Registration Date) so the report
 * degrades gracefully (blank column) instead of throwing if the real header
 * text differs slightly from these guesses.
 */
function findColumnIndex(headers, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var idx = headers.indexOf(candidates[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

/** Normalizes a mobile number cell (which may be stored as a number) to a plain digit string. */
function formatMobileNumber(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (typeof v === 'number') return String(Math.round(v));
  return String(v).trim();
}

/**
 * Pulls Name/Mobile/Gender/Email/RegistrationDate/PhotoURL/PhotoPath from
 * Form Responses 1. Name (col C) and Mobile (col E) use the same fixed
 * positions already proven correct elsewhere in this file (the Helper/
 * YearlyViewHelper sheets). Gender/Email/RegistrationDate use header-name
 * lookup with fallbacks since their exact header text isn't confirmed —
 * verify these show up correctly on first run.
 */
function getMemberReportRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MEMBER_FORM_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });

  var nameIdx = 2;   // column C
  var mobileIdx = 4; // column E
  var genderIdx = findColumnIndex(headers, ['Gender', 'Sex']);
  var emailIdx = findColumnIndex(headers, ['Email Address', 'Email']);
  var regDateIdx = findColumnIndex(headers, ['Registration Date', 'Timestamp']);
  var photoUrlIdx = findColumnIndex(headers, ['Member Photo']);
  var photoPathIdx = findColumnIndex(headers, ['PhotoPath']);

  var results = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var name = row[nameIdx];
    if (!name) continue;

    results.push({
      name: String(name).trim(),
      mobile: formatMobileNumber(row[mobileIdx]),
      gender: genderIdx !== -1 ? row[genderIdx] : '',
      email: emailIdx !== -1 ? row[emailIdx] : '',
      registrationDate: regDateIdx !== -1 ? row[regDateIdx] : '',
      photoUrl: photoUrlIdx !== -1 ? row[photoUrlIdx] : '',
      photoPath: photoPathIdx !== -1 ? row[photoPathIdx] : ''
    });
  }
  return results;
}

/**
 * Builds the full Member report: paginated member-list slides first
 * (SlNo/Member Name/Mobile No/Gender/Email Address/Registration Date),
 * followed by the photo-collage slides (PHOTO_GRID_COLUMNS x PHOTO_GRID_ROWS
 * per page). "Page X of Y" spans the whole deck (list pages + photo pages).
 */
function buildMemberReportPresentation(rows) {
  var presentation = SlidesApp.create('TEMP_Member_' + new Date().getTime());
  presentation.setPageSize(PAGE_WIDTH_PORTRAIT, PAGE_HEIGHT_PORTRAIT);

  // Slides always creates one default slide — remove it, we add our own.
  var defaultSlides = presentation.getSlides();
  for (var d = 0; d < defaultSlides.length; d++) defaultSlides[d].remove();

  var contentWidth = PAGE_WIDTH_PORTRAIT - (PAGE_MARGIN * 2);

  rows.sort(function (a, b) { return String(a.name).localeCompare(String(b.name)); });

  var listPageCount = Math.max(1, Math.ceil(rows.length / MEMBER_ROWS_PER_PAGE));
  var photosPerPage = PHOTO_GRID_COLUMNS * PHOTO_GRID_ROWS;
  var photoPageCount = Math.max(1, Math.ceil(rows.length / photosPerPage));
  var totalPages = listPageCount + photoPageCount;

  for (var p = 0; p < listPageCount; p++) {
    var slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    var topY = appendSlideHeaderBlock(slide, 'MEMBER LIST', contentWidth);
    var chunk = rows.slice(p * MEMBER_ROWS_PER_PAGE, (p + 1) * MEMBER_ROWS_PER_PAGE);
    insertMemberListTable(slide, chunk, p * MEMBER_ROWS_PER_PAGE, topY, contentWidth);
    appendSlidePageLabel(slide, p + 1, totalPages, contentWidth);
  }

  appendMemberPhotoSlides(presentation, rows, listPageCount + 1, totalPages, contentWidth);

  return presentation;
}

/** Org name/address + subtitle, same content as the Docs-based reports' header. Returns the Y position content should start at. */
function appendSlideHeaderBlock(slide, subtitle, contentWidth) {
  var title = slide.insertTextBox(ORG_NAME, PAGE_MARGIN, PAGE_MARGIN, contentWidth, 22);
  title.getText().getTextStyle().setBold(true).setFontSize(16);

  var address = slide.insertTextBox(ORG_ADDRESS, PAGE_MARGIN, PAGE_MARGIN + 22, contentWidth, 16);
  address.getText().getTextStyle().setFontSize(10).setForegroundColor('#555555');

  var sub = slide.insertTextBox(subtitle, PAGE_MARGIN, PAGE_MARGIN + 44, contentWidth, 18);
  sub.getText().getTextStyle().setBold(true).setFontSize(11);

  return PAGE_MARGIN + 70; // Y position where table/grid content should start
}

function appendSlidePageLabel(slide, pageNum, totalPages, contentWidth) {
  var label = slide.insertTextBox('Page ' + pageNum + ' of ' + totalPages,
    PAGE_MARGIN, PAGE_HEIGHT_PORTRAIT - PAGE_MARGIN - 16, contentWidth, 14);
  label.getText().getTextStyle().setFontSize(8).setForegroundColor('#888888');
  label.getText().getParagraphs()[0].getRange().getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

/** Builds the SlNo/Member Name/Mobile No/Gender/Email Address/Registration Date table for one page. */
function insertMemberListTable(slide, chunkRows, startIndex, topY, contentWidth) {
  var headerRow = ['SlNo', 'Member Name', 'Mobile No', 'Gender', 'Email Address', 'Registration Date'];
  var numRows = chunkRows.length + 1;
  var table = slide.insertTable(numRows, headerRow.length, PAGE_MARGIN, topY, contentWidth, 16 * numRows);

  var colWidths = computeMemberTableColumnWidths(contentWidth);
  for (var c = 0; c < headerRow.length; c++) {
    var hCell = table.getCell(0, c);
    hCell.getText().setText(headerRow[c]);
    hCell.getText().getTextStyle().setBold(true).setFontSize(8);
    hCell.getFill().setSolidFill('#f1f1f1');
    table.getColumn(c).setWidth(colWidths[c]);
  }
  table.getRow(0).setMinimumHeight(16);

  chunkRows.forEach(function (r, i) {
    var slNo = startIndex + i + 1;
    var values = [
      String(slNo),
      r.name.toUpperCase(),
      r.mobile,
      r.gender ? String(r.gender) : '',
      r.email ? String(r.email) : '',
      formatYearlyDate(r.registrationDate)
    ];
    for (var c2 = 0; c2 < values.length; c2++) {
      var cell = table.getCell(i + 1, c2);
      cell.getText().setText(values[c2]);
      cell.getText().getTextStyle().setFontSize(8);
    }
    table.getRow(i + 1).setMinimumHeight(15);
  });

  return table;
}

function computeMemberTableColumnWidths(contentWidth) {
  var slNoWidth = 30;
  var mobileWidth = 65;
  var genderWidth = 50;
  var regDateWidth = 70;
  var nameWidth = 140;
  var emailWidth = contentWidth - (slNoWidth + mobileWidth + genderWidth + regDateWidth + nameWidth);
  return [slNoWidth, nameWidth, mobileWidth, genderWidth, emailWidth, regDateWidth];
}

/**
 * Builds the photo-collage pages: PHOTO_GRID_COLUMNS x PHOTO_GRID_ROWS uniform
 * boxes per page, each photo crop-filled to the box (see insertFilledPhoto),
 * with the member's name and mobile centered below each photo.
 */
function appendMemberPhotoSlides(presentation, rows, startPageNumber, totalPages, contentWidth) {
  var perPage = PHOTO_GRID_COLUMNS * PHOTO_GRID_ROWS;
  var totalPhotoPages = Math.max(1, Math.ceil(rows.length / perPage));
  var folderCache = {};

  var headerReserve = 70; // matches appendSlideHeaderBlock's returned offset
  var contentHeight = PAGE_HEIGHT_PORTRAIT - (PAGE_MARGIN * 2) - headerReserve;

  var cellWidth = (contentWidth - PHOTO_GRID_GAP * (PHOTO_GRID_COLUMNS - 1)) / PHOTO_GRID_COLUMNS;
  var cellHeight = (contentHeight - PHOTO_GRID_GAP * (PHOTO_GRID_ROWS - 1)) / PHOTO_GRID_ROWS;
  var textHeight = 22;
  var photoHeight = cellHeight - textHeight - 4;
  var photoWidth = cellWidth;

  for (var p = 0; p < totalPhotoPages; p++) {
    var slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    var topY = appendSlideHeaderBlock(slide, 'MEMBER PHOTOS', contentWidth);

    var chunk = rows.slice(p * perPage, (p + 1) * perPage);
    chunk.forEach(function (member, i) {
      var col = i % PHOTO_GRID_COLUMNS;
      var gridRow = Math.floor(i / PHOTO_GRID_COLUMNS);
      var cellLeft = PAGE_MARGIN + col * (cellWidth + PHOTO_GRID_GAP);
      var cellTop = topY + gridRow * (cellHeight + PHOTO_GRID_GAP);

      var blob = getMemberPhotoBlob(member, folderCache);
      if (blob) {
        insertFilledPhoto(slide, blob, cellLeft, cellTop, photoWidth, photoHeight);
      } else {
        var placeholder = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cellLeft, cellTop, photoWidth, photoHeight);
        placeholder.getFill().setSolidFill('#eeeeee');
        placeholder.getBorder().setTransparent();
        var noPhotoText = placeholder.getText();
        noPhotoText.setText('No Photo');
        noPhotoText.getTextStyle().setFontSize(8).setForegroundColor('#999999');
      }

      var caption = slide.insertTextBox(member.name.toUpperCase() + '\n' + member.mobile,
        cellLeft, cellTop + photoHeight + 2, photoWidth, textHeight);
      caption.getText().getTextStyle().setFontSize(7);
      var paragraphs = caption.getText().getParagraphs();
      for (var pg = 0; pg < paragraphs.length; pg++) {
        paragraphs[pg].getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      }
    });

    appendSlidePageLabel(slide, startPageNumber + p, totalPages, contentWidth);
  }

  return totalPhotoPages;
}

/**
 * Inserts an image at the exact box size with true crop-to-fill (cover),
 * instead of the default stretch. Slides is inserted at box size first
 * (which stretches), then .replace(blob, true) re-crops the same source
 * to fill that exact frame — this is the documented crop-on-replace
 * behavior. Falls back to the stretched version if the crop-replace call
 * fails for any reason (e.g. unsupported image format).
 */
function insertFilledPhoto(slide, blob, left, top, width, height) {
  var img = slide.insertImage(blob, left, top, width, height);
  try {
    img.replace(blob, true);
  } catch (e) {
    Logger.log('Photo crop-fill failed, using stretched image instead: ' + e);
  }
  return img;
}

/** Resolves a member's photo: prefers "Member Photo" (direct Drive file ID), falls back to "PhotoPath". */
function getMemberPhotoBlob(memberRow, folderCache) {
  if (memberRow.photoUrl) {
    var match = String(memberRow.photoUrl).match(/[-\w]{25,}/);
    if (match) {
      try {
        return DriveApp.getFileById(match[0]).getBlob();
      } catch (e) {
        // fall through to PhotoPath
      }
    }
  }

  if (memberRow.photoPath) {
    var file = getFileFromPhotoPath(memberRow.photoPath, folderCache);
    if (file) {
      try {
        return file.getBlob();
      } catch (e) {
        return null;
      }
    }
  }

  return null;
}

/** Like getFileUrlFromPath, but returns the actual Drive File object instead of a URL string. */
function getFileFromPhotoPath(path, folderCache) {
  folderCache = folderCache || {};
  var relativeParts = String(path).split('/').filter(Boolean);
  if (relativeParts.length === 0) return null;

  var fileName = relativeParts.pop();
  var folderParts = INVOICE_BASE_FOLDER_PATH.concat(relativeParts);
  var folder = getFolderByPath(folderParts, folderCache);
  if (!folder) return null;

  var files = folder.getFilesByName(fileName);
  if (!files.hasNext()) return null;
  return files.next();
}

/* -------------------------------------------------------------------------
 * Shared helpers
 * ---------------------------------------------------------------------- */

function styleInfoTable(table) {
  for (var r = 0; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    for (var c = 0; c < row.getNumCells(); c++) {
      var cell = row.getCell(c);
      cell.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(5).setPaddingRight(5);
      if (c % 2 === 0) {
        // Label columns (Report Name, Category Name, Collection Frequency, Generated DateTime)
        cell.setBackgroundColor('#f1f1f1');
        cell.editAsText().setBold(true);
      } else {
        // Value columns
        cell.setBackgroundColor('#ffffff');
      }
    }
  }
}

/** Centers (or otherwise aligns) all paragraph text inside a table cell. */
function setCellAlignment(cell, alignment) {
  for (var i = 0; i < cell.getNumChildren(); i++) {
    var child = cell.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      child.asParagraph().setAlignment(alignment);
    }
  }
}

function getFileNameFromPath(path) {
  var parts = String(path).split('/').filter(Boolean);
  return parts[parts.length - 1];
}

/** Saves a blob into Drive at a path like "/GeneratedReports/abc123.pdf", creating folders as needed. */
function saveBlobToPath(blob, path) {
  var relativeParts = String(path).split('/').filter(Boolean);
  var fileName = relativeParts.pop();
  var folderParts = INVOICE_BASE_FOLDER_PATH.concat(relativeParts);
  var folder = getOrCreateFolderByPath(folderParts);

  var existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true); // overwrite if regenerating

  folder.createFile(blob).setName(fileName);
}

/** Like getFolderByPath, but creates any missing folder segments instead of failing. */
function getOrCreateFolderByPath(folderNames) {
  var folder = DriveApp.getRootFolder();
  for (var i = 0; i < folderNames.length; i++) {
    var subFolders = folder.getFoldersByName(folderNames[i]);
    folder = subFolders.hasNext() ? subFolders.next() : folder.createFolder(folderNames[i]);
  }
  return folder;
}
