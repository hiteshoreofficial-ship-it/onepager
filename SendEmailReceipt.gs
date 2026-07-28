/* ==========================================================================
 * Contribution Receipt Emails
 *
 * Completely separate from Code.gs on purpose — this file only touches the
 * Contributions sheet and the ReceiptEmailBody.html template, so it can be
 * edited/restyled without any risk to the PDF report generation logic.
 * ========================================================================== */

var RECEIPT_CONTRIBUTIONS_SHEET = 'Contributions';
var RECEIPT_SENT_COLUMN_HEADER = 'ReceiptEmailSent'; // <-- add this column to Contributions (any position, TRUE/FALSE)
var RECEIPT_ORG_NAME = 'Keishamthong Apunba Chaokhat Marup';

// TEST MODE: every receipt email currently goes here instead of the real
// member, since the app isn't live yet. See sendReceiptEmailForRow() below
// for the one line to change when you're ready to go live.
var RECEIPT_TEST_EMAIL = 'doren007@gmail.com';

/**
 * Run this ONCE from the Apps Script editor to install the trigger that
 * watches Contributions for new/edited rows.
 *
 * IMPORTANT: If AppSheet inserts these rows via the Sheets API, onEdit will
 * NOT fire — only an installable onChange trigger catches API-driven edits
 * (same reasoning as every other onChange trigger in this project).
 *
 * Safe to re-run any time — it won't create duplicate triggers.
 */
function setupContributionReceiptTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onContributionSheetChangeForReceipt') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onContributionSheetChangeForReceipt')
    .forSpreadsheet(ss)
    .onChange()
    .create();
  return 'Contribution receipt email trigger installed.';
}

/**
 * Optional safety net: wire this function to a time-driven trigger
 * (Triggers UI -> Add Trigger -> processReceiptEmailSafetyNet -> Time-driven
 * -> every 5 minutes) in case onChange ever misses an edit.
 */
function processReceiptEmailSafetyNet() {
  processPendingReceiptEmails();
}

function onContributionSheetChangeForReceipt(e) {
  processPendingReceiptEmails();
}

/**
 * Scans Contributions for rows where ReceiptEmailSent isn't TRUE yet, sends
 * a receipt email for each, and marks it sent so it's never re-sent.
 */
function processPendingReceiptEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RECEIPT_CONTRIBUTIONS_SHEET);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });

  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i + 1; }); // 1-based column numbers

  if (!idx[RECEIPT_SENT_COLUMN_HEADER]) {
    Logger.log('Contributions is missing the "' + RECEIPT_SENT_COLUMN_HEADER +
      '" column — add it (any position, holds TRUE/FALSE) to enable receipt emails.');
    return;
  }

  var numRows = lastRow - 1;
  var data = sheet.getRange(2, 1, numRows, sheet.getLastColumn()).getValues();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var alreadySent = row[idx[RECEIPT_SENT_COLUMN_HEADER] - 1];
    if (alreadySent === true || String(alreadySent).toUpperCase() === 'TRUE') continue;

    var contributionId = idx['ContributionID'] ? row[idx['ContributionID'] - 1] : null;
    if (!contributionId) continue; // incomplete/blank row, skip for now

    try {
      var success = sendReceiptEmailForRow(row, idx);
      if (success) {
        sheet.getRange(2 + i, idx[RECEIPT_SENT_COLUMN_HEADER]).setValue(true);
      }
    } catch (err) {
      Logger.log('Failed to send receipt email for ContributionID ' + contributionId + ': ' + err);
      // Left as not-sent so it's retried on the next run.
    }
  }
}

/** Builds the email content and sends it. Returns true on success. */
function sendReceiptEmailForRow(row, idx) {
  var context = buildReceiptContext(row, idx);
  var htmlBody = buildReceiptEmailHtml(context);

  var subject = 'Contribution Receipt - ' + context.categoryName +
    (context.period ? ' (' + context.period + ')' : '');

  // ---- TEST MODE (current) ----
  var recipientEmail = RECEIPT_TEST_EMAIL;

  // ---- LIVE MODE -- uncomment these 5 lines and delete the test-mode line
  //      above when the app is ready to email real members ----
  // recipientEmail = lookupMemberEmailForReceipt(row[idx['Member'] - 1]);
  // if (!recipientEmail) {
  //   Logger.log('No email found for member "' + row[idx['Member'] - 1] + '" — skipping this row.');
  //   return false;
  // }

  MailApp.sendEmail({
    to: recipientEmail,
    subject: subject,
    htmlBody: htmlBody,
    body: 'Your contribution of Rs ' + context.amount + ' to ' + context.categoryName +
      ' has been recorded. Please view this email in an HTML-capable mail client to see the full receipt.'
  });

  return true;
}

/**
 * LIVE-MODE HELPER (not called yet — see the commented-out block in
 * sendReceiptEmailForRow above). Looks up a member's email from
 * Form Responses 1 by matching their name against the Contributions
 * row's Member field.
 */
function lookupMemberEmailForReceipt(memberString) {
  var parsed = splitMemberAndMobileForReceipt(memberString);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Form Responses 1');
  if (!sheet) return '';

  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });

  var emailIdx = headers.indexOf('Email Address');
  if (emailIdx === -1) emailIdx = headers.indexOf('Email');
  if (emailIdx === -1) return '';

  for (var r = 1; r < data.length; r++) {
    var name = String(data[r][2] || '').trim(); // column C, same position used elsewhere in this project
    if (name === parsed.name) {
      return String(data[r][emailIdx] || '').trim();
    }
  }
  return '';
}

/**
 * Pulls together everything the email template needs from one Contributions
 * row. Every field degrades to '' if the column isn't found, so a missing
 * column never throws — it just shows blank in the email.
 */
function buildReceiptContext(row, idx) {
  var memberRaw = idx['Member'] ? row[idx['Member'] - 1] : '';
  var parsedMember = splitMemberAndMobileForReceipt(memberRaw);

  var amount = idx['Amount'] ? (Number(row[idx['Amount'] - 1]) || 0) : 0;
  var categoryName = idx['CategoryName'] ? row[idx['CategoryName'] - 1] : '';
  var month = idx['Contribution Month'] ? row[idx['Contribution Month'] - 1] : '';
  var year = idx['Contribution Year'] ? row[idx['Contribution Year'] - 1] : '';
  var period = (month && year) ? (capitalizeMonthForReceipt(month) + '-' + year) : '';
  var frequency = idx['Contribution Frequncy'] ? row[idx['Contribution Frequncy'] - 1] : '';
  var transactionDateRaw = idx['Transaction Date'] ? row[idx['Transaction Date'] - 1] : '';
  var paymentMethod = idx['TransactionBy'] ? row[idx['TransactionBy'] - 1] : '';
  var remarks = idx['Remarks'] ? row[idx['Remarks'] - 1] : '';
  var invoiceUrl = idx['InvoiceUrl'] ? row[idx['InvoiceUrl'] - 1] : '';

  return {
    orgName: RECEIPT_ORG_NAME,
    memberName: parsedMember.name,
    amount: formatIndianNumberForReceipt(amount),
    categoryName: String(categoryName || ''),
    period: period,
    frequency: String(frequency || ''),
    transactionDate: formatReceiptDate(transactionDateRaw),
    paymentMethod: String(paymentMethod || ''),
    recordedBy: Session.getActiveUser().getEmail() || '',
    remarks: String(remarks || ''),
    invoiceUrl: String(invoiceUrl || '')
  };
}

/** Renders ReceiptEmailBody.html with `context` available as `data`. */
function buildReceiptEmailHtml(context) {
  var template = HtmlService.createTemplateFromFile('ReceiptEmailBody');
  template.data = context;
  return template.evaluate().getContent();
}

/* -------------------------------------------------------------------------
 * Small local helpers (deliberately duplicated from Code.gs rather than
 * shared, so this file stays fully independent and editable on its own).
 * ---------------------------------------------------------------------- */

/** Splits "AHANTHEM JODHA: 9089680302" into {name: "AHANTHEM JODHA", mobile: "9089680302"}. */
function splitMemberAndMobileForReceipt(memberString) {
  var parts = String(memberString || '').split(':');
  var mobile = parts.length > 1 ? parts.pop().trim() : '';
  var name = parts.join(':').trim();
  return { name: name, mobile: mobile };
}

/** "JUN" -> "Jun" */
function capitalizeMonthForReceipt(m) {
  var s = String(m || '').trim();
  if (!s) return '';
  return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
}

/** Formats a Transaction Date value as "23 July 2026 10:36 AM". */
function formatReceiptDate(value) {
  if (!value) return '';
  var d = (value instanceof Date) ? value : new Date(value);
  if (isNaN(d.getTime())) return String(value);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "d MMMM yyyy hh:mm a");
}

/** Indian-style thousands separators, e.g. 120000 -> "1,20,000". */
function formatIndianNumberForReceipt(num) {
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
