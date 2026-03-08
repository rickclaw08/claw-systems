/**
 * Google Apps Script - ClawOps Intake Form Handler
 *
 * SETUP INSTRUCTIONS:
 * 1. Go to https://script.google.com and create a new project
 * 2. Paste this entire file into Code.gs
 * 3. Click Deploy > New Deployment
 * 4. Select "Web app" as the type
 * 5. Set "Execute as" to your Google account
 * 6. Set "Who has access" to "Anyone"
 * 7. Click Deploy and copy the Web App URL
 * 8. Paste that URL into the APPS_SCRIPT_URL variable in intake/index.html
 *
 * SHEET SETUP:
 * 1. Create a Google Sheet called "ClawOps Intake Submissions"
 * 2. Copy the Sheet ID from the URL (the long string between /d/ and /edit)
 * 3. Paste it into SHEET_ID below
 * 4. The script will auto-create headers on the first submission
 *
 * The sheet will have these columns:
 *   Timestamp | Full Name | Email | Phone | Business Name | Business Type |
 *   Calls Per Week | Phone Challenge | Referral Source | Source
 */

// REPLACE THIS with your Google Sheet ID
var SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE';

// Sheet name (tab name)
var SHEET_NAME = 'Submissions';

// Email to notify on new submission
var NOTIFY_EMAIL = 'contact@theclawops.com';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Basic server-side validation
    if (!data.fullName || !data.email || !data.phone) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Missing required fields'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Sanitize inputs (strip HTML)
    var fullName = stripHtml(data.fullName).substring(0, 100);
    var email = stripHtml(data.email).substring(0, 254);
    var phone = stripHtml(data.phone).substring(0, 20);
    var businessName = stripHtml(data.businessName || '').substring(0, 200);
    var businessType = stripHtml(data.businessType || '').substring(0, 50);
    var callsPerWeek = stripHtml(data.callsPerWeek || '').substring(0, 20);
    var phoneChallenge = stripHtml(data.phoneChallenge || '').substring(0, 2000);
    var referralSource = stripHtml(data.referralSource || '').substring(0, 200);
    var source = stripHtml(data.source || 'unknown').substring(0, 50);

    // Email format check
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Invalid email format'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Open the sheet
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      // Add headers
      sheet.appendRow([
        'Timestamp',
        'Full Name',
        'Email',
        'Phone',
        'Business Name',
        'Business Type',
        'Calls Per Week',
        'Phone Challenge',
        'Referral Source',
        'Source'
      ]);
      // Bold the header row
      sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
    }

    // Append the data row
    var timestamp = new Date().toISOString();
    sheet.appendRow([
      timestamp,
      fullName,
      email,
      phone,
      businessName,
      businessType,
      callsPerWeek,
      phoneChallenge,
      referralSource,
      source
    ]);

    // Send notification email
    if (NOTIFY_EMAIL) {
      try {
        var subject = '[ClawOps Intake] New lead: ' + (businessName || fullName);
        var body = 'New intake form submission:\n\n' +
          'Name: ' + fullName + '\n' +
          'Email: ' + email + '\n' +
          'Phone: ' + phone + '\n' +
          'Business: ' + businessName + '\n' +
          'Type: ' + businessType + '\n' +
          'Calls/Week: ' + callsPerWeek + '\n' +
          'Challenge: ' + phoneChallenge + '\n' +
          'Referral: ' + referralSource + '\n\n' +
          'Submitted: ' + timestamp;
        MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
      } catch (emailErr) {
        // Don't fail the whole thing if email fails
        Logger.log('Email notification failed: ' + emailErr.message);
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Submission recorded'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Error: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Server error'
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    service: 'ClawOps Intake Handler'
  })).setMimeType(ContentService.MimeType.JSON);
}

function stripHtml(str) {
  if (!str) return '';
  return str.toString().replace(/<[^>]*>/g, '').trim();
}
