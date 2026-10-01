// ============================================
// FLEXI VISITOR SIGN-IN -- EMAIL NOTIFIER (v2)
// ============================================
// This replaces the old Code.gs. Visitor data now lives in Supabase, not
// this spreadsheet -- this script's only job is sending the two Gmail
// notifications the kiosk used to trigger on every sign-in:
//   1. Arrival notification to the host
//   2. Optional "thanks for visiting" follow-up to the visitor
//
// DEPLOY: paste this over the existing Code.gs in the SAME Apps Script
// project, then Deploy > Manage deployments > (pencil icon) > New version.
// Keeping it in the same deployment keeps the same /exec URL, so nothing
// in visitor-kiosk.html needs to change.
// ============================================

const CONFIG = {
  EMAIL_FROM_NAME: 'Flexi Reception',
  EMAIL_SUBJECT: 'Visitor Arrival Notification',
  COMPANY_NAME: 'Narrow Aisle Ltd (Flexi)',
  COMPANY_ADDRESS: 'Tipton, West Midlands',

  // Where watchlist alerts (and anything addressed to "Reception") land.
  SECURITY_EMAIL: 'reception@narrowaisle.co.uk',

  // Optional: paste a Microsoft Teams "Incoming Webhook" URL here (right-click
  // the channel > Workflows / Connectors > Incoming Webhook) to also post
  // every arrival/delivery/watchlist alert into Teams. Leave blank to skip --
  // these are hooks, not requirements.
  TEAMS_WEBHOOK_URL: '',

  // Optional: fill in all three to also text the host via Twilio
  // (https://www.twilio.com -- Account SID + Auth Token from the console,
  // plus a Twilio phone number). Leave any one blank to skip SMS entirely.
  TWILIO_ACCOUNT_SID: '',
  TWILIO_AUTH_TOKEN: '',
  TWILIO_FROM_NUMBER: '',

  // Optional: map a host's email (as used on the kiosk) to their mobile
  // number so SMS knows who to text. Only hosts listed here get texted.
  HOST_SMS_NUMBERS: {
    // 'e.holt@narrowaisle.co.uk': '+447700900000',
  }
};

function corsResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doOptions(e) {
  return corsResponse({ status: 'ok' });
}

function doGet(e) {
  return corsResponse({ status: 'ok', service: 'flexi visitor mailer' });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.alertType === 'watchlist') {
      // Silent watchlist match -- the visitor never sees this, only reception.
      sendWatchlistAlertEmail(data);
    } else {
      // Normal arrival or delivery notification to the host/reception.
      sendEmailNotification(data);
      if (data.visitorEmail && data.visitorEmail.trim() !== '') {
        sendVisitorFollowUpEmail(data);
      }
    }

    // Both hooks below no-op silently until their CONFIG values are filled in.
    sendTeamsAlert(data);
    sendSmsAlert(data);

    return corsResponse({ success: true, message: 'Notification sent' });
  } catch (error) {
    Logger.log('Error in doPost: ' + error);
    return corsResponse({ success: false, message: error.toString() });
  }
}

// ============================================
// EMAIL FUNCTIONS (unchanged from the original system)
// ============================================

function sendEmailNotification(data) {
  const subject = CONFIG.EMAIL_SUBJECT;

  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
        .container { max-width: 600px; margin: 0 auto; padding: 20px; }
        .header {
          background: linear-gradient(135deg, #E31E24 0%, #B81820 100%);
          color: white; padding: 30px; text-align: center; border-radius: 8px 8px 0 0;
        }
        .header h1 { margin: 0; font-size: 24px; }
        .content { background: #f8f9fa; padding: 30px; border-radius: 0 0 8px 8px; }
        .info-box {
          background: white; padding: 20px; border-radius: 6px;
          margin: 20px 0; border-left: 4px solid #E31E24;
        }
        .info-row { margin: 10px 0; }
        .info-label { font-weight: bold; color: #666; }
        .info-value { color: #333; font-size: 16px; }
        .footer { text-align: center; margin-top: 20px; color: #666; font-size: 12px; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>🔔 Visitor Arrival Notification</h1>
        </div>
        <div class="content">
          <p>Hi <strong>${data.hostName}</strong>,</p>
          <p>You have a visitor waiting for you at reception.</p>
          <div class="info-box">
            <div class="info-row">
              <span class="info-label">Visitor Name:</span>
              <span class="info-value">${data.visitorName}</span>
            </div>
            <div class="info-row">
              <span class="info-label">Company:</span>
              <span class="info-value">${data.companyName}</span>
            </div>
            <div class="info-row">
              <span class="info-label">Purpose:</span>
              <span class="info-value">${data.visitPurpose}</span>
            </div>
            <div class="info-row">
              <span class="info-label">Time:</span>
              <span class="info-value">${data.timeIn} on ${data.date}</span>
            </div>
            ${data.vehicleReg ? `
            <div class="info-row">
              <span class="info-label">Vehicle:</span>
              <span class="info-value">${data.vehicleReg}</span>
            </div>` : ''}
          </div>
          <p><strong>Please proceed to reception when convenient.</strong></p>
          <div class="footer">
            <p>${CONFIG.COMPANY_NAME}</p>
            <p>${CONFIG.COMPANY_ADDRESS}</p>
            <p style="margin-top: 10px; font-size: 11px;">
              This is an automated notification from the Flexi Visitor Sign-In System.
            </p>
          </div>
        </div>
      </div>
    </body>
    </html>
  `;

  const plainBody = `Hi ${data.hostName},\n\nYou have a visitor waiting for you at reception.\n\nVisitor Details:\n- Name: ${data.visitorName}\n- Company: ${data.companyName}\n- Purpose: ${data.visitPurpose}\n- Time: ${data.timeIn} on ${data.date}\n${data.vehicleReg ? `- Vehicle: ${data.vehicleReg}` : ''}\n\nPlease proceed to reception when convenient.\n\n${CONFIG.COMPANY_NAME}\n${CONFIG.COMPANY_ADDRESS}`;

  try {
    GmailApp.sendEmail(data.hostEmail, subject, plainBody, {
      htmlBody: htmlBody,
      name: CONFIG.EMAIL_FROM_NAME
    });
    Logger.log('Host email sent successfully to: ' + data.hostEmail);
  } catch (error) {
    Logger.log('Error sending host email: ' + error);
  }
}

function sendVisitorFollowUpEmail(data) {
  const subject = 'Thank You for Visiting Narrow Aisle / Flexi';

  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
        .container { max-width: 600px; margin: 0 auto; padding: 20px; }
        .header {
          background: linear-gradient(135deg, #E31E24 0%, #B81820 100%);
          color: white; padding: 40px; text-align: center; border-radius: 8px 8px 0 0;
        }
        .header h1 { margin: 0; font-size: 28px; }
        .content { background: #f8f9fa; padding: 30px; border-radius: 0 0 8px 8px; }
        .visit-box {
          background: white; padding: 25px; border-radius: 6px;
          margin: 20px 0; border-left: 4px solid #E31E24;
        }
        .footer {
          text-align: center; margin-top: 30px; padding-top: 20px;
          border-top: 2px solid #E31E24; color: #666; font-size: 14px;
        }
        .website-link {
          display: inline-block; background: #E31E24; color: white;
          padding: 12px 30px; border-radius: 5px; text-decoration: none;
          margin: 20px 0; font-weight: bold;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>Thank You for Visiting!</h1>
        </div>
        <div class="content">
          <p>Dear <strong>${data.visitorName}</strong>,</p>
          <p>Thank you for visiting <strong>Narrow Aisle / Flexi</strong> today. We hope you enjoyed your visit.</p>
          <div class="visit-box">
            <h3 style="color: #E31E24; margin-top: 0;">Your Visit Details</h3>
            <p><strong>Date:</strong> ${data.date}</p>
            <p><strong>Time:</strong> ${data.timeIn}</p>
            <p><strong>Host:</strong> ${data.hostName}</p>
            <p><strong>Company:</strong> ${data.companyName}</p>
          </div>
          <p>We hope to see you again soon!</p>
          <div style="text-align: center;">
            <a href="https://www.flexi.co.uk" class="website-link">Visit www.flexi.co.uk</a>
          </div>
          <div class="footer">
            <p><strong>Narrow Aisle Ltd (Flexi)</strong></p>
            <p>Your Space, Our Expertise</p>
            <p>Tipton, West Midlands</p>
            <p><a href="https://www.flexi.co.uk" style="color: #E31E24;">www.flexi.co.uk</a></p>
            <p style="font-size: 12px; opacity: 0.8;">This is an automated message from the Flexi Visitor Sign-In System.</p>
          </div>
        </div>
      </div>
    </body>
    </html>
  `;

  const plainBody = `Dear ${data.visitorName},\n\nThank you for visiting Narrow Aisle / Flexi today.\n\nYour Visit Details:\n- Date: ${data.date}\n- Time: ${data.timeIn}\n- Host: ${data.hostName}\n- Company: ${data.companyName}\n\nWe hope to see you again soon!\n\nwww.flexi.co.uk\n\nNarrow Aisle Ltd (Flexi)\nTipton, West Midlands`;

  try {
    GmailApp.sendEmail(data.visitorEmail, subject, plainBody, {
      htmlBody: htmlBody,
      name: 'Flexi Reception'
    });
    Logger.log('Follow-up email sent successfully to: ' + data.visitorEmail);
  } catch (error) {
    Logger.log('Error sending follow-up email: ' + error);
  }
}

// ============================================
// WATCHLIST ALERT (silent -- reception only, never the visitor)
// ============================================
function sendWatchlistAlertEmail(data) {
  const subject = '⚠️ Watchlist Match — Visitor Signed In';
  const body =
    'A visitor matching a watchlist entry has just signed in at reception.\n\n' +
    'Name: ' + data.visitorName + '\n' +
    'Company: ' + (data.companyName || '—') + '\n' +
    'Reason on file: ' + (data.flaggedReason || '—') + '\n' +
    'Time: ' + data.timeIn + ' on ' + data.date + '\n\n' +
    'Please attend reception.';

  try {
    GmailApp.sendEmail(CONFIG.SECURITY_EMAIL, subject, body, { name: CONFIG.EMAIL_FROM_NAME });
    Logger.log('Watchlist alert email sent to: ' + CONFIG.SECURITY_EMAIL);
  } catch (error) {
    Logger.log('Error sending watchlist alert email: ' + error);
  }
}

// ============================================
// MICROSOFT TEAMS ALERT (optional -- no-ops until TEAMS_WEBHOOK_URL is set)
// ============================================
function sendTeamsAlert(data) {
  if (!CONFIG.TEAMS_WEBHOOK_URL) return;

  let text;
  if (data.alertType === 'watchlist') {
    text = '⚠️ **Watchlist match at reception**: ' + data.visitorName +
      ' (' + (data.companyName || 'no company given') + ') has just signed in. ' +
      (data.flaggedReason ? 'Reason: ' + data.flaggedReason : '');
  } else if (data.alertType === 'delivery') {
    text = '📦 **Delivery logged**: ' + data.visitorName + ' at ' + data.timeIn + '.' +
      (data.vehicleReg ? ' Notes: ' + data.vehicleReg : '');
  } else {
    text = '🔔 **Visitor arrived**: ' + data.visitorName + ' (' + data.companyName + ') is here to see **' +
      data.hostName + '** — ' + data.visitPurpose + '.';
  }

  try {
    UrlFetchApp.fetch(CONFIG.TEAMS_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: text }),
      muteHttpExceptions: true
    });
  } catch (error) {
    Logger.log('Teams alert failed: ' + error);
  }
}

// ============================================
// SMS ALERT VIA TWILIO (optional -- no-ops until all three Twilio fields,
// plus a number for this host in HOST_SMS_NUMBERS, are filled in)
// ============================================
function sendSmsAlert(data) {
  if (!CONFIG.TWILIO_ACCOUNT_SID || !CONFIG.TWILIO_AUTH_TOKEN || !CONFIG.TWILIO_FROM_NUMBER) return;
  const toNumber = CONFIG.HOST_SMS_NUMBERS[data.hostEmail];
  if (!toNumber) return;

  const body = data.alertType === 'watchlist'
    ? 'Reception alert: watchlist match ' + data.visitorName + ' has signed in.'
    : 'Visitor arrived: ' + data.visitorName + ' (' + data.companyName + ') for ' + data.visitPurpose + '. Please head to reception.';

  const url = 'https://api.twilio.com/2010-04-01/Accounts/' + CONFIG.TWILIO_ACCOUNT_SID + '/Messages.json';
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      headers: { Authorization: 'Basic ' + Utilities.base64Encode(CONFIG.TWILIO_ACCOUNT_SID + ':' + CONFIG.TWILIO_AUTH_TOKEN) },
      payload: { To: toNumber, From: CONFIG.TWILIO_FROM_NUMBER, Body: body },
      muteHttpExceptions: true
    });
    Logger.log('SMS alert sent to: ' + toNumber);
  } catch (error) {
    Logger.log('SMS alert failed: ' + error);
  }
}
