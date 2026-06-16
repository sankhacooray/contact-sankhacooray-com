/**
 * sankhacooray.com — Contact-inquiry form backend.
 *
 * Surface:
 *   - POST /exec    → JSON-in / JSON-out endpoint, called by the portfolio's
 *                     contact form (v3.sankhacooray.com) via fetch() with
 *                     Content-Type: text/plain.
 *   - GET  /exec    → a tiny status page (handy to confirm the deploy is live).
 *
 * What it does on each submit:
 *   1. Appends the inquiry to a Google Sheet (your private inbox of leads).
 *   2. Sends the sender a branded HTML "thanks, I'll be in touch" email.
 *   3. Emails YOU an internal alert with the full message, with reply-to set
 *      to the sender so you can respond directly from your inbox.
 *
 * Wiring (one-time, after you paste this into your own Google account):
 *   1. Apps Script IDE → Run → `setup`  → authorize scopes. This creates a
 *      Google Sheet titled SHEET_TITLE in your Drive and stores its ID in
 *      Script Properties (key: SHEET_ID). Log shows the Sheet URL.
 *   2. Deploy → New deployment → type "Web app", execute as "Me",
 *      access "Anyone". Copy the /exec URL into the website's script.js.
 *
 * Optional Script Properties (Project Settings → Script Properties):
 *   NOTIFY_EMAIL — where to send the "new inquiry" alert (e.g. your inbox).
 *                  If unset, the alert goes to OWNER_EMAIL (defined below).
 *   FROM_EMAIL   — alias to send the auto-reply FROM (e.g. hello@sankhacooray.com).
 *                  Must be a verified "Send mail as" alias in your Gmail.
 *                  If unset/unverified, falls back to your primary address.
 *   FROM_NAME    — sender display name. Defaults to "Sankha Cooray".
 */

const SHEET_NAME  = "Inquiries";
const SHEET_TITLE = "sankhacooray.com — Contact inquiries";
const BRAND       = "Sankha Cooray";
const SITE        = "sankhacooray.com";
// Default inbox for the internal "new inquiry" alert. Override with the
// NOTIFY_EMAIL Script Property without touching code.
const OWNER_EMAIL = "bsc2fast@gmail.com";
const HEADERS = [
  "Timestamp", "Name", "Email", "Message", "User agent", "Source"
];

// ---------------------------------------------------------------------------
// Web app entrypoints
// ---------------------------------------------------------------------------

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: "sankhacooray.com contact endpoint" }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Cross-origin endpoint used by the portfolio's contact form.
 *
 * Browsers treat Content-Type: text/plain as a "simple request" so no CORS
 * preflight is sent. Apps Script automatically attaches
 * Access-Control-Allow-Origin: * for ANYONE_ANONYMOUS deployments, so the
 * response is readable from any origin.
 */
function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) || "{}";
    const payload = JSON.parse(raw);
    const result = submitInquiry(payload);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error("doPost failed:", err);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, message: "Server error: " + err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Business logic.
 *
 * @param {{name:string, email:string, message:string,
 *          company?:string, userAgent?:string, source?:string}} payload
 *          `company` is a hidden honeypot — real users leave it blank.
 * @return {{ok:boolean, message:string}}
 */
function submitInquiry(payload) {
  try {
    const data = sanitise_(payload || {});

    // Honeypot: bots fill the hidden `company` field; humans never see it.
    if (data.company) return { ok: true, message: "Thanks — your message has been sent." };

    const err = validate_(data);
    if (err) return { ok: false, message: err };

    appendToSheet_(data);
    sendThankYouEmail_(data);
    notifyInternal_(data);

    return { ok: true, message: "Thanks — your message landed. I'll get back to you soon." };
  } catch (err) {
    console.error("submitInquiry failed:", err);
    return { ok: false, message: "Sorry, something went wrong on my end. Please try again in a moment." };
  }
}

// ---------------------------------------------------------------------------
// One-time setup (run from the IDE)
// ---------------------------------------------------------------------------

function setup() {
  const props = PropertiesService.getScriptProperties();
  let sheetId = props.getProperty("SHEET_ID");
  let ss;

  if (sheetId) {
    ss = SpreadsheetApp.openById(sheetId);
    console.log("Reusing existing sheet:", ss.getName(), "—", ss.getUrl());
  } else {
    ss = SpreadsheetApp.create(SHEET_TITLE);
    sheetId = ss.getId();
    props.setProperty("SHEET_ID", sheetId);
    console.log("Created new sheet:", ss.getUrl());
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const blank = ss.getSheets().find((s) => s.getName() === "Sheet1");
    if (blank && ss.getSheets().length > 1) ss.deleteSheet(blank);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, HEADERS.length);
  }

  console.log("Setup complete. Sheet URL:", ss.getUrl());
  return ss.getUrl();
}

// ---------------------------------------------------------------------------
// Sheet helpers
// ---------------------------------------------------------------------------

function getSheet_() {
  const sheetId = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
  if (!sheetId) throw new Error("SHEET_ID is not set. Run setup() once from the IDE.");
  return SpreadsheetApp.openById(sheetId).getSheetByName(SHEET_NAME);
}

function appendToSheet_(d) {
  getSheet_().appendRow([
    new Date(),
    d.name,
    d.email,
    d.message,
    d.userAgent,
    d.source,
  ]);
}

// ---------------------------------------------------------------------------
// Email
// ---------------------------------------------------------------------------

function sendThankYouEmail_(d) {
  const firstName = (d.name.split(/\s+/)[0] || "there").trim();
  const subject = "Thanks for reaching out — " + BRAND;

  const body =
    "Hi " + firstName + ",\n\n" +
    "Thanks for getting in touch through " + SITE + ". Your message has reached me and " +
    "I'll get back to you as soon as I can.\n\n" +
    "For reference, here's what you sent:\n\n" +
    d.message + "\n\n" +
    "— " + BRAND + "\n";

  sendBrandedEmail_({
    to: d.email,
    subject: subject,
    body: body,
    htmlBody: renderThankYouHtml_(firstName, d.message),
  });
}

function notifyInternal_(d) {
  // Always alert the owner. NOTIFY_EMAIL (Script Property) overrides the default.
  const to = PropertiesService.getScriptProperties().getProperty("NOTIFY_EMAIL") || OWNER_EMAIL;
  if (!to) return;
  const lines = [
    "New contact inquiry via " + (d.source || SITE),
    "",
    "Name:       " + d.name,
    "Email:      " + d.email,
    "User agent: " + (d.userAgent || "-"),
    "Source:     " + (d.source || "-"),
    "",
    "Message:",
    d.message,
    "",
    "Reply to this email to respond directly to " + d.name + ".",
  ];
  sendBrandedEmail_({
    to: to,
    subject: "[" + (d.source || SITE) + "] New inquiry from " + d.name,
    body: lines.join("\n"),
    htmlBody: renderInternalHtml_(d),
    overrideName: SITE + " contact",
    replyToOverride: d.email, // reply goes straight to the sender
  });
}

/**
 * Sends email using the FROM_EMAIL alias (via GmailApp) when configured,
 * otherwise falls back to MailApp (which always sends as the deployer).
 * Falls back transparently if the alias isn't a verified "Send mail as".
 */
function sendBrandedEmail_(opts) {
  const props = PropertiesService.getScriptProperties();
  const fromEmail = props.getProperty("FROM_EMAIL");
  const senderName = opts.overrideName || props.getProperty("FROM_NAME") || BRAND;
  const replyTo = opts.replyToOverride || fromEmail || undefined;

  if (fromEmail) {
    try {
      GmailApp.sendEmail(opts.to, opts.subject, opts.body, {
        htmlBody: opts.htmlBody,
        name: senderName,
        from: fromEmail,
        replyTo: replyTo,
      });
      return;
    } catch (err) {
      console.warn(
        "Send-as alias '" + fromEmail + "' failed (" + err.message +
        "). Verify it in Gmail → Settings → Accounts → Send mail as. Falling back."
      );
    }
  }

  MailApp.sendEmail({
    to: opts.to,
    subject: opts.subject,
    body: opts.body,
    htmlBody: opts.htmlBody,
    name: senderName,
    replyTo: opts.replyToOverride || undefined,
  });
}

/**
 * Branded HTML auto-reply — table layout + inline styles for cross-client
 * consistency. Dark header band echoes the portfolio's look.
 */
function renderThankYouHtml_(firstName, message) {
  const name = escapeHtml_(firstName);
  const msg  = escapeHtml_(message).replace(/\n/g, "<br>");
  const INK     = "#0e1320";   // dark header / strong text
  const PRIMARY = "#5b8cff";   // accent (blue)
  const ACCENT2 = "#34d6a5";   // accent (mint)
  const SURFACE = "#f4f5f7";   // footer + quote band
  const BG      = "#fafafa";   // outer page bg
  return [
    '<!DOCTYPE html>',
    '<html lang="en"><head>',
    '<meta charset="UTF-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    '<title>Thanks for reaching out</title>',
    '</head>',
    '<body style="margin:0;padding:0;background:' + BG + ';font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Arial,sans-serif;color:' + INK + ';-webkit-font-smoothing:antialiased;">',

    '<div style="display:none;max-height:0;overflow:hidden;mso-hide:all;font-size:1px;line-height:1px;color:' + BG + ';">',
    'Got your message — I\'ll be in touch soon.',
    '</div>',

    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:' + BG + ';">',
    '<tr><td align="center" style="padding:40px 16px;">',

    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:560px;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(14,19,32,0.08);">',

    // Header band
    '<tr><td style="background:linear-gradient(135deg,' + INK + ' 0%,' + PRIMARY + ' 100%);background-color:' + INK + ';padding:38px 40px 34px;">',
    '<div style="font-size:13px;font-weight:600;letter-spacing:0.18em;text-transform:uppercase;color:rgba(255,255,255,0.7);margin-bottom:12px;">' + BRAND + '</div>',
    '<div style="font-size:26px;font-weight:700;color:#ffffff;line-height:1.15;letter-spacing:-0.01em;">Thanks for reaching out</div>',
    '</td></tr>',

    // Body
    '<tr><td style="padding:34px 40px 8px;">',
    '<p style="margin:0 0 18px;font-size:16px;line-height:1.6;">Hi <strong>' + name + '</strong>,</p>',
    '<p style="margin:0 0 18px;font-size:16px;line-height:1.6;">',
    'Thanks for getting in touch through <strong style="color:' + PRIMARY + ';">' + SITE + '</strong>. ',
    'Your message has reached me and I\'ll get back to you as soon as I can.',
    '</p>',
    '<p style="margin:0 0 10px;font-size:13px;font-weight:600;letter-spacing:0.04em;text-transform:uppercase;color:rgba(14,19,32,0.5);">Your message</p>',
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:' + SURFACE + ';border-left:3px solid ' + ACCENT2 + ';border-radius:8px;">',
    '<tr><td style="padding:16px 18px;font-size:15px;line-height:1.6;color:rgba(14,19,32,0.85);">' + msg + '</td></tr>',
    '</table>',

    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:24px;border-top:1px solid ' + SURFACE + ';">',
    '<tr><td style="padding-top:22px;padding-bottom:6px;">',
    '<p style="margin:0;font-size:15px;line-height:1.5;">— <strong>' + BRAND + '</strong></p>',
    '</td></tr>',
    '</table>',

    '</td></tr>',

    // Footer
    '<tr><td style="padding:22px 40px;background:' + SURFACE + ';font-size:12px;line-height:1.55;color:rgba(14,19,32,0.55);">',
    'You\'re receiving this because you sent a message through <strong style="color:rgba(14,19,32,0.72);">' + SITE + '</strong>. ',
    'If that wasn\'t you, you can ignore this email.',
    '</td></tr>',

    '</table>',
    '</td></tr></table>',
    '</body></html>',
  ].join('');
}

/**
 * Branded HTML for the internal "new inquiry" alert sent to the owner.
 * Same visual language as the auto-reply: dark header band, message quote
 * block, plus the sender's details and a one-tap reply button.
 */
function renderInternalHtml_(d) {
  const name   = escapeHtml_(d.name);
  const email  = escapeHtml_(d.email);
  const source = escapeHtml_(d.source || SITE);
  const ua     = escapeHtml_(d.userAgent || "—");
  const msg    = escapeHtml_(d.message).replace(/\n/g, "<br>");
  let received = "";
  try {
    received = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Etc/UTC", "d MMM yyyy, h:mm a");
  } catch (e) { received = ""; }

  const INK     = "#0e1320";   // dark header / strong text
  const PRIMARY = "#5b8cff";   // accent (blue)
  const ACCENT2 = "#34d6a5";   // accent (mint)
  const SURFACE = "#f4f5f7";   // footer + quote band
  const BG      = "#fafafa";   // outer page bg
  const MUTED   = "rgba(14,19,32,0.5)";

  // One label/value row in the details table.
  const row = (label, value) =>
    '<tr>' +
    '<td style="padding:7px 0;font-size:13px;font-weight:600;letter-spacing:0.03em;text-transform:uppercase;color:' + MUTED + ';width:84px;vertical-align:top;">' + label + '</td>' +
    '<td style="padding:7px 0;font-size:15px;line-height:1.5;color:' + INK + ';">' + value + '</td>' +
    '</tr>';

  const replySubject = encodeURIComponent("Re: your message via " + (d.source || SITE));

  return [
    '<!DOCTYPE html>',
    '<html lang="en"><head>',
    '<meta charset="UTF-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    '<title>New contact inquiry</title>',
    '</head>',
    '<body style="margin:0;padding:0;background:' + BG + ';font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Arial,sans-serif;color:' + INK + ';-webkit-font-smoothing:antialiased;">',

    // Hidden preheader (preview text in the inbox list)
    '<div style="display:none;max-height:0;overflow:hidden;mso-hide:all;font-size:1px;line-height:1px;color:' + BG + ';">',
    name + ' sent a message via ' + source,
    '</div>',

    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:' + BG + ';">',
    '<tr><td align="center" style="padding:40px 16px;">',

    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:560px;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(14,19,32,0.08);">',

    // Header band
    '<tr><td style="background:linear-gradient(135deg,' + INK + ' 0%,' + PRIMARY + ' 100%);background-color:' + INK + ';padding:32px 40px 28px;">',
    '<div style="font-size:13px;font-weight:600;letter-spacing:0.18em;text-transform:uppercase;color:rgba(255,255,255,0.7);margin-bottom:10px;">New inquiry</div>',
    '<div style="font-size:25px;font-weight:700;color:#ffffff;line-height:1.15;letter-spacing:-0.01em;">' + name + '</div>',
    '<div style="font-size:13px;color:rgba(255,255,255,0.66);margin-top:6px;">via ' + source + (received ? ' &middot; ' + received : '') + '</div>',
    '</td></tr>',

    // Sender details
    '<tr><td style="padding:28px 40px 4px;">',
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">',
    row("From", name),
    row("Email", '<a href="mailto:' + email + '" style="color:' + PRIMARY + ';text-decoration:none;">' + email + '</a>'),
    row("Source", source),
    '</table>',

    // Message
    '<p style="margin:22px 0 10px;font-size:13px;font-weight:600;letter-spacing:0.04em;text-transform:uppercase;color:' + MUTED + ';">Message</p>',
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:' + SURFACE + ';border-left:3px solid ' + ACCENT2 + ';border-radius:8px;">',
    '<tr><td style="padding:16px 18px;font-size:15px;line-height:1.6;color:rgba(14,19,32,0.85);">' + msg + '</td></tr>',
    '</table>',

    // Reply button
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:26px 0 4px;"><tr>',
    '<td style="border-radius:10px;background:' + PRIMARY + ';">',
    '<a href="mailto:' + email + '?subject=' + replySubject + '" style="display:inline-block;padding:13px 26px;font-size:15px;font-weight:600;color:#ffffff;text-decoration:none;border-radius:10px;">Reply to ' + name + ' &rarr;</a>',
    '</td></tr></table>',
    '<p style="margin:10px 0 0;font-size:13px;line-height:1.5;color:' + MUTED + ';">Or just hit reply &mdash; this email\'s reply-to is already set to the sender.</p>',

    '</td></tr>',

    // Footer — technical context
    '<tr><td style="padding:20px 40px;background:' + SURFACE + ';font-size:12px;line-height:1.55;color:rgba(14,19,32,0.55);">',
    '<strong style="color:rgba(14,19,32,0.7);">User agent</strong><br>' + ua,
    '</td></tr>',

    '</table>',
    '</td></tr></table>',
    '</body></html>',
  ].join('');
}

// ---------------------------------------------------------------------------
// Input handling
// ---------------------------------------------------------------------------

function sanitise_(p) {
  const str = (v) => (v == null ? "" : String(v).trim());
  return {
    name:      str(p.name).slice(0, 200),
    email:     str(p.email).slice(0, 200).toLowerCase(),
    message:   str(p.message).slice(0, 5000),
    company:   str(p.company).slice(0, 200), // honeypot
    userAgent: str(p.userAgent).slice(0, 500),
    source:    str(p.source).slice(0, 200) || SITE + " contact form",
  };
}

function validate_(d) {
  if (!d.name)    return "Please enter your name.";
  if (!d.email)   return "Please enter your email.";
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email)) return "That email doesn't look right — please check it.";
  if (!d.message) return "Please enter a message.";
  if (d.message.length < 5) return "Your message is a little short — add a few more words.";
  return null;
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}
