// ════════════════════════════════════════════════════════
//  EliteWriters — Google Apps Script Backend
//  Tasks: 1) Log to Sheet  2) Email client receipt  3) Alert tawk.to
// ════════════════════════════════════════════════════════

const SHEET_NAME    = "Orders";
const TAWKTO_EMAIL  = "tickets@elitewriters-oefbaw.p.tawk.email";
const BUSINESS_EMAIL = "elitewriters804@gmail.com";
const BUSINESS_NAME  = "EliteWriters";

// ── Handle POST from website ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 1. Log to Google Sheet
    const ticketId = logToSheet(data);

    // 2. Send receipt email to client (if email provided)
    if (data.clientEmail) {
      sendClientReceipt(data, ticketId);
    }

    // 3. Alert tawk.to dashboard (creates a support ticket)
    alertTawkTo(data, ticketId, data.clientEmail || "");

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, ticketId: ticketId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Health check ──
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "EliteWriters API is live ✅" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════
//  TASK 1: Log Order to Google Sheet
// ════════════════════════════════════════════════════════
function logToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Create sheet with headers if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Ticket #", "Date Submitted", "Client Email", "Service",
      "Topic", "Paper Type", "Education Level", "Pages",
      "Word Count", "Sources", "Citation Style", "Language",
      "Deadline", "Add-ons", "Files Attached",
      "Instructions", "Estimated Price", "Status"
    ]);
    // Style the header row
    const header = sheet.getRange(1, 1, 1, 18);
    header.setBackground("#0a3d2e");
    header.setFontColor("#ffffff");
    header.setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  // Use ticket ID from frontend (already generated) or generate one
  const lastRow = sheet.getLastRow();
  const ticketNum = lastRow < 2 ? 1 : lastRow; // row 1 is header
  const ticketId = data.ticketId || ("EW-" + String(ticketNum).padStart(3, "0"));

  sheet.appendRow([
    ticketId,
    new Date().toLocaleString("en-GB", { timeZone: "Africa/Nairobi" }),
    data.clientEmail    || "Not provided",
    data.service        || "",
    data.topic          || "",
    data.paperType      || "",
    data.educationLevel || "",
    data.pages          || "",
    data.wordCount      || "",
    data.sources        || "",
    data.citationStyle  || "",
    data.language       || "",
    data.deadline       || "",
    data.addons         || "",
    data.filesAttached  || "None",
    data.instructions   || "",
    data.estimatedPrice || "$0.00",
    "New ⏳"
  ]);

  return ticketId;
}

// ════════════════════════════════════════════════════════
//  TASK 2: Send HTML Receipt Email to Client
// ════════════════════════════════════════════════════════
function sendClientReceipt(data, ticketId) {
  const subject = `✅ Order Confirmed — Ticket ${ticketId} | EliteWriters`;

  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <style>
    body { margin:0; padding:0; background:#f4f4f4; font-family: 'Helvetica Neue', Arial, sans-serif; }
    .wrapper { max-width:600px; margin:0 auto; background:#ffffff; }
    .header { background:#0a3d2e; padding:32px 40px; text-align:center; }
    .header h1 { color:#ffffff; margin:0; font-size:24px; letter-spacing:0.5px; }
    .header p { color:rgba(255,255,255,0.7); margin:6px 0 0; font-size:14px; }
    .ticket-box { background:#0a3d2e; border-radius:10px; padding:20px 28px; margin:28px 40px; text-align:center; }
    .ticket-label { font-size:11px; color:rgba(255,255,255,0.6); text-transform:uppercase; letter-spacing:0.12em; margin-bottom:6px; }
    .ticket-num { font-family:monospace; font-size:28px; font-weight:700; color:#fbbf24; letter-spacing:0.06em; }
    .ticket-hint { font-size:11px; color:rgba(255,255,255,0.5); margin-top:6px; }
    .content { padding:28px 40px; }
    .greeting { font-size:16px; color:#1a1a1a; margin-bottom:16px; }
    .section-title { font-size:11px; font-weight:700; color:#0a3d2e; text-transform:uppercase; letter-spacing:0.1em; border-bottom:2px solid #e5e7eb; padding-bottom:8px; margin:24px 0 16px; }
    .detail-row { display:flex; justify-content:space-between; padding:8px 0; border-bottom:1px solid #f3f4f6; font-size:14px; }
    .detail-label { color:#6b7280; }
    .detail-value { color:#1a1a1a; font-weight:500; text-align:right; max-width:60%; }
    .price-box { background:#f0fdf4; border:1px solid #bbf7d0; border-radius:8px; padding:16px 20px; margin:20px 0; display:flex; justify-content:space-between; align-items:center; }
    .price-label { font-size:14px; color:#166534; font-weight:600; }
    .price-amount { font-size:24px; font-weight:700; color:#0a3d2e; }
    .free-tag { display:inline-block; background:#dcfce7; color:#166534; font-size:11px; font-weight:700; padding:3px 10px; border-radius:100px; }
    .next-steps { background:#f9fafb; border-radius:8px; padding:20px; margin:20px 0; }
    .step-item { display:flex; align-items:flex-start; gap:10px; margin-bottom:12px; font-size:14px; color:#374151; }
    .step-item:last-child { margin-bottom:0; }
    .cta-btn { display:block; width:fit-content; margin:24px auto; background:#0a3d2e; color:#ffffff; text-decoration:none; padding:14px 32px; border-radius:8px; font-weight:600; font-size:15px; text-align:center; }
    .footer { background:#f9fafb; border-top:1px solid #e5e7eb; padding:24px 40px; text-align:center; }
    .footer p { font-size:12px; color:#9ca3af; margin:4px 0; }
    .footer a { color:#0a3d2e; text-decoration:none; }
  </style>
</head>
<body>
<div class="wrapper">

  <div class="header">
    <h1>EliteWriters</h1>
    <p>Professional Academic & Content Writing</p>
  </div>

  <div class="content">
    <p class="greeting">Hi there,</p>
    <p style="font-size:14px;color:#374151;line-height:1.6;">Thank you for placing your order with <strong>EliteWriters</strong>. Your request has been received and logged in our system. Here are your order details:</p>

    <div class="ticket-box">
      <div class="ticket-label">Your Order Ticket Number</div>
      <div class="ticket-num">${ticketId}</div>
      <div class="ticket-hint">Save this number — quote it when following up</div>
    </div>

    <div class="section-title">Order Details</div>

    <div class="detail-row"><span class="detail-label">Service</span><span class="detail-value">${data.service || '—'}</span></div>
    <div class="detail-row"><span class="detail-label">Topic</span><span class="detail-value">${data.topic || '—'}</span></div>
    <div class="detail-row"><span class="detail-label">Paper Type</span><span class="detail-value">${data.paperType || '—'}</span></div>
    <div class="detail-row"><span class="detail-label">Education Level</span><span class="detail-value">${data.educationLevel || '—'}</span></div>
    <div class="detail-row"><span class="detail-label">Pages</span><span class="detail-value">${data.pages || '—'}</span></div>
    <div class="detail-row"><span class="detail-label">Word Count</span><span class="detail-value">${data.wordCount ? data.wordCount + ' words' : '—'}</span></div>
    <div class="detail-row"><span class="detail-label">Sources Required</span><span class="detail-value">${data.sources || '—'}</span></div>
    <div class="detail-row"><span class="detail-label">Citation Style</span><span class="detail-value">${data.citationStyle || '—'}</span></div>
    <div class="detail-row"><span class="detail-label">Language</span><span class="detail-value">${data.language || '—'}</span></div>
    <div class="detail-row"><span class="detail-label">Deadline</span><span class="detail-value">${data.deadline || '—'}</span></div>
    <div class="detail-row"><span class="detail-label">Files Attached</span><span class="detail-value">${data.filesAttached || 'None'}</span></div>
    <div class="detail-row"><span class="detail-label">Add-ons</span><span class="detail-value">${data.addons || '—'}</span></div>

    <div class="section-title">Pricing</div>

    <div class="price-box">
      <span class="price-label">Estimated Total</span>
      <span class="price-amount">${data.estimatedPrice || '$0.00'}</span>
    </div>
    <p style="font-size:12px;color:#6b7280;margin-top:-8px;">
      ✅ Plagiarism Report &amp; AI Report: <span class="free-tag">INCLUDED FREE</span><br>
      <span style="display:block;margin-top:6px;">Orders under $30: full payment upfront. Orders above $30: 50% upfront, 50% on delivery.</span>
    </p>

    ${data.instructions ? `
    <div class="section-title">Your Instructions</div>
    <p style="font-size:14px;color:#374151;line-height:1.6;background:#f9fafb;padding:14px;border-radius:8px;">${data.instructions}</p>
    ` : ''}

    <div class="section-title">What Happens Next</div>
    <div class="next-steps">
      <div class="step-item">✅ <span>Your order is confirmed and logged in our system</span></div>
      <div class="step-item">🔍 <span>We'll review your requirements within <strong>1 hour</strong></span></div>
      <div class="step-item">💬 <span>We'll reach out to confirm and discuss payment</span></div>
      <div class="step-item">✍️ <span>Writing begins immediately after payment</span></div>
      <div class="step-item">🔄 <span>2 free revisions included on every order</span></div>
    </div>

    <a href="https://wa.me/16624002088" class="cta-btn">💬 Chat with Us on WhatsApp →</a>
  </div>

  <div class="footer">
    <p><strong>EliteWriters</strong> — Professional Writing Services</p>
    <p><a href="mailto:elitewriters804@gmail.com">elitewriters804@gmail.com</a> · <a href="https://wa.me/16624002088">+1 (662) 400-2088</a></p>
    <p style="margin-top:10px;">Fast · Reliable · Confidential · 100% Original</p>
  </div>

</div>
</body>
</html>
  `;

  MailApp.sendEmail({
    to: data.clientEmail,
    subject: subject,
    htmlBody: htmlBody,
    name: BUSINESS_NAME,
    replyTo: BUSINESS_EMAIL
  });
}

// ════════════════════════════════════════════════════════
//  TASK 3: Alert tawk.to — Creates a Support Ticket
// ════════════════════════════════════════════════════════
function alertTawkTo(data, ticketId, clientEmail) {
  const subject = `[${ticketId}] New Order: ${data.topic || 'Untitled'} — ${data.service || ''}`;

  const body = `
NEW ORDER RECEIVED — EliteWriters
══════════════════════════════════

Ticket:     ${ticketId}
Submitted:  ${new Date().toLocaleString("en-GB", { timeZone: "Africa/Nairobi" })}
Client:     ${clientEmail || "Not provided"}

─── ORDER DETAILS ───────────────
Service:          ${data.service || '—'}
Topic:            ${data.topic || '—'}
Paper Type:       ${data.paperType || '—'}
Education Level:  ${data.educationLevel || '—'}
Pages:            ${data.pages || '—'}
Word Count:       ${data.wordCount || '—'}
Sources:          ${data.sources || '—'}
Citation Style:   ${data.citationStyle || '—'}
Language:         ${data.language || '—'}
Deadline:         ${data.deadline || '—'}
Files Attached:   ${data.filesAttached || 'None'}
Add-ons:          ${data.addons || '—'}

─── PRICE ───────────────────────
Estimated Total:  ${data.estimatedPrice || '$0.00'}
Plagiarism & AI:  FREE (included)

─── INSTRUCTIONS ─────────────────
${data.instructions || 'None provided'}

══════════════════════════════════
Reply to this email to contact the client directly.
  `;

  // Set replyTo to client's email so tawk.to replies go straight to them
  MailApp.sendEmail({
    to: TAWKTO_EMAIL,
    subject: subject,
    body: body,
    name: "EliteWriters Orders",
    replyTo: clientEmail || BUSINESS_EMAIL
  });
}
