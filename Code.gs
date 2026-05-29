var SHEET_ID = '1unEDBCYyVN2wA_s0dO-gSvk0NMlkEAbfEtbHtg0EcZA';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

    // Add headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Ticket ID','Date','Service','Subject','Topic','Paper Type',
        'Level','Pages','Sources','Citation','Language',
        'Deadline','Add-ons','Files','Price','Instructions','Client Email']);
    }

    // Generate sequential ticket ID from sheet row count
    var nextRow = sheet.getLastRow(); // includes header row
    var ticketId = 'EW-' + String(nextRow).padStart(3,'0');

    sheet.appendRow([
      ticketId,
      new Date().toLocaleString(),
      data.service     || '',
      data.subject     || '',
      data.topic       || '',
      data.paperType   || '',
      data.educationLevel || '',
      data.pages       || '',
      data.sources     || '',
      data.citationStyle || '',
      data.language    || '',
      data.deadline    || '',
      data.addons      || '',
      data.filesAttached || '',
      data.estimatedPrice || '',
      data.instructions || '',
      data.clientEmail || ''
    ]);

    // Notify owner
    MailApp.sendEmail({
      to: 'elitewriters804@gmail.com',
      subject: '🎫 New Order: ' + ticketId + ' — ' + (data.topic || ''),
      body: 'New order received!\n\n' +
        'Ticket: '   + ticketId + '\n' +
        'Client: '   + (data.clientEmail || 'Unknown') + '\n' +
        'Service: '  + data.service + '\n' +
        'Subject: '  + data.subject + '\n' +
        'Topic: '    + data.topic + '\n' +
        'Level: '    + data.educationLevel + '\n' +
        'Pages: '    + data.pages + '\n' +
        'Deadline: ' + data.deadline + '\n' +
        'Price: '    + data.estimatedPrice + '\n' +
        'Add-ons: '  + data.addons + '\n' +
        'Instructions: ' + data.instructions
    });

    // Confirm to client
    if (data.clientEmail) {
      MailApp.sendEmail({
        to: data.clientEmail,
        subject: '✅ Order Confirmed — ' + ticketId + ' | EliteWriters',
        body: 'Hi!\n\nYour order has been received successfully.\n\n' +
          '📋 Ticket: '    + ticketId + '\n' +
          '📌 Topic: '     + data.topic + '\n' +
          '📚 Subject: '   + data.subject + '\n' +
          '⏰ Deadline: '  + data.deadline + '\n' +
          '💰 Price: '     + data.estimatedPrice + '\n\n' +
          'We will contact you within 1 hour.\n\n' +
          'WhatsApp: https://wa.me/16624002088\n\n' +
          'Thank you for choosing EliteWriters!'
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', ticketId: ticketId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'EliteWriters API is live ✅', timestamp: new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}

function testEmail() {
  MailApp.sendEmail({
    to: 'elitewriters804@gmail.com',
    subject: '✅ Test Email — EliteWriters Script Working',
    body: 'If you received this, email sending is working correctly!'
  });
  Logger.log('Test email sent!');
}
