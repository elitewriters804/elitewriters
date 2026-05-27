// ============================================
// ELITEWRITERS — GOOGLE APPS SCRIPT BACKEND
// Deploy as Web App: Execute as Me, Anyone can access
// ============================================

const SHEET_ID = '1unEDBCYyVN2wA_s0dO-gSvk0NMlkEAbfEtbHtg0EcZA';
const SHEET_NAME = 'Orders';

function doPost(e) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json'
  };

  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'createTicket') {
      const result = createTicket(data);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'getTicket') {
      const result = getTicket(data.ticketNumber);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'getAllTickets') {
      const result = getAllTickets();
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getTicket') {
    const result = getTicket(e.parameter.ticketNumber);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getAllTickets') {
    const result = getAllTickets();
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, message: 'EliteWriters API running' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function createTicket(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Create sheet with headers if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'Ticket #', 'Date', 'Service', 'Topic', 'Type', 'Level',
      'Pages', 'Words', 'Sources', 'Citation', 'Language',
      'Deadline', 'Instructions', 'Status', 'Price', 'Paid'
    ]);
    // Format header row
    sheet.getRange(1, 1, 1, 16).setFontWeight('bold').setBackground('#0a3d2e').setFontColor('#ffffff');
  }

  // Get next ticket number
  const lastRow = sheet.getLastRow();
  const ticketNum = lastRow; // row 1 = headers, row 2 = #001
  const ticketFormatted = '#' + String(ticketNum).padStart(3, '0');

  // Append order row
  sheet.appendRow([
    ticketFormatted,
    new Date().toLocaleString(),
    data.service || '',
    data.topic || '',
    data.type || '',
    data.level || '',
    data.pages || '',
    data.words || '',
    data.sources || '',
    data.citation || '',
    data.language || '',
    data.deadline || '',
    data.instructions || '',
    'Pending',
    '',
    'Unpaid'
  ]);

  // Auto-resize columns
  sheet.autoResizeColumns(1, 16);

  return {
    success: true,
    ticketNumber: ticketFormatted,
    message: 'Ticket created successfully'
  };
}

function getTicket(ticketNumber) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { success: false, error: 'No orders found' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === ticketNumber) {
      return {
        success: true,
        ticket: {
          ticketNumber: data[i][0],
          date: data[i][1],
          service: data[i][2],
          topic: data[i][3],
          type: data[i][4],
          level: data[i][5],
          pages: data[i][6],
          words: data[i][7],
          sources: data[i][8],
          citation: data[i][9],
          language: data[i][10],
          deadline: data[i][11],
          instructions: data[i][12],
          status: data[i][13],
          price: data[i][14],
          paid: data[i][15]
        }
      };
    }
  }
  return { success: false, error: 'Ticket not found' };
}

function getAllTickets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { success: true, tickets: [] };

  const data = sheet.getDataRange().getValues();
  const tickets = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      tickets.push({
        ticketNumber: data[i][0],
        date: data[i][1],
        service: data[i][2],
        topic: data[i][3],
        type: data[i][4],
        level: data[i][5],
        pages: data[i][6],
        deadline: data[i][11],
        status: data[i][13],
        price: data[i][14],
        paid: data[i][15]
      });
    }
  }
  return { success: true, tickets: tickets.reverse() };
}
