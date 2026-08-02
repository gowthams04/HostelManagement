function doGet(e) {
  var output = ContentService.createTextOutput('Hostel API is running. Use POST requests.');
  return output;
}

function doPost(e) {
  var output;
  try {
    // The frontend sends Content-Type: text/plain with a JSON string body.
    // This avoids CORS preflight (which Apps Script can't answer properly)
    // while still letting us send structured data. Parse it from postData.
    var body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }

    var func = body.func;
    var spreadsheetId = body.spreadsheetId;

    // Any other keys sent alongside func/spreadsheetId
    var params = {};
    for (var key in body) {
      if (key !== 'func' && key !== 'spreadsheetId') {
        params[key] = body[key];
      }
    }

    var result;
    switch (func) {
      case 'testConnection':
        result = { success: true };
        break;
      case 'getDashboardStats':
        result = getDashboardStats(spreadsheetId);
        break;
      case 'getHostels':
        result = getSheetData(spreadsheetId, 'Hostels');
        break;
      case 'getRooms':
        result = getSheetData(spreadsheetId, 'Rooms');
        break;
      case 'getStudents':
        result = getSheetData(spreadsheetId, 'Students');
        break;
      case 'addHostel':
        result = addHostel(spreadsheetId, params.data);
        break;
      case 'addHostels':
        result = addHostels(spreadsheetId, params.data);
        break;
      default:
        throw new Error('Unknown function: ' + func);
    }

    output = ContentService.createTextOutput(JSON.stringify({ result: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    output = ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Note: no CORS headers are set here on purpose.
  // ContentService.TextOutput has no setHeader()/setHeaders() method in Apps
  // Script, and calling one would throw a runtime error. Apps Script Web Apps
  // deployed with "Anyone" access already allow cross-origin GET/POST requests
  // as long as the request doesn't trigger a CORS preflight (see the frontend
  // fix: use Content-Type: text/plain instead of application/json).
  return output;
}

// Helper functions remain the same
function getSheetData(id, name) {
  var sheet = SpreadsheetApp.openById(id).getSheetByName(name);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data.shift();
  return data.map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

// ===== HOSTEL CREATION (manual "type it in" form + Excel bulk upload) =====

// Adds a single hostel. `data` is a plain object, e.g.
// { name, type, floors, address, chiefWarden, phone, status }
function addHostel(spreadsheetId, data) {
  if (!data) throw new Error('No hostel data received.');
  validateHostel(data);
  var sheet = getOrCreateHostelsSheet(spreadsheetId);
  appendRowByHeaders(sheet, data);
  return { success: true };
}

// Adds many hostels at once (used by the Excel upload flow). Invalid rows
// are skipped rather than aborting the whole batch, so one bad row in a
// spreadsheet doesn't block the rest.
function addHostels(spreadsheetId, rows) {
  if (!rows || !rows.length) throw new Error('No rows to upload.');
  var sheet = getOrCreateHostelsSheet(spreadsheetId);
  var added = 0, skipped = 0, errors = [];
  rows.forEach(function(row, i) {
    try {
      validateHostel(row);
      appendRowByHeaders(sheet, row);
      added++;
    } catch (err) {
      skipped++;
      errors.push('Row ' + (i + 1) + ': ' + err.message);
    }
  });
  return { added: added, skipped: skipped, errors: errors };
}

function validateHostel(data) {
  if (!data.name || !String(data.name).trim()) {
    throw new Error('Hostel Name is required.');
  }
  if (data.type && data.type !== 'Boys Hostel' && data.type !== 'Girls Hostel') {
    throw new Error('Hostel Type must be "Boys Hostel" or "Girls Hostel".');
  }
}

// Gets the Hostels sheet, creating it with default headers if it
// doesn't exist yet.
function getOrCreateHostelsSheet(spreadsheetId) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName('Hostels');
  if (!sheet) {
    sheet = ss.insertSheet('Hostels');
    sheet.appendRow(['id', 'name', 'type', 'floors', 'address', 'chiefWarden', 'phone', 'status']);
  }
  return sheet;
}

// Appends a row, matching each column to the sheet's existing header row
// (by name) so column order in the sheet doesn't have to match the object.
// An 'id' column, if present and not supplied, is auto-generated.
function appendRowByHeaders(sheet, data) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];

  if (!headers.length || !headers[0]) {
    headers = ['id', 'name', 'type', 'floors', 'address', 'chiefWarden', 'phone', 'status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  var row = headers.map(function(h, colIndex) {
    if (h === 'id' && (data.id === undefined || data.id === '')) {
      return generateNextId(sheet, colIndex + 1);
    }
    return data[h] !== undefined ? data[h] : '';
  });
  sheet.appendRow(row);
}

// Simple auto-incrementing ID like H001, H002, ... based on the last row's ID.
function generateNextId(sheet, idColIndex) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'H001';
  var lastId = String(sheet.getRange(lastRow, idColIndex).getValue());
  var num = parseInt(lastId.replace(/\D/g, ''), 10);
  if (isNaN(num)) num = lastRow - 1;
  var next = num + 1;
  return 'H' + (next < 100 ? ('00' + next).slice(-3) : next);
}

function getDashboardStats(id) {
  return {
    totalHostels: getSheetData(id, 'Hostels').length,
    totalRooms: getSheetData(id, 'Rooms').length,
    totalStudents: getSheetData(id, 'Students').length,
    totalGuestBookings: getSheetData(id, 'GuestBookings').length
  };
}
