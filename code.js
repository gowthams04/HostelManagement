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

function getDashboardStats(id) {
  return {
    totalHostels: getSheetData(id, 'Hostels').length,
    totalRooms: getSheetData(id, 'Rooms').length,
    totalStudents: getSheetData(id, 'Students').length,
    totalGuestBookings: getSheetData(id, 'GuestBookings').length
  };
}