/**
 * Hostel Management System - Apps Script Backend
 * Handles: Hostels, Rooms, Students, Staff/Wardens (read + create, single or bulk)
 */

function doGet(e) {
  return ContentService.createTextOutput('Hostel API is running. Use POST requests.');
}

function doPost(e) {
  var output;
  try {
    var body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }

    var func = body.func;
    var spreadsheetId = body.spreadsheetId;

    var params = {};
    for (var key in body) {
      if (key !== 'func' && key !== 'spreadsheetId') params[key] = body[key];
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
      case 'getStaff':
        result = getSheetData(spreadsheetId, 'Staff');
        break;

      case 'addHostel':
        result = addHostel(spreadsheetId, params.data);
        break;
      case 'addRoom':
        result = addRoom(spreadsheetId, params.data);
        break;
      case 'addStudent':
        result = addStudent(spreadsheetId, params.data);
        break;
      case 'addStaff':
        result = addStaff(spreadsheetId, params.data);
        break;

      case 'addHostels':
        result = addHostels(spreadsheetId, params.data);
        break;
      case 'addRooms':
        result = addRooms(spreadsheetId, params.data);
        break;
      case 'addStudents':
        result = addStudents(spreadsheetId, params.data);
        break;
      case 'addStaffs':
        result = addStaffs(spreadsheetId, params.data);
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

  // No CORS headers here on purpose: ContentService.TextOutput has no
  // setHeader()/setHeaders() method (calling one throws). Apps Script Web
  // Apps deployed with "Anyone" access already allow cross-origin GET/POST
  // as long as the request doesn't trigger a preflight (frontend uses
  // Content-Type: text/plain for that reason).
  return output;
}

// ===================================================================
// READ HELPERS
// ===================================================================

function getSheetData(id, name) {
  var sheet = SpreadsheetApp.openById(id).getSheetByName(name);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data.shift();
  return data.map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function getDashboardStats(id) {
  var hostels = getSheetData(id, 'Hostels');
  var rooms = getSheetData(id, 'Rooms');
  var students = getSheetData(id, 'Students');
  var staff = getSheetData(id, 'Staff');
  var bookings = getSheetData(id, 'GuestBookings');

  // Distinct hostel+floor combinations found across Rooms, as a stand-in
  // "total floors" count (there's no separate Floor master sheet yet).
  var floorSet = {};
  rooms.forEach(function (r) {
    if (r.floor !== undefined && r.floor !== '') {
      floorSet[String(r.hostel || '') + '||' + String(r.floor)] = true;
    }
  });

  function roomBreakdown(category) {
    var subset = rooms.filter(function (r) { return String(r.category || '') === category; });
    var byStatus = function (status) {
      return subset.filter(function (r) { return String(r.status || '').toLowerCase() === status; }).length;
    };
    return {
      total: subset.length,
      available: byStatus('available'),
      occupied: byStatus('occupied'),
      maintenance: byStatus('maintenance')
    };
  }

  var byGender = function (gender) {
    return students.filter(function (s) { return String(s.gender || '').toLowerCase() === gender; }).length;
  };

  return {
    totalHostels: hostels.length,
    totalFloors: Object.keys(floorSet).length,
    totalRooms: rooms.length,
    studentRooms: roomBreakdown('Student Room'),
    guestRooms: roomBreakdown('Guest Room'),
    totalStudents: students.length,
    maleStudents: byGender('male'),
    femaleStudents: byGender('female'),
    totalStaff: staff.length,
    totalGuestBookings: bookings.length
  };
}

// ===================================================================
// GENERIC CREATE ENGINE (used by Hostels / Rooms / Students / Staff)
// ===================================================================

var SHEET_HEADERS = {
  Hostels: ['id', 'name', 'type', 'floors', 'address', 'chiefWarden', 'phone', 'status'],
  Rooms: ['id', 'hostel', 'floor', 'roomNumber', 'category', 'type', 'capacity', 'bathroom', 'ac', 'status', 'remarks'],
  Students: ['id', 'rollNumber', 'name', 'gender', 'department', 'year', 'phone', 'parentPhone', 'email', 'bloodGroup', 'address', 'status'],
  Staff: ['id', 'facultyName', 'employeeId', 'phone', 'email', 'hostel', 'floor', 'status']
};

var ID_PREFIXES = { Hostels: 'H', Rooms: 'R', Students: 'S', Staff: 'W' };

function addRecord(spreadsheetId, sheetName, data, validateFn) {
  if (!data) throw new Error('No data received.');
  validateFn(data, spreadsheetId);
  var sheet = getOrCreateSheet(spreadsheetId, sheetName);
  appendRowByHeaders(sheet, data, ID_PREFIXES[sheetName]);
  return { success: true };
}

function addRecords(spreadsheetId, sheetName, rows, validateFn) {
  if (!rows || !rows.length) throw new Error('No rows to upload.');
  var sheet = getOrCreateSheet(spreadsheetId, sheetName);
  var added = 0, skipped = 0, errors = [];
  rows.forEach(function (row, i) {
    try {
      validateFn(row, spreadsheetId);
      appendRowByHeaders(sheet, row, ID_PREFIXES[sheetName]);
      added++;
    } catch (err) {
      skipped++;
      errors.push('Row ' + (i + 1) + ': ' + err.message);
    }
  });
  return { added: added, skipped: skipped, errors: errors };
}

function getOrCreateSheet(spreadsheetId, sheetName) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(SHEET_HEADERS[sheetName]);
  }
  return sheet;
}

function appendRowByHeaders(sheet, data, idPrefix) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];

  if (!headers.length || !headers[0]) {
    headers = Object.keys(data);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  var row = headers.map(function (h, colIndex) {
    if (h === 'id' && (data.id === undefined || data.id === '')) {
      return generateNextId(sheet, colIndex + 1, idPrefix || 'X');
    }
    return data[h] !== undefined ? data[h] : '';
  });
  sheet.appendRow(row);
}

function generateNextId(sheet, idColIndex, prefix) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return prefix + '001';
  var lastId = String(sheet.getRange(lastRow, idColIndex).getValue());
  var num = parseInt(lastId.replace(/\D/g, ''), 10);
  if (isNaN(num)) num = lastRow - 1;
  var next = num + 1;
  return prefix + (next < 100 ? ('00' + next).slice(-3) : String(next));
}

// ===================================================================
// ENTITY-SPECIFIC WRAPPERS + VALIDATION
// ===================================================================

function addHostel(spreadsheetId, data) { return addRecord(spreadsheetId, 'Hostels', data, validateHostel); }
function addHostels(spreadsheetId, rows) { return addRecords(spreadsheetId, 'Hostels', rows, validateHostel); }
function validateHostel(data) {
  if (!data.name || !String(data.name).trim()) throw new Error('Hostel Name is required.');
  if (data.type && data.type !== 'Boys Hostel' && data.type !== 'Girls Hostel') {
    throw new Error('Hostel Type must be "Boys Hostel" or "Girls Hostel".');
  }
}

function addRoom(spreadsheetId, data) { return addRecord(spreadsheetId, 'Rooms', data, validateRoom); }
function addRooms(spreadsheetId, rows) { return addRecords(spreadsheetId, 'Rooms', rows, validateRoom); }
function validateRoom(data, spreadsheetId) {
  if (!data.roomNumber || !String(data.roomNumber).trim()) throw new Error('Room Number is required.');
  if (!data.hostel || !String(data.hostel).trim()) throw new Error('Hostel is required.');
  var validCategories = ['Student Room', 'Guest Room'];
  if (data.category && validCategories.indexOf(data.category) === -1) {
    throw new Error('Room Category must be "Student Room" or "Guest Room".');
  }
  var validTypes = ['Single', 'Double Sharing', 'Triple Sharing', 'Four Sharing', 'Five Sharing', 'Six Sharing'];
  if (data.type && validTypes.indexOf(data.type) === -1) {
    throw new Error('Invalid Room Type: ' + data.type);
  }
  if (data.capacity !== undefined && data.capacity !== '' && isNaN(Number(data.capacity))) {
    throw new Error('Capacity must be a number.');
  }
  // A room can only be saved for a hostel that already exists — this is
  // what stops rooms from being created against a typo'd or missing hostel.
  if (spreadsheetId && !hostelExists(spreadsheetId, data.hostel)) {
    throw new Error('Hostel "' + data.hostel + '" was not found. Please add this hostel first, then add the room.');
  }
}

// Checks whether a hostel with this name already exists (case-insensitive,
// trimmed) in the Hostels sheet.
function hostelExists(spreadsheetId, hostelName) {
  var target = String(hostelName || '').trim().toLowerCase();
  if (!target) return false;
  var hostels = getSheetData(spreadsheetId, 'Hostels');
  return hostels.some(function (h) {
    return String(h.name || '').trim().toLowerCase() === target;
  });
}

function addStudent(spreadsheetId, data) { return addRecord(spreadsheetId, 'Students', data, validateStudent); }
function addStudents(spreadsheetId, rows) { return addRecords(spreadsheetId, 'Students', rows, validateStudent); }
function validateStudent(data) {
  if (!data.name || !String(data.name).trim()) throw new Error('Student Name is required.');
  if (!data.rollNumber || !String(data.rollNumber).trim()) throw new Error('Roll Number is required.');
  if (data.gender && data.gender !== 'Male' && data.gender !== 'Female') {
    throw new Error('Gender must be Male or Female.');
  }
}

function addStaff(spreadsheetId, data) { return addRecord(spreadsheetId, 'Staff', data, validateStaff); }
function addStaffs(spreadsheetId, rows) { return addRecords(spreadsheetId, 'Staff', rows, validateStaff); }
function validateStaff(data, spreadsheetId) {
  if (!data.facultyName || !String(data.facultyName).trim()) throw new Error('Faculty Name is required.');
  // Hostel is optional for staff, but if one is given it must be a real,
  // already-saved hostel (same rule as Rooms).
  if (spreadsheetId && data.hostel && !hostelExists(spreadsheetId, data.hostel)) {
    throw new Error('Hostel "' + data.hostel + '" was not found. Please add this hostel first.');
  }
}
