// ============================================================
//  CUPOD APPS SCRIPT — v5
//  Supports: addStudent, lookupStudent, recordPayment
// ============================================================

var SHEET_NAME = 'CUPOD Student Tracker';

function doGet(e) {
  try {
    if (!e || !e.parameter || !e.parameter.action) {
      return respond({ status: 'ok', message: 'CUPOD script is live!' });
    }
    var a = e.parameter.action;
    if (a === 'addStudent')    return respond(addStudent(e.parameter));
    if (a === 'lookupStudent') return respond(lookupStudent(e.parameter));
    if (a === 'recordPayment') return respond(recordPayment(e.parameter));
    return respond({ status: 'error', message: 'Unknown action: ' + a });
  } catch (err) {
    Logger.log('ERROR: ' + err.toString());
    return respond({ status: 'error', message: err.toString() });
  }
}

function doPost(e) { return doGet(e); }

// ============================================================
//  ADD STUDENT
// ============================================================
function addStudent(p) {
  if (!p.studentName || p.studentName.trim() === '')
    return { status: 'error', message: 'Student name is required' };
  if (!p.phone || p.phone.trim() === '')
    return { status: 'error', message: 'Phone is required' };
  if (!p.enrollDate)
    return { status: 'error', message: 'Enrollment date is required' };

  var sheet = getSheet();
  if (sheet.error) return sheet;

  var lastRow  = sheet.getLastRow();
  var nextRow  = lastRow + 1;
  var studentId    = 'CU' + pad(nextRow - 1, 5);
  var coursePrice  = parseFloat(p.coursePrice)  || 0;
  var firstPayment = parseFloat(p.firstPayment) || 0;
  var accessDays   = parseInt(p.accessDays)     || 90;
  var enrollDate   = parseDate(p.enrollDate);
  var r = nextRow;

  var rowData = [
    studentId,
    p.studentName.trim(),
    (p.email || '').trim(),
    p.phone.trim(),
    coursePrice,
    firstPayment > 0 ? firstPayment : '',
    firstPayment > 0 ? enrollDate   : '',
    '', '', '', '',
    '', '',
    enrollDate,
    '', '', ''
  ];

  sheet.getRange(r, 1, 1, rowData.length).setValues([rowData]);
  sheet.getRange(r, 12).setFormula('=SUM(F'+r+',H'+r+',J'+r+')');
  sheet.getRange(r, 13).setFormula('=IF(E'+r+'<>"",E'+r+'-L'+r+',"")');
  sheet.getRange(r, 15).setFormula('=IF(N'+r+'<>"",N'+r+'+'+accessDays+',"")');
  sheet.getRange(r, 16).setFormula('=IF(O'+r+'="","",IF(TODAY()>O'+r+',"Expired",IF(TODAY()>=O'+r+'-10,"Expiring Soon","Active")))');
  sheet.getRange(r, 17).setFormula('=IF(P'+r+'="Expired","Stop access!",IF(P'+r+'="Expiring Soon","Renewal/payment due soon","No action needed"))');

  applyFormats(sheet, r);
  Logger.log('Added ' + studentId + ' at row ' + r);
  return { status: 'success', studentId: studentId, rowNumber: r };
}

// ============================================================
//  LOOKUP STUDENT — searches whole sheet, flexible matching
// ============================================================
function lookupStudent(p) {
  if (!p.studentId || p.studentId.trim() === '')
    return { status: 'error', message: 'Student ID is required' };

  var sheet = getSheet();
  if (sheet.error) return sheet;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'error', message: 'No students in the sheet yet.' };

  // Read ALL rows (skip row 1 which is header)
  var allData = sheet.getRange(2, 1, lastRow - 1, 17).getValues();

  var query = p.studentId.trim().toUpperCase().replace(/\s+/g, '');

  // Try to find a match — check column A (Student ID)
  for (var i = 0; i < allData.length; i++) {
    var cellId = String(allData[i][0] || '').trim().toUpperCase().replace(/\s+/g, '');
    if (cellId === '' ) continue;
    if (cellId === query) {
      return buildStudentResult(allData[i], i + 2); // +2: row 1 is header
    }
  }

  // No exact match — try partial match (in case user typed CU45 instead of CU00045)
  for (var j = 0; j < allData.length; j++) {
    var cid = String(allData[j][0] || '').trim().toUpperCase().replace(/\s+/g, '');
    if (cid === '') continue;
    // Strip leading zeros from both and compare
    var stripped1 = cid.replace(/^CU0+/, 'CU');
    var stripped2 = query.replace(/^CU0+/, 'CU');
    if (stripped1 === stripped2) {
      return buildStudentResult(allData[j], j + 2);
    }
  }

  // Still not found — return helpful debug info
  var sample = [];
  for (var k = 0; k < Math.min(5, allData.length); k++) {
    var sid = String(allData[k][0] || '').trim();
    if (sid !== '') sample.push(sid);
  }

  return {
    status: 'error',
    message: 'Student ID "' + p.studentId.trim() + '" not found.\n' +
             'Example IDs in your sheet: ' + (sample.join(', ') || 'none found') + '\n' +
             'Total rows checked: ' + allData.length
  };
}

function buildStudentResult(row, rowNumber) {
  // Count filled payment slots (F=idx5, H=idx7, J=idx9)
  var slots = 0;
  if (row[5] !== '' && row[5] !== null && row[5] !== undefined) slots++;
  if (row[7] !== '' && row[7] !== null && row[7] !== undefined) slots++;
  if (row[9] !== '' && row[9] !== null && row[9] !== undefined) slots++;

  var enrollVal = row[13];
  var enrollStr = '';
  try {
    if (enrollVal && enrollVal !== '') {
      enrollStr = Utilities.formatDate(new Date(enrollVal), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  } catch(e) { enrollStr = String(enrollVal || ''); }

  return {
    status: 'success',
    student: {
      id:                 String(row[0] || '').trim(),
      name:               String(row[1] || '').trim(),
      email:              String(row[2] || '').trim(),
      phone:              String(row[3] || '').trim(),
      coursePrice:        parseFloat(row[4]) || 0,
      totalPaid:          parseFloat(row[11]) || 0,
      balance:            parseFloat(row[12]) || 0,
      enrollDate:         enrollStr,
      status:             String(row[15] || '').trim(),
      paymentSlotsFilled: slots,
      rowNumber:          rowNumber,
    }
  };
}

// ============================================================
//  RECORD PAYMENT
// ============================================================
function recordPayment(p) {
  if (!p.studentId) return { status: 'error', message: 'Student ID is required' };
  if (!p.amount || isNaN(parseFloat(p.amount)) || parseFloat(p.amount) <= 0)
    return { status: 'error', message: 'A valid payment amount is required' };
  if (!p.paymentDate) return { status: 'error', message: 'Payment date is required' };

  var lookup = lookupStudent({ studentId: p.studentId });
  if (lookup.status !== 'success') return lookup;

  var s     = lookup.student;
  var sheet = getSheet();
  if (sheet.error) return sheet;

  var r      = s.rowNumber;
  var slots  = s.paymentSlotsFilled;
  var amount = parseFloat(p.amount);
  var payDate = parseDate(p.paymentDate);

  var amtCol, dateCol, slotNum;
  if      (slots === 0) { amtCol=6;  dateCol=7;  slotNum=1; }
  else if (slots === 1) { amtCol=8;  dateCol=9;  slotNum=2; }
  else if (slots === 2) { amtCol=10; dateCol=11; slotNum=3; }
  else return { status: 'error', message: 'All 3 payment slots are already full for this student.' };

  sheet.getRange(r, amtCol).setValue(amount);
  sheet.getRange(r, dateCol).setValue(payDate);
  sheet.getRange(r, amtCol).setNumberFormat('"$"#,##0.00');
  sheet.getRange(r, dateCol).setNumberFormat('yyyy-mm-dd');

  SpreadsheetApp.flush();
  var newBalance = sheet.getRange(r, 13).getValue();

  Logger.log('Payment recorded: ' + p.studentId + ' slot ' + slotNum + ' $' + amount);
  return {
    status:     'success',
    studentId:  p.studentId,
    slotUsed:   slotNum,
    amount:     amount,
    newBalance: parseFloat(newBalance) || 0,
  };
}

// ============================================================
//  HELPERS
// ============================================================
function getSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    var all = ss.getSheets().map(function(s){ return s.getName(); }).join(', ');
    return { error: true, status: 'error',
             message: 'Sheet "'+SHEET_NAME+'" not found. Available tabs: '+all };
  }
  return sheet;
}

function applyFormats(sheet, r) {
  var money = '"$"#,##0.00';
  [5,6,8,10,12,13].forEach(function(c){ sheet.getRange(r,c).setNumberFormat(money); });
  [7,9,11,14,15].forEach(function(c){ sheet.getRange(r,c).setNumberFormat('yyyy-mm-dd'); });
}

function parseDate(value) {
  if (value instanceof Date) return value;
  var s = String(value).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(parseInt(m[1]), parseInt(m[2])-1, parseInt(m[3]));
  return new Date(s);
}

function pad(num, size) {
  var s = String(num);
  while (s.length < size) s = '0' + s;
  return s;
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  TEST — run these from Apps Script editor to debug
// ============================================================
function testLookup() {
  // This will show you what IDs are actually in your sheet
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  var ids = sheet.getRange(2, 1, Math.min(10, lastRow-1), 2).getValues();
  Logger.log('First 10 rows (ID | Name):');
  ids.forEach(function(r, i){ Logger.log((i+2)+': "'+r[0]+'" | "'+r[1]+'"'); });
}

function testLookupById() {
  // Change CU00001 to a real ID from your sheet
  var r = lookupStudent({ studentId: 'CU00001' });
  Logger.log(JSON.stringify(r, null, 2));
}
