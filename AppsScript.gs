// ============================================================
//  CUPOD APPS SCRIPT — UPDATED VERSION
//  Supports: addStudent, lookupStudent, recordPayment
// ============================================================

const SHEET_NAME = 'CUPOD Student Tracker';

function doGet(e) {
  try {
    if (!e || !e.parameter || !e.parameter.action) {
      return respond({ status: 'ok', message: 'CUPOD script is live!' });
    }
    const a = e.parameter.action;
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
  var studentId = 'CU' + pad(nextRow - 1, 5);
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
//  LOOKUP STUDENT
// ============================================================
function lookupStudent(p) {
  if (!p.studentId) return { status: 'error', message: 'Student ID is required' };

  var sheet  = getSheet();
  if (sheet.error) return sheet;

  var idCol  = sheet.getRange('A2:A' + sheet.getLastRow()).getValues();
  var targetId = p.studentId.trim().toUpperCase();

  for (var i = 0; i < idCol.length; i++) {
    if (String(idCol[i][0]).toUpperCase() === targetId) {
      var r    = i + 2; // +2 because data starts at row 2
      var row  = sheet.getRange(r, 1, 1, 17).getValues()[0];

      // Count filled payment slots
      var slots = 0;
      if (row[5] !== '' && row[5] !== null) slots++;
      if (row[7] !== '' && row[7] !== null) slots++;
      if (row[9] !== '' && row[9] !== null) slots++;

      var totalPaid = parseFloat(row[11]) || 0;
      var balance   = parseFloat(row[12]) || 0;

      return {
        status: 'success',
        student: {
          id:                  String(row[0]),
          name:                String(row[1]),
          email:               String(row[2]),
          phone:               String(row[3]),
          coursePrice:         parseFloat(row[4]) || 0,
          totalPaid:           totalPaid,
          balance:             balance,
          enrollDate:          row[13] ? Utilities.formatDate(new Date(row[13]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
          status:              String(row[15] || ''),
          paymentSlotsFilled:  slots,
          rowNumber:           r,
        }
      };
    }
  }

  return { status: 'error', message: 'Student ID "' + targetId + '" not found.' };
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

  var r     = s.rowNumber;
  var slots = s.paymentSlotsFilled;
  var amount = parseFloat(p.amount);
  var payDate = parseDate(p.paymentDate);

  // Payment slots: F/G = slot1 (cols 6/7), H/I = slot2 (cols 8/9), J/K = slot3 (cols 10/11)
  var amtCol, dateCol, slotNum;
  if      (slots === 0) { amtCol=6;  dateCol=7;  slotNum=1; }
  else if (slots === 1) { amtCol=8;  dateCol=9;  slotNum=2; }
  else if (slots === 2) { amtCol=10; dateCol=11; slotNum=3; }
  else return { status: 'error', message: 'All 3 payment slots are already full for this student.' };

  sheet.getRange(r, amtCol).setValue(amount);
  sheet.getRange(r, dateCol).setValue(payDate);
  sheet.getRange(r, amtCol).setNumberFormat('"$"#,##0.00');
  sheet.getRange(r, dateCol).setNumberFormat('yyyy-mm-dd');

  // Re-read balance after formula recalc (force recalc)
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
    return { error: true, status: 'error', message: 'Sheet "'+SHEET_NAME+'" not found. Available: '+all };
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
//  TEST FUNCTIONS — run from Apps Script editor to verify
// ============================================================
function testAddStudent() {
  var r = addStudent({
    studentName: 'TEST DELETE ME', email: 'test@test.com',
    phone: '+961 70 000 000', coursePrice: '350', firstPayment: '100',
    enrollDate: new Date().toISOString().slice(0,10), accessDays: '90'
  });
  Logger.log(JSON.stringify(r));
}

function testLookup() {
  var r = lookupStudent({ studentId: 'CU00001' });
  Logger.log(JSON.stringify(r, null, 2));
}

function testRecordPayment() {
  // First run testLookup() to find a real ID, then paste it below
  var r = recordPayment({ studentId: 'CU00001', amount: '50', paymentDate: '2025-05-10' });
  Logger.log(JSON.stringify(r, null, 2));
}
