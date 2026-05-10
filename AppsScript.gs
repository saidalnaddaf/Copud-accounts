// ============================================================
//  CUPOD APPS SCRIPT — FIXED VERSION
//  Uses doGet (URL parameters) — zero CORS issues guaranteed
// ============================================================

const SHEET_NAME = 'CUPOD Student Tracker';

function doGet(e) {
  try {
    if (!e || !e.parameter || !e.parameter.action) {
      return respond({ status: 'ok', message: 'CUPOD script is live!' });
    }
    if (e.parameter.action === 'addStudent') {
      return respond(addStudent(e.parameter));
    }
    return respond({ status: 'error', message: 'Unknown action' });
  } catch (err) {
    Logger.log('ERROR: ' + err.toString());
    return respond({ status: 'error', message: err.toString() });
  }
}

function doPost(e) {
  return doGet(e);
}

function addStudent(p) {
  if (!p.studentName || p.studentName.trim() === '') {
    return { status: 'error', message: 'Student name is required' };
  }
  if (!p.phone || p.phone.trim() === '') {
    return { status: 'error', message: 'Phone is required' };
  }
  if (!p.enrollDate) {
    return { status: 'error', message: 'Enrollment date is required' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    const allSheets = ss.getSheets().map(function(s) { return s.getName(); }).join(', ');
    return { status: 'error', message: 'Sheet not found. Available: ' + allSheets };
  }

  const lastRow = sheet.getLastRow();
  const nextRow = lastRow + 1;

  const coursePrice = parseFloat(p.coursePrice) || 0;
  const firstPayment = parseFloat(p.firstPayment) || 0;
  const accessDays = parseInt(p.accessDays) || 90;
  const studentId = 'CU' + String(nextRow - 1).padStart(5, '0');
  const enrollDate = new Date(p.enrollDate + 'T00:00:00');
  const r = nextRow;

  const rowData = [
    studentId,
    p.studentName.trim(),
    (p.email || '').trim(),
    p.phone.trim(),
    coursePrice,
    firstPayment > 0 ? firstPayment : '',
    firstPayment > 0 ? enrollDate : '',
    '', '', '', '',
    '', '', enrollDate, '', '', ''
  ];

  sheet.getRange(r, 1, 1, rowData.length).setValues([rowData]);
  sheet.getRange(r, 12).setFormula('=SUM(F'+r+',H'+r+',J'+r+')');
  sheet.getRange(r, 13).setFormula('=IF(E'+r+'<>"",E'+r+'-L'+r+',"")');
  sheet.getRange(r, 15).setFormula('=IF(N'+r+'<>"",N'+r+'+'+accessDays+',"")');
  sheet.getRange(r, 16).setFormula('=IF(O'+r+'="","",IF(TODAY()>O'+r+',"Expired",IF(TODAY()>=O'+r+'-10,"Expiring Soon","Active")))');
  sheet.getRange(r, 17).setFormula('=IF(P'+r+'="Expired","Stop access!",IF(P'+r+'="Expiring Soon","Renewal/payment due soon","No action needed"))');

  var moneyFmt = '"$"#,##0.00';
  [5,6,8,10,12,13].forEach(function(c){ sheet.getRange(r,c).setNumberFormat(moneyFmt); });
  [7,9,11,14,15].forEach(function(c){ sheet.getRange(r,c).setNumberFormat('yyyy-mm-dd'); });

  Logger.log('SUCCESS: ' + studentId + ' row ' + r);
  return { status: 'success', studentId: studentId, rowNumber: r, studentName: p.studentName.trim(), message: 'Student added' };
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function testAddStudent() {
  var result = addStudent({
    action: 'addStudent',
    studentName: 'TEST DELETE ME',
    email: 'test@test.com',
    phone: '+961 70 000 000',
    coursePrice: '350',
    firstPayment: '100',
    enrollDate: new Date().toISOString().slice(0,10),
    accessDays: '90'
  });
  Logger.log(JSON.stringify(result));
}
