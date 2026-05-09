/* ============================================
   CUPOD APPS SCRIPT — BACKEND
   ============================================
   This is the file you paste into Google Apps Script.
   Detailed instructions in SETUP_GUIDE.md (Step 3).
   ============================================ */

const TRACKER_SHEET = 'CUPOD Student Tracker';
const ACCESS_DAYS_DEFAULT = 90;

/**
 * doGet — used to verify the script is alive (open the Web App URL in a browser).
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      message: 'CUPOD Apps Script is alive.',
      time: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * doPost — receives the form submission from the website.
 */
function doPost(e) {
  try {
    if (!e || !e.parameter || !e.parameter.payload) {
      return jsonOut({ status: 'error', message: 'No payload received' });
    }

    const payload = JSON.parse(e.parameter.payload);
    const result = addStudent(payload);
    return jsonOut(result);

  } catch (err) {
    Logger.log('doPost error: ' + err);
    return jsonOut({ status: 'error', message: 'Server error: ' + err.message });
  }
}

/**
 * Add a new student row to the tracker.
 */
function addStudent(p) {
  // ---- validate ----
  if (!p.studentName || String(p.studentName).trim() === '') {
    return { status: 'error', message: 'Student name is required' };
  }
  if (!p.phone || String(p.phone).trim() === '') {
    return { status: 'error', message: 'Phone is required' };
  }
  if (isNaN(parseFloat(p.coursePrice))) {
    return { status: 'error', message: 'Course price must be a number' };
  }
  if (!p.enrollDate) {
    return { status: 'error', message: 'Enrollment date is required' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TRACKER_SHEET);
  if (!sheet) {
    return { status: 'error', message: 'Sheet "' + TRACKER_SHEET + '" not found' };
  }

  // ---- find next empty row ----
  // Header is row 1; data starts row 2. Use column B (Student Name) to find last filled.
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const targetRow = lastRow + 1;

  // ---- generate Student ID ----
  // Pattern: CU + 5-digit padded sequence based on data row index
  const sequenceNum = targetRow - 1;
  const studentId = 'CU' + pad(sequenceNum, 5);

  // ---- prepare values ----
  const courseStrPrice = parseFloat(p.coursePrice);
  const firstPay = parseFloat(p.firstPayment) || 0;
  const enrollDate = parseDate(p.enrollDate);
  const accessDays = parseInt(p.accessDays, 10) || ACCESS_DAYS_DEFAULT;

  // Column layout (matches the original Excel structure):
  // A: Student ID
  // B: Student Name
  // C: Email
  // D: Phone
  // E: Course Price
  // F: Payment 1 Amount
  // G: Payment 1 Date
  // H: Payment 2 Amount
  // I: Payment 2 Date
  // J: Payment 3 Amount
  // K: Payment 3 Date
  // L: Total Paid (formula)
  // M: Balance Due (formula)
  // N: Enrollment Date
  // O: Access End Date (formula)
  // P: Status (formula)
  // Q: Reminder (formula)

  const r = targetRow;
  const row = [
    studentId,
    String(p.studentName).trim(),
    String(p.email || '').trim(),
    String(p.phone).trim(),
    courseStrPrice,
    firstPay > 0 ? firstPay : '',
    firstPay > 0 ? enrollDate : '',
    '', '',  // Payment 2
    '', '',  // Payment 3
    '=SUM(F' + r + ',H' + r + ',J' + r + ')',                                              // L: Total Paid
    '=IF(E' + r + '<>"", E' + r + '-L' + r + ', "")',                                        // M: Balance
    enrollDate,                                                                              // N: Enrollment
    '=IF(N' + r + '<>"", N' + r + '+' + accessDays + ', "")',                                // O: Access End
    '=IF(O' + r + '="", "", IF(TODAY()>O' + r + ', "Expired", IF(TODAY()>=O' + r + '-10, "Expiring Soon", "Active")))',  // P: Status
    '=IF(P' + r + '="Expired", "Stop access!", IF(P' + r + '="Expiring Soon", "Renewal/payment due soon", "No action needed"))',  // Q: Reminder
  ];

  // ---- write row ----
  sheet.getRange(r, 1, 1, row.length).setValues([row]);

  // ---- formatting ----
  // Money columns
  [5, 6, 8, 10, 12, 13].forEach(c => sheet.getRange(r, c).setNumberFormat('"$"#,##0.00'));
  // Date columns
  [7, 9, 11, 14, 15].forEach(c => sheet.getRange(r, c).setNumberFormat('yyyy-mm-dd'));

  return {
    status: 'success',
    studentId: studentId,
    rowNumber: r,
    message: 'Student added successfully',
  };
}

// ---------- helpers ----------
function pad(num, size) {
  let s = String(num);
  while (s.length < size) s = '0' + s;
  return s;
}

function parseDate(value) {
  // Accepts ISO yyyy-mm-dd
  if (value instanceof Date) return value;
  const s = String(value).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  }
  return new Date(s);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================
   TEST FUNCTION
   ============================================
   Run this once from the Apps Script editor to confirm
   the connection works (Run → testAddStudent).
   ============================================ */
function testAddStudent() {
  const result = addStudent({
    studentName: 'Test Student (delete me)',
    email: 'test@example.com',
    phone: '+961 70 000 000',
    coursePrice: 350,
    firstPayment: 100,
    enrollDate: new Date().toISOString().slice(0, 10),
    accessDays: 90,
  });
  Logger.log(JSON.stringify(result, null, 2));
}
