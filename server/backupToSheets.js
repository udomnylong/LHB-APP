// Backup: pushes current Cloud SQL data for the 5 migrated resources
// (User, StaffInfo, CheckIn, CheckOut, Attendance) back into the Google Sheet as a
// snapshot — the reverse direction of migrate-sheets-to-sql.js's Phase 1 ETL.
//
// Deliberately scoped to only these 5 tabs: the other 7 sheets (StaffLeave, Project,
// StaffOT, Food, WorkPlace, Comment, EvaluateStaff) are still edited directly in the
// Sheet via Apps Script and are the source of truth for those — overwriting them from
// Cloud SQL would clobber live edits with a stale Phase-1 snapshot.

const { google } = require('googleapis');
const { pool } = require('./db');

async function getSheetsClient() {
  // On Cloud Run this uses the service's runtime service account (ADC) — the same
  // account used by syncCheckEvents.js, now also granted Editor access to the Sheet.
  const auth = new google.auth.GoogleAuth({ scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
  const client = await auth.getClient();
  return google.sheets({ version: 'v4', auth: client });
}

function fmtDate(d) {
  if (!d) return '';
  const dt = new Date(d);
  return Number.isNaN(dt.getTime()) ? '' : dt.toISOString().slice(0, 10);
}
function fmtTime(t) {
  return t ? String(t).slice(0, 8) : '';
}
function cell(v) {
  return v === null || v === undefined ? '' : String(v);
}

async function writeSheet(sheets, sheetId, sheetName, headers, rows) {
  const values = [headers, ...rows.map((r) => headers.map((h) => cell(r[h])))];
  await sheets.spreadsheets.values.clear({ spreadsheetId: sheetId, range: sheetName });
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${sheetName}!A1`,
    valueInputOption: 'RAW',
    requestBody: { values },
  });
  return rows.length;
}

async function backupUsers(sheets, sheetId) {
  const { rows } = await pool.query(
    'SELECT username, password_plain, name, role, email, department, position FROM users ORDER BY username'
  );
  const mapped = rows.map((r) => ({
    Username: r.username, Password: r.password_plain, Name: r.name, Role: r.role,
    Email: r.email, Department: r.department, Position: r.position,
    SessionToken: '', SessionTime: '',
  }));
  const headers = ['Username', 'Password', 'Name', 'Role', 'Email', 'Department', 'Position', 'SessionToken', 'SessionTime'];
  return writeSheet(sheets, sheetId, 'User', headers, mapped);
}

async function backupStaff(sheets, sheetId) {
  const { rows } = await pool.query('SELECT * FROM staff ORDER BY staff_code');
  const mapped = rows.map((r) => ({
    ID: r.staff_code, Name: r.name, NameLatin: r.name_latin, Sex: r.sex, LV: r.lv,
    Position: r.position, Department: r.department, ProjectName: r.project_name,
    DateOfBirth: fmtDate(r.date_of_birth), StartingDate: fmtDate(r.starting_date), ResignDate: fmtDate(r.resign_date),
    Salary: r.salary, Gmail: r.gmail, BankName: r.bank_name, BankNumber: r.bank_number,
    Photo: r.photo_url, Phone: r.phone, EmploymentStatus: r.employment_status, TelegramChatId: r.telegram_chat_id,
    OTP: '', OTPExpire: '',
  }));
  const headers = ['ID', 'Name', 'NameLatin', 'Sex', 'LV', 'Position', 'Department', 'ProjectName',
    'DateOfBirth', 'StartingDate', 'ResignDate', 'Salary', 'Gmail', 'BankName', 'BankNumber',
    'Photo', 'Phone', 'EmploymentStatus', 'TelegramChatId', 'OTP', 'OTPExpire'];
  return writeSheet(sheets, sheetId, 'StaffInfo', headers, mapped);
}

async function backupAttendance(sheets, sheetId) {
  const { rows } = await pool.query(`
    SELECT a.*, s.name, s.position AS staff_position, s.department AS staff_department, s.project_name
    FROM attendance a LEFT JOIN staff s ON s.staff_code = a.staff_code
    ORDER BY a.date, a.staff_code`);
  const mapped = rows.map((r) => ({
    ID: r.staff_code, Name: r.name, Position: r.staff_position, Department: r.staff_department,
    ProjectName: r.project_name, CheckIn: fmtTime(r.check_in), CheckOut: fmtTime(r.check_out),
    Late: r.late, Early: r.early, Status: r.status, Date: fmtDate(r.date),
  }));
  const headers = ['ID', 'Name', 'Position', 'Department', 'ProjectName', 'CheckIn', 'CheckOut', 'Late', 'Early', 'Status', 'Date'];
  return writeSheet(sheets, sheetId, 'Attendance', headers, mapped);
}

async function backupCheckEvents(sheets, sheetId, table, sheetName) {
  const { rows } = await pool.query(`
    SELECT e.*, s.name, s.gmail
    FROM ${table} e LEFT JOIN staff s ON s.staff_code = e.staff_code
    ORDER BY e.event_date, e.event_time`);
  const mapped = rows.map((r) => ({
    ID: r.staff_code, Name: r.name, Gmail: r.gmail, ProjectName: r.project_name,
    Date: fmtDate(r.event_date), Time: fmtTime(r.event_time),
    Timestamp: r.event_timestamp ? new Date(r.event_timestamp).getTime() : '',
    Latitude: r.latitude, Longitude: r.longitude, Accuracy: r.accuracy,
    LateEarly: r.late_early, Minutes: r.minutes, Position: r.position, Department: r.department,
  }));
  const headers = ['ID', 'Name', 'Gmail', 'ProjectName', 'Date', 'Time', 'Timestamp',
    'Latitude', 'Longitude', 'Accuracy', 'LateEarly', 'Minutes', 'Position', 'Department'];
  return writeSheet(sheets, sheetId, sheetName, headers, mapped);
}

async function runBackup() {
  const sheetId = process.env.SHEET_ID;
  const sheets = await getSheetsClient();
  const users = await backupUsers(sheets, sheetId);
  const staff = await backupStaff(sheets, sheetId);
  const attendance = await backupAttendance(sheets, sheetId);
  const checkIns = await backupCheckEvents(sheets, sheetId, 'check_ins', 'CheckIn');
  const checkOuts = await backupCheckEvents(sheets, sheetId, 'check_outs', 'CheckOut');
  return { users, staff, attendance, checkIns, checkOuts };
}

module.exports = { runBackup };
