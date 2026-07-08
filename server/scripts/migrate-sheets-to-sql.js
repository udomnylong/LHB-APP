// One-time (re-runnable) ETL: Google Sheets -> Cloud SQL for PostgreSQL.
// Phase 1 of the Sheets -> Cloud SQL migration (see plan: mossy-toasting-treasure).
//
// Each table is loaded with DELETE + INSERT inside a transaction, so re-running
// this script always leaves the DB matching the current sheet contents exactly
// (safe to run repeatedly while iterating, before Phase 2 cutover).
//
// Required env vars (see ../.env.example):
//   SHEET_ID                      Google Sheet spreadsheet ID (Code.gs SS_ID)
//   GOOGLE_APPLICATION_CREDENTIALS  path to a service-account JSON key with
//                                    read access to the sheet (share the sheet
//                                    with the service account's email)
//   DATABASE_URL                  postgres connection string

require('dotenv').config();
const { google } = require('googleapis');
const { Pool } = require('pg');

const SHEET_ID = process.env.SHEET_ID;
if (!SHEET_ID) throw new Error('SHEET_ID env var is required');

const pool = new Pool({ connectionString: process.env.DATABASE_URL });

const orphanReport = []; // { table, sheetRow, staffCode }

// ── Sheets helpers ──────────────────────────────────────────────

async function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
  const client = await auth.getClient();
  return google.sheets({ version: 'v4', auth: client });
}

async function loadSheetRows(sheets, sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: sheetName,
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER',
  });
  const values = res.data.values || [];
  if (values.length < 2) return [];
  const headers = values[0].map((h) => String(h).trim());
  return values.slice(1)
    .filter((row) => row.some((c) => c !== '' && c !== null && c !== undefined))
    .map((row) => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

// ── Cell coercion (mirrors the Sheets serial-date quirks Code.gs already handles) ──

const SHEETS_EPOCH_UTC = Date.UTC(1899, 11, 30);
const pad = (n) => String(n).padStart(2, '0');

function serialToUtcDate(serial) {
  return new Date(SHEETS_EPOCH_UTC + Math.round(serial * 86400000));
}

function toStr(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  return s === '' ? null : s;
}

function toNumber(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, ''));
  return Number.isNaN(n) ? null : n;
}

function toDateStr(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') {
    const d = serialToUtcDate(v);
    return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  }
  const s = String(v).trim();
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
}

function toTimeStr(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') {
    const totalSeconds = Math.round((v % 1) * 86400);
    const h = Math.floor(totalSeconds / 3600);
    const m = Math.floor((totalSeconds % 3600) / 60);
    const s = totalSeconds % 60;
    return `${pad(h)}:${pad(m)}:${pad(s)}`;
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2}):(\d{2})(:(\d{2}))?/);
  return m ? `${pad(+m[1])}:${m[2]}:${m[4] || '00'}` : null;
}

// Sheet 'Timestamp' columns store epoch millis as a string (set server-side by Code.gs).
function toTimestampTz(v, fallbackDateStr, fallbackTimeStr) {
  const millis = toNumber(v);
  if (millis) return new Date(millis).toISOString();
  if (fallbackDateStr && fallbackTimeStr) return new Date(`${fallbackDateStr}T${fallbackTimeStr}Z`).toISOString();
  return null;
}

// ── Table loaders ───────────────────────────────────────────────
// Each loader: (client, sheetRows, ctx) -> { inserted, skipped }

async function reload(client, table, columns, rows) {
  await client.query(`DELETE FROM ${table}`);
  if (rows.length === 0) return 0;
  const placeholdersPerRow = columns.length;
  const values = [];
  const tuples = rows.map((row, i) => {
    const base = i * placeholdersPerRow;
    columns.forEach((col) => values.push(row[col]));
    return `(${columns.map((_, j) => `$${base + j + 1}`).join(', ')})`;
  });
  const sql = `INSERT INTO ${table} (${columns.join(', ')}) VALUES ${tuples.join(', ')}`;
  await client.query(sql, values);
  return rows.length;
}

async function loadUsers(client, sheets) {
  const raw = await loadSheetRows(sheets, 'User');
  const rows = raw.map((r) => ({
    username: toStr(r.Username),
    password_plain: toStr(r.Password),
    name: toStr(r.Name),
    role: toStr(r.Role) || 'user',
    email: toStr(r.Email),
    department: toStr(r.Department),
    position: toStr(r.Position),
  })).filter((r) => r.username);
  const inserted = await reload(client, 'users',
    ['username', 'password_plain', 'name', 'role', 'email', 'department', 'position'], rows);
  return { table: 'users', inserted, skipped: raw.length - rows.length };
}

async function loadStaff(client, sheets) {
  const raw = await loadSheetRows(sheets, 'StaffInfo');
  const rows = raw.map((r) => ({
    staff_code: toStr(r.ID),
    name: toStr(r.Name),
    name_latin: toStr(r.NameLatin),
    sex: toStr(r.Sex),
    lv: toStr(r.LV),
    position: toStr(r.Position),
    department: toStr(r.Department),
    project_name: toStr(r.ProjectName),
    date_of_birth: toDateStr(r.DateOfBirth),
    starting_date: toDateStr(r.StartingDate),
    resign_date: toDateStr(r.ResignDate),
    salary: toNumber(r.Salary),
    gmail: toStr(r.Gmail),
    bank_name: toStr(r.BankName),
    bank_number: toStr(r.BankNumber),
    photo_url: toStr(r.Photo),
    phone: toStr(r.Phone),
    employment_status: toStr(r.EmploymentStatus),
    telegram_chat_id: toStr(r.TelegramChatId),
  })).filter((r) => r.staff_code);
  const inserted = await reload(client, 'staff', [
    'staff_code', 'name', 'name_latin', 'sex', 'lv', 'position', 'department', 'project_name',
    'date_of_birth', 'starting_date', 'resign_date', 'salary', 'gmail', 'bank_name', 'bank_number',
    'photo_url', 'phone', 'employment_status', 'telegram_chat_id',
  ], rows);
  return { table: 'staff', inserted, skipped: raw.length - rows.length, staffCodes: new Set(rows.map((r) => r.staff_code)) };
}

async function loadProjects(client, sheets) {
  const raw = await loadSheetRows(sheets, 'Project');
  const rows = raw.map((r) => ({
    project_id: toStr(r.ProjectID) || toStr(r.ProjectName),
    project_name: toStr(r.ProjectName),
    location: toStr(r.Location),
    latitude: toNumber(r.Latitude),
    longitude: toNumber(r.Longitude),
    radius: toNumber(r.Radius),
    status: toStr(r.Status),
  })).filter((r) => r.project_id);
  const inserted = await reload(client, 'projects',
    ['project_id', 'project_name', 'location', 'latitude', 'longitude', 'radius', 'status'], rows);
  return { table: 'projects', inserted, skipped: raw.length - rows.length };
}

async function loadAttendance(client, sheets, staffCodes) {
  const raw = await loadSheetRows(sheets, 'Attendance');
  const seen = new Set();
  const rows = [];
  raw.forEach((r) => {
    const staff_code = toStr(r.ID);
    const date = toDateStr(r.Date);
    if (!staff_code || !date) return orphanReport.push({ table: 'attendance', reason: 'missing ID/Date', row: r });
    if (!staffCodes.has(staff_code)) return orphanReport.push({ table: 'attendance', reason: 'unknown staff_code', staffCode: staff_code });
    const key = staff_code + '|' + date;
    if (seen.has(key)) return orphanReport.push({ table: 'attendance', reason: 'duplicate staff_code+date, kept first', staffCode: staff_code, row: r });
    seen.add(key);
    rows.push({
      staff_code, date,
      check_in: toTimeStr(r.CheckIn),
      check_out: toTimeStr(r.CheckOut),
      late: toStr(r.Late),
      early: toStr(r.Early),
      status: toStr(r.Status),
    });
  });
  const inserted = await reload(client, 'attendance',
    ['staff_code', 'date', 'check_in', 'check_out', 'late', 'early', 'status'], rows);
  return { table: 'attendance', inserted, skipped: raw.length - rows.length };
}

async function loadCheckEvents(client, sheets, sheetName, table, staffCodes) {
  const raw = await loadSheetRows(sheets, sheetName);
  const rows = [];
  raw.forEach((r) => {
    const staff_code = toStr(r.ID);
    const event_date = toDateStr(r.Date);
    const event_time = toTimeStr(r.Time);
    if (!staff_code || !event_date || !event_time) return orphanReport.push({ table, reason: 'missing ID/Date/Time', row: r });
    if (!staffCodes.has(staff_code)) return orphanReport.push({ table, reason: 'unknown staff_code', staffCode: staff_code });
    rows.push({
      staff_code, project_name: toStr(r.ProjectName), event_date, event_time,
      event_timestamp: toTimestampTz(r.Timestamp, event_date, event_time),
      latitude: toNumber(r.Latitude), longitude: toNumber(r.Longitude), accuracy: toNumber(r.Accuracy),
      late_early: toStr(r.LateEarly), minutes: toNumber(r.Minutes),
      position: toStr(r.Position), department: toStr(r.Department),
    });
  });
  const inserted = await reload(client, table, [
    'staff_code', 'project_name', 'event_date', 'event_time', 'event_timestamp',
    'latitude', 'longitude', 'accuracy', 'late_early', 'minutes', 'position', 'department',
  ], rows);
  return { table, inserted, skipped: raw.length - rows.length };
}

async function loadStaffLeave(client, sheets, staffCodes) {
  const raw = await loadSheetRows(sheets, 'StaffLeave');
  const rows = [];
  raw.forEach((r) => {
    const staff_code = toStr(r.ID);
    if (!staff_code || !staffCodes.has(staff_code)) return orphanReport.push({ table: 'staff_leave', reason: 'unknown/missing staff_code', staffCode: staff_code });
    rows.push({
      staff_code, type_of_leave: toStr(r.TypeOfLeave),
      start_date: toDateStr(r.StartDate), end_date: toDateStr(r.EndDate),
      days: toNumber(r.Days), reason: toStr(r.Reason), status: toStr(r.Status),
    });
  });
  const inserted = await reload(client, 'staff_leave',
    ['staff_code', 'type_of_leave', 'start_date', 'end_date', 'days', 'reason', 'status'], rows);
  return { table: 'staff_leave', inserted, skipped: raw.length - rows.length };
}

async function loadStaffOt(client, sheets, staffCodes) {
  const raw = await loadSheetRows(sheets, 'StaffOT');
  const rows = [];
  raw.forEach((r) => {
    const staff_code = toStr(r.ID);
    if (!staff_code || !staffCodes.has(staff_code)) return orphanReport.push({ table: 'staff_ot', reason: 'unknown/missing staff_code', staffCode: staff_code });
    rows.push({
      staff_code, date: toDateStr(r.Date), hours: toNumber(r.Hours),
      time_from: toTimeStr(r.TimeFrom), time_to: toTimeStr(r.TimeTo),
      type_of_work: toStr(r.TypeOfWork), reason: toStr(r.Reason), status: toStr(r.Status),
      // The real hr-system.html UI only ever wrote TypeOfWork + Remark (TimeFrom/TimeTo/
      // Reason above are legacy columns from setupHeaders() that were never populated).
      remark: toStr(r.Remark),
    });
  });
  const inserted = await reload(client, 'staff_ot',
    ['staff_code', 'date', 'hours', 'time_from', 'time_to', 'type_of_work', 'reason', 'status', 'remark'], rows);
  return { table: 'staff_ot', inserted, skipped: raw.length - rows.length };
}

async function loadFood(client, sheets, staffCodes) {
  const raw = await loadSheetRows(sheets, 'Food');
  const rows = raw.map((r) => {
    const staff_code = toStr(r.ID);
    if (staff_code && !staffCodes.has(staff_code)) orphanReport.push({ table: 'food_records', reason: 'unknown staff_code kept as NULL', staffCode: staff_code });
    return {
      date: toDateStr(r.Date), staff_code: staff_code && staffCodes.has(staff_code) ? staff_code : null,
      name: toStr(r.Name), sex: toStr(r.Sex), position: toStr(r.Position), project_name: toStr(r.ProjectName),
      morning: toStr(r.Morning), lunch: toStr(r.Lunch), evening: toStr(r.Evening),
      total: toNumber(r.Total), unit_price: toNumber(r.UnitPrice), total_price: toNumber(r.TotalPrice),
      photo_morning_url: toStr(r.PhotoMorning), photo_lunch_url: toStr(r.PhotoLunch), photo_evening_url: toStr(r.PhotoEvening),
      comment: toStr(r.Comment), remark: toStr(r.Remark),
    };
  });
  const inserted = await reload(client, 'food_records', [
    'date', 'staff_code', 'name', 'sex', 'position', 'project_name', 'morning', 'lunch', 'evening',
    'total', 'unit_price', 'total_price', 'photo_morning_url', 'photo_lunch_url', 'photo_evening_url', 'comment', 'remark',
  ], rows);
  return { table: 'food_records', inserted, skipped: 0 };
}

async function loadReportShaped(client, sheets, sheetName, table, staffCodes) {
  const raw = await loadSheetRows(sheets, sheetName);
  const rows = raw.map((r) => {
    const staff_code = toStr(r.ID);
    if (staff_code && !staffCodes.has(staff_code)) orphanReport.push({ table, reason: 'unknown staff_code kept as NULL', staffCode: staff_code });
    return {
      date: toDateStr(r.Date), time: toTimeStr(r.Time),
      staff_code: staff_code && staffCodes.has(staff_code) ? staff_code : null,
      department: toStr(r.Department), project_name: toStr(r.ProjectName),
      comment: toStr(r.Comment), photo_url: toStr(r.Photo), status: toStr(r.Status),
    };
  });
  const inserted = await reload(client, table,
    ['date', 'time', 'staff_code', 'department', 'project_name', 'comment', 'photo_url', 'status'], rows);
  return { table, inserted, skipped: 0 };
}

async function loadEvaluations(client, sheets, staffCodes) {
  const raw = await loadSheetRows(sheets, 'EvaluateStaff');
  const rows = raw.map((r) => {
    const staff_code = toStr(r.StaffID);
    if (staff_code && !staffCodes.has(staff_code)) orphanReport.push({ table: 'staff_evaluations', reason: 'unknown staff_code kept as NULL', staffCode: staff_code });
    return {
      request_by: toStr(r.RequestBy), staff_code: staff_code && staffCodes.has(staff_code) ? staff_code : null,
      staff_name: toStr(r.StaffName), date_evaluate: toDateStr(r.DateEvaluate), kpi_score: toNumber(r.KPIScore),
      previous_salary: toNumber(r.PreviousSalary), current_salary: toNumber(r.CurrentSalary),
      approved_by: toStr(r.ApprovedBy), remark: toStr(r.Remark),
    };
  });
  const inserted = await reload(client, 'staff_evaluations', [
    'request_by', 'staff_code', 'staff_name', 'date_evaluate', 'kpi_score', 'previous_salary', 'current_salary', 'approved_by', 'remark',
  ], rows);
  return { table: 'staff_evaluations', inserted, skipped: 0 };
}

// ── Orchestration ───────────────────────────────────────────────

async function main() {
  const sheets = await getSheetsClient();
  const client = await pool.connect();
  const summary = [];
  try {
    // Independent tables first (no FK dependency)
    await client.query('BEGIN');
    summary.push(await loadUsers(client, sheets));
    await client.query('COMMIT');

    await client.query('BEGIN');
    const staffResult = await loadStaff(client, sheets);
    summary.push(staffResult);
    await client.query('COMMIT');

    await client.query('BEGIN');
    summary.push(await loadProjects(client, sheets));
    await client.query('COMMIT');

    const staffCodes = staffResult.staffCodes;

    // Dependent tables (FK on staff)
    await client.query('BEGIN');
    summary.push(await loadAttendance(client, sheets, staffCodes));
    await client.query('COMMIT');

    await client.query('BEGIN');
    summary.push(await loadCheckEvents(client, sheets, 'CheckIn', 'check_ins', staffCodes));
    await client.query('COMMIT');

    await client.query('BEGIN');
    summary.push(await loadCheckEvents(client, sheets, 'CheckOut', 'check_outs', staffCodes));
    await client.query('COMMIT');

    await client.query('BEGIN');
    summary.push(await loadStaffLeave(client, sheets, staffCodes));
    await client.query('COMMIT');

    await client.query('BEGIN');
    summary.push(await loadStaffOt(client, sheets, staffCodes));
    await client.query('COMMIT');

    await client.query('BEGIN');
    summary.push(await loadFood(client, sheets, staffCodes));
    await client.query('COMMIT');

    await client.query('BEGIN');
    summary.push(await loadReportShaped(client, sheets, 'WorkPlace', 'work_place_reports', staffCodes));
    await client.query('COMMIT');

    await client.query('BEGIN');
    summary.push(await loadReportShaped(client, sheets, 'Comment', 'comments', staffCodes));
    await client.query('COMMIT');

    await client.query('BEGIN');
    summary.push(await loadEvaluations(client, sheets, staffCodes));
    await client.query('COMMIT');
  } catch (err) {
    await client.query('ROLLBACK');
    throw err;
  } finally {
    client.release();
    await pool.end();
  }

  console.log('\n=== ETL summary ===');
  summary.forEach((s) => console.log(`${s.table.padEnd(20)} inserted=${s.inserted}  skipped=${s.skipped ?? 0}`));

  if (orphanReport.length) {
    console.log(`\n=== Orphan/data-quality punch list (${orphanReport.length} rows) ===`);
    const byTable = {};
    orphanReport.forEach((o) => { (byTable[o.table] ||= []).push(o); });
    Object.entries(byTable).forEach(([table, items]) => {
      console.log(`\n${table}: ${items.length} issue(s)`);
      items.slice(0, 20).forEach((o) => console.log(`  - ${o.reason}${o.staffCode ? ` (staffCode=${o.staffCode})` : ''}`));
      if (items.length > 20) console.log(`  ... and ${items.length - 20} more`);
    });
  } else {
    console.log('\nNo orphaned rows found.');
  }
}

main().catch((err) => {
  console.error('ETL failed:', err);
  process.exit(1);
});
