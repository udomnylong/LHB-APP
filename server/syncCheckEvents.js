// Incremental, additive-only sync: pulls CheckIn/CheckOut rows from the Google Sheet
// that aren't already in Cloud SQL and inserts them — does NOT delete or touch any
// existing Cloud SQL rows (unlike migrate-sheets-to-sql.js's full-refresh Phase 1 ETL).
//
// Needed because the admin "manual attendance edit" forms in hr-system.html (kept on
// Apps Script on purpose — see the migration plan) write straight to the Sheet instead
// of the Cloud Run API. Run on a schedule (see routes/admin.js) to keep Cloud SQL caught up.

const { google } = require('googleapis');
const { pool } = require('./db');

async function getSheetsClient() {
  // On Cloud Run this uses the service's runtime service account (ADC) — no key file
  // needed there. Locally, GOOGLE_APPLICATION_CREDENTIALS in .env points at the key file.
  const auth = new google.auth.GoogleAuth({ scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'] });
  const client = await auth.getClient();
  return google.sheets({ version: 'v4', auth: client });
}

async function loadSheetRows(sheets, sheetId, sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId, range: sheetName,
    valueRenderOption: 'UNFORMATTED_VALUE', dateTimeRenderOption: 'SERIAL_NUMBER',
  });
  const values = res.data.values || [];
  if (values.length < 2) return [];
  const headers = values[0].map((h) => String(h).trim());
  return values.slice(1)
    .filter((row) => row.some((c) => c !== '' && c !== null && c !== undefined))
    .map((row) => { const obj = {}; headers.forEach((h, i) => { obj[h] = row[i]; }); return obj; });
}

const SHEETS_EPOCH_UTC = Date.UTC(1899, 11, 30);
const pad = (n) => String(n).padStart(2, '0');
function toStr(v) { if (v === null || v === undefined) return null; const s = String(v).trim(); return s === '' ? null : s; }
function toNumber(v) { if (v === null || v === undefined || v === '') return null; const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, '')); return Number.isNaN(n) ? null : n; }
function toDateStr(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') { const d = new Date(SHEETS_EPOCH_UTC + Math.round(v * 86400000)); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; }
  const s = String(v).trim(); const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
  const d = new Date(s); return Number.isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
}
function toTimeStr(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') { const ts = Math.round((v % 1) * 86400); return `${pad(Math.floor(ts/3600))}:${pad(Math.floor((ts%3600)/60))}:${pad(ts%60)}`; }
  const s = String(v).trim(); const m = s.match(/^(\d{1,2}):(\d{2})(:(\d{2}))?/);
  return m ? `${pad(+m[1])}:${m[2]}:${m[4] || '00'}` : null;
}
function toTimestampTz(v, fallbackDateStr, fallbackTimeStr) {
  const millis = toNumber(v);
  if (millis) return new Date(millis).toISOString();
  if (fallbackDateStr && fallbackTimeStr) return new Date(`${fallbackDateStr}T${fallbackTimeStr}Z`).toISOString();
  return null;
}

async function syncTable(sheets, sheetId, sheetName, table) {
  const raw = await loadSheetRows(sheets, sheetId, sheetName);
  const staffCodes = new Set((await pool.query('SELECT staff_code FROM staff')).rows.map(r => r.staff_code));

  const existing = new Set(
    (await pool.query(`SELECT staff_code, event_date::text, event_time::text FROM ${table}`)).rows
      .map(r => `${r.staff_code}|${r.event_date}|${r.event_time}`)
  );

  let inserted = 0, skippedExisting = 0, skippedBadData = 0;
  const affected = new Set();

  for (const r of raw) {
    const staff_code = toStr(r.ID);
    const event_date = toDateStr(r.Date);
    const event_time = toTimeStr(r.Time);
    if (!staff_code || !event_date || !event_time || !staffCodes.has(staff_code)) { skippedBadData++; continue; }
    const key = `${staff_code}|${event_date}|${event_time}`;
    if (existing.has(key)) { skippedExisting++; continue; }

    await pool.query(
      `INSERT INTO ${table} (staff_code, project_name, event_date, event_time, event_timestamp,
                              latitude, longitude, accuracy, late_early, minutes, position, department)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)`,
      [staff_code, toStr(r.ProjectName), event_date, event_time,
       toTimestampTz(r.Timestamp, event_date, event_time),
       toNumber(r.Latitude), toNumber(r.Longitude), toNumber(r.Accuracy),
       toStr(r.LateEarly), toNumber(r.Minutes), toStr(r.Position), toStr(r.Department)]
    );
    existing.add(key);
    affected.add(`${staff_code}|${event_date}`);
    inserted++;
  }
  return { table, inserted, skippedExisting, skippedBadData, affected };
}

async function recomputeAttendance(pairs) {
  let updated = 0;
  for (const pair of pairs) {
    const [staff_code, date] = pair.split('|');
    const ci = await pool.query(
      `SELECT event_time, late_early, minutes FROM check_ins WHERE staff_code=$1 AND event_date=$2 ORDER BY event_time ASC LIMIT 1`,
      [staff_code, date]
    );
    const co = await pool.query(
      `SELECT event_time, late_early, minutes FROM check_outs WHERE staff_code=$1 AND event_date=$2 ORDER BY event_time DESC LIMIT 1`,
      [staff_code, date]
    );
    const ciRow = ci.rows[0], coRow = co.rows[0];
    const late = ciRow && ciRow.late_early === 'Late' ? `Late ${ciRow.minutes} min` : '';
    const early = coRow && coRow.late_early === 'Early' ? `Early ${coRow.minutes} min` : '';
    const status = ciRow ? (ciRow.late_early === 'Late' ? 'Late' : 'Present') : 'Present';

    await pool.query(
      `INSERT INTO attendance (staff_code, date, check_in, check_out, late, early, status)
       VALUES ($1,$2,$3,$4,$5,$6,$7)
       ON CONFLICT (staff_code, date) DO UPDATE SET
         check_in = COALESCE(EXCLUDED.check_in, attendance.check_in),
         check_out = COALESCE(EXCLUDED.check_out, attendance.check_out),
         late = CASE WHEN EXCLUDED.check_in IS NOT NULL THEN EXCLUDED.late ELSE attendance.late END,
         early = CASE WHEN EXCLUDED.check_out IS NOT NULL THEN EXCLUDED.early ELSE attendance.early END,
         status = CASE WHEN EXCLUDED.check_in IS NOT NULL THEN EXCLUDED.status ELSE attendance.status END,
         updated_at = now()`,
      [staff_code, date, ciRow ? ciRow.event_time : null, coRow ? coRow.event_time : null, late, early, status]
    );
    updated++;
  }
  return updated;
}

async function runSync() {
  const sheetId = process.env.SHEET_ID;
  const sheets = await getSheetsClient();
  const ci = await syncTable(sheets, sheetId, 'CheckIn', 'check_ins');
  const co = await syncTable(sheets, sheetId, 'CheckOut', 'check_outs');
  const allAffected = new Set([...ci.affected, ...co.affected]);
  const attendanceUpdated = allAffected.size ? await recomputeAttendance(allAffected) : 0;
  return {
    check_ins: { inserted: ci.inserted, skippedExisting: ci.skippedExisting, skippedBadData: ci.skippedBadData },
    check_outs: { inserted: co.inserted, skippedExisting: co.skippedExisting, skippedBadData: co.skippedBadData },
    attendanceUpdated,
  };
}

module.exports = { runSync };
