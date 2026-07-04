const express = require('express');
const { pool, withTransaction } = require('../db');
const asyncHandler = require('../asyncHandler');
const { phNow, phDateStr, phTimeStr, phTimeStrFull } = require('../phTime');
const { sendTelegramMessage } = require('../telegramClient');

// No requireAuth here, on purpose: staff-portal.html's check-in (the primary caller)
// has no session/token concept at all — its "login" is a client-side Gmail+phone
// match with no server verification, exactly like the Code.gs backend it replaces,
// which likewise has zero server-side auth on any action. Matching that faithfully
// (rather than silently breaking check-ins) is the Phase 2 goal; real auth hardening
// for the whole app is Phase 5's job.
const router = express.Router();

const CHECKIN_CUTOFF_MIN = 8 * 60;   // 08:00
const CHECKOUT_CUTOFF_MIN = 17 * 60; // 17:00
const TELEGRAM_GROUP = process.env.TELEGRAM_GROUP || process.env.TELEGRAM_CHAT || '';

function minutesOf(timeStr) {
  const [h, m] = timeStr.split(':').map(Number);
  return h * 60 + m;
}

function toNum(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

async function recordEvent({ table, staffCode, req, res }) {
  const staffCodeTrim = String(staffCode || '').trim();
  if (!staffCodeTrim) return res.status(400).json({ status: 'error', msg: 'staff_code required' });

  const staffResult = await pool.query('SELECT * FROM staff WHERE staff_code = $1', [staffCodeTrim]);
  const staff = staffResult.rows[0];
  if (!staff) return res.status(404).json({ status: 'error', msg: 'Staff not found: ' + staffCodeTrim });

  const now = phNow();
  const eventDate = phDateStr(now);
  const eventTime = phTimeStr(now);
  const eventTimestamp = new Date(); // real instant, TIMESTAMPTZ column
  const totalMin = minutesOf(eventTime);
  const projectName = String(req.body.project_name || staff.project_name || '');
  const latitude = toNum(req.body.latitude);
  const longitude = toNum(req.body.longitude);
  const accuracy = toNum(req.body.accuracy);

  let lateEarly = '';
  let minutes = 0;
  let attendanceSql;
  let attendanceParams;

  if (table === 'check_ins') {
    if (totalMin > CHECKIN_CUTOFF_MIN) { lateEarly = 'Late'; minutes = totalMin - CHECKIN_CUTOFF_MIN; }
    const late = lateEarly ? `Late ${minutes} min` : '';
    const status = lateEarly ? 'Late' : 'Present';
    attendanceSql = `
      INSERT INTO attendance (staff_code, date, check_in, late, status)
      VALUES ($1, $2, $3, $4, $5)
      ON CONFLICT (staff_code, date) DO UPDATE
        SET check_in = EXCLUDED.check_in, late = EXCLUDED.late, status = EXCLUDED.status, updated_at = now()`;
    attendanceParams = [staffCodeTrim, eventDate, eventTime, late, status];
  } else {
    if (totalMin < CHECKOUT_CUTOFF_MIN) { lateEarly = 'Early'; minutes = CHECKOUT_CUTOFF_MIN - totalMin; }
    const early = lateEarly ? `Early ${minutes} min` : '';
    attendanceSql = `
      INSERT INTO attendance (staff_code, date, check_out, early, status)
      VALUES ($1, $2, $3, $4, COALESCE((SELECT status FROM attendance WHERE staff_code=$5 AND date=$6), 'Present'))
      ON CONFLICT (staff_code, date) DO UPDATE
        SET check_out = EXCLUDED.check_out, early = EXCLUDED.early, updated_at = now()`;
    attendanceParams = [staffCodeTrim, eventDate, eventTime, early, staffCodeTrim, eventDate];
  }

  await withTransaction(async (client) => {
    await client.query(
      `INSERT INTO ${table}
        (staff_code, project_name, event_date, event_time, event_timestamp, latitude, longitude, accuracy, late_early, minutes, position, department)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)`,
      [staffCodeTrim, projectName, eventDate, eventTime, eventTimestamp, latitude, longitude, accuracy, lateEarly, minutes, staff.position, staff.department]
    );
    await client.query(attendanceSql, attendanceParams);
  });

  const emoji = table === 'check_ins' ? '🟢' : '🟡';
  const label = table === 'check_ins' ? 'CHECK IN' : 'CHECK OUT';
  if (TELEGRAM_GROUP) {
    sendTelegramMessage(TELEGRAM_GROUP,
      `${emoji} ${label} | ${eventTime}\n\n${staff.name || ''}  ${staff.staff_code}\n${staff.position || ''} | ${staff.department || ''}\n${projectName}`
    );
  }

  res.json({ status: 'ok', date: eventDate, time: eventTime, lateEarly, minutes });
}

// POST /api/checkins
router.post('/checkins', asyncHandler((req, res) => recordEvent({ table: 'check_ins', staffCode: req.body.staff_code, req, res })));

// POST /api/checkouts
router.post('/checkouts', asyncHandler((req, res) => recordEvent({ table: 'check_outs', staffCode: req.body.staff_code, req, res })));

// GET /api/attendance?staffCode=&date=&from=&to=
router.get('/attendance', asyncHandler(async (req, res) => {
  const { staffCode, date, from, to } = req.query;
  const clauses = [];
  const params = [];

  if (staffCode) { params.push(staffCode); clauses.push(`a.staff_code = $${params.length}`); }
  if (date) { params.push(date); clauses.push(`a.date = $${params.length}`); }
  if (from) { params.push(from); clauses.push(`a.date >= $${params.length}`); }
  if (to) { params.push(to); clauses.push(`a.date <= $${params.length}`); }

  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  // Joined to staff for display fields (Name/Position/Department/ProjectName) — these
  // reflect the staff member's *current* assignment, not a point-in-time snapshot the
  // old Attendance sheet stored per row. Acceptable for a live dashboard view.
  const { rows } = await pool.query(
    `SELECT a.*, s.name, s.position AS staff_position, s.department AS staff_department, s.project_name
     FROM attendance a LEFT JOIN staff s ON s.staff_code = a.staff_code
     ${where}
     ORDER BY a.date DESC, a.staff_code`,
    params
  );
  res.json({ status: 'ok', data: rows });
}));

async function listEvents(table, req, res) {
  const { staffCode, date, from, to } = req.query;
  const clauses = [];
  const params = [];

  if (staffCode) { params.push(staffCode); clauses.push(`e.staff_code = $${params.length}`); }
  if (date) { params.push(date); clauses.push(`e.event_date = $${params.length}`); }
  if (from) { params.push(from); clauses.push(`e.event_date >= $${params.length}`); }
  if (to) { params.push(to); clauses.push(`e.event_date <= $${params.length}`); }

  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  const { rows } = await pool.query(
    `SELECT e.*, s.name, s.gmail
     FROM ${table} e LEFT JOIN staff s ON s.staff_code = e.staff_code
     ${where} ORDER BY e.event_date DESC, e.event_time DESC`,
    params
  );
  res.json({ status: 'ok', data: rows });
}

// GET /api/checkins?staffCode=&date=&from=&to= — raw event log (dashboards/history)
router.get('/checkins', asyncHandler((req, res) => listEvents('check_ins', req, res)));
// GET /api/checkouts?staffCode=&date=&from=&to=
router.get('/checkouts', asyncHandler((req, res) => listEvents('check_outs', req, res)));

module.exports = router;
