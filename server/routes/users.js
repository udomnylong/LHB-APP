const express = require('express');
const { pool } = require('../db');
const { requireAuth, requireAdmin } = require('../middleware/auth');
const asyncHandler = require('../asyncHandler');

const router = express.Router();

// GET /api/users — HR system account list for the admin Settings screen.
// No auth required, matching the old Apps Script doGet (fully unauthenticated) —
// and this is actually *safer* than that: it omits password_plain/password_hash,
// which the old Sheets-era client used to fetch in plaintext for every user just
// to do local login matching (no longer needed now that /api/auth/login does it
// server-side).
router.get('/', asyncHandler(async (req, res) => {
  const { rows } = await pool.query(
    'SELECT username, name, role, email, department, position FROM users ORDER BY username'
  );
  res.json({ status: 'ok', data: rows });
}));

// Everything below creates accounts, changes roles, or deletes accounts — admin only,
// checked server-side (not just hidden client-side like most other admin actions in
// this app), since a regular logged-in user could otherwise grant themselves 'admin'.
router.use(requireAuth, requireAdmin);

// POST /api/users — create. Password stored in password_plain, matching the existing
// login comparison in auth.js (bcrypt/password_hash migration is a separate future task).
router.post('/', asyncHandler(async (req, res) => {
  const { username, password, name, role, email, department, position } = req.body || {};
  if (!username || !password) return res.status(400).json({ status: 'error', msg: 'username and password required' });

  try {
    await pool.query(
      `INSERT INTO users (username, password_plain, name, role, email, department, position)
       VALUES ($1, $2, $3, $4, $5, $6, $7)`,
      [username, password, name || '', role || 'user', email || '', department || '', position || '']
    );
    res.json({ status: 'ok' });
  } catch (err) {
    if (err.code === '23505') return res.status(409).json({ status: 'error', msg: 'Username already exists: ' + username });
    throw err;
  }
}));

// PUT /api/users/:username — edit profile fields and/or role; password only changes if
// a new one is sent.
router.put('/:username', asyncHandler(async (req, res) => {
  const { username } = req.params;
  const { password, name, role, email, department, position } = req.body || {};
  const setCols = [];
  const values = [username];
  const add = (col, val) => { values.push(val); setCols.push(`${col} = $${values.length}`); };
  if (password !== undefined)   add('password_plain', password);
  if (name !== undefined)       add('name', name);
  if (role !== undefined)       add('role', role);
  if (email !== undefined)      add('email', email);
  if (department !== undefined) add('department', department);
  if (position !== undefined)   add('position', position);
  if (!setCols.length) return res.status(400).json({ status: 'error', msg: 'No fields to update' });
  setCols.push('updated_at = now()');

  const result = await pool.query(`UPDATE users SET ${setCols.join(', ')} WHERE username = $1`, values);
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'User not found: ' + username });
  res.json({ status: 'ok' });
}));

// DELETE /api/users/:username — user_sessions.user_id has ON DELETE CASCADE, so this
// also immediately revokes any session the deleted user is currently using.
router.delete('/:username', asyncHandler(async (req, res) => {
  const { username } = req.params;
  if (req.user.username === username) return res.status(400).json({ status: 'error', msg: 'Cannot delete your own account while logged in as it' });
  const result = await pool.query('DELETE FROM users WHERE username = $1', [username]);
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'User not found: ' + username });
  res.json({ status: 'ok' });
}));

module.exports = router;
