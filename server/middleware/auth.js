const { pool } = require('../db');

const SESSION_TTL_MS = 10 * 60 * 1000; // matches the old Code.gs SESSION_TIMEOUT

function extractToken(req) {
  const header = req.headers.authorization || '';
  const match = header.match(/^Bearer (.+)$/i);
  return match ? match[1].trim() : null;
}

async function requireAuth(req, res, next) {
  const token = extractToken(req);
  if (!token) return res.status(401).json({ status: 'error', msg: 'Missing bearer token' });

  const { rows } = await pool.query(
    `SELECT s.id AS session_id, s.expires_at, s.revoked_at, u.id AS user_id, u.username, u.name, u.role, u.email, u.department, u.position
     FROM user_sessions s JOIN users u ON u.id = s.user_id
     WHERE s.session_token = $1`,
    [token]
  );
  const session = rows[0];
  if (!session || session.revoked_at || new Date(session.expires_at) < new Date()) {
    return res.status(401).json({ status: 'error', msg: 'Session expired or invalid' });
  }

  req.session = { id: session.session_id, token };
  req.user = {
    id: session.user_id, username: session.username, name: session.name,
    role: session.role, email: session.email, department: session.department, position: session.position,
  };
  next();
}

// Chain after requireAuth (needs req.user). Every other write route in this app only
// checks "is logged in" and relies on hr-system.html's client-side page permissions to
// hide admin-only actions — fine for approving a leave request, not fine for creating
// accounts/changing roles, so user management gets a real server-side check.
function requireAdmin(req, res, next) {
  if (!req.user || (req.user.role || '').toLowerCase() !== 'admin') {
    return res.status(403).json({ status: 'error', msg: 'Admin role required' });
  }
  next();
}

module.exports = { requireAuth, requireAdmin, extractToken, SESSION_TTL_MS };
