const assert = require('node:assert/strict');
const express = require('express');
const http = require('node:http');
const test = require('node:test');
const { pool } = require('../db');
const { hashPassword, verifyPassword } = require('../passwords');
const { requireAuth, requirePrivileged, requireProjectManager } = require('../middleware/auth');
const authRoutes = require('../routes/auth');
const staffRoutes = require('../routes/staff');

test('password hashes verify only the matching password', async () => {
  const encoded = await hashPassword('123456');
  assert.match(encoded, /^scrypt\$[0-9a-f]{32}\$[0-9a-f]{128}$/);
  assert.equal(await verifyPassword('123456', encoded), true);
  assert.equal(await verifyPassword('654321', encoded), false);
  assert.equal(await verifyPassword('123456', 'invalid'), false);
});

test('authentication middleware forwards database errors to Express', async () => {
  const originalQuery = pool.query;
  const databaseError = new Error('database unavailable');
  pool.query = async () => { throw databaseError; };
  let forwardedError;

  try {
    await requireAuth(
      { headers: { authorization: 'Bearer test-token' } },
      {},
      (error) => { forwardedError = error; }
    );
    assert.equal(forwardedError, databaseError);
  } finally {
    pool.query = originalQuery;
  }
});

test('project writes allow HR managers but reject ordinary users', () => {
  let nextCalled = false;
  const response = {
    statusCode: null,
    status(code) { this.statusCode = code; return this; },
    json() { return this; },
  };

  requireProjectManager({ user: { role: 'hr' } }, response, () => { nextCalled = true; });
  assert.equal(nextCalled, true);

  nextCalled = false;
  requireProjectManager({ user: { role: 'user' } }, response, () => { nextCalled = true; });
  assert.equal(nextCalled, false);
  assert.equal(response.statusCode, 403);
});

test('privileged staff operations reject ordinary users', () => {
  let nextCalled = false;
  const response = {
    statusCode: null,
    status(code) { this.statusCode = code; return this; },
    json() { return this; },
  };

  requirePrivileged({ user: { role: 'user' } }, response, () => { nextCalled = true; });
  assert.equal(nextCalled, false);
  assert.equal(response.statusCode, 403);
});

test('staff login is server-verified and public staff reads omit sensitive fields', async () => {
  const originalQuery = pool.query;
  const queryLog = [];
  const staff = {
    staff_code: 'EMP001',
    name: 'Test Staff',
    gmail: 'staff@example.com',
    phone: '0123456789',
    employment_status: 'active',
    salary: '1000',
    bank_number: '123456',
    telegram_chat_id: 'private-chat',
    portal_password_hash: null,
  };
  pool.query = async (sql) => {
    queryLog.push(sql);
    if (sql.includes('lower(trim(gmail))')) return { rows: [staff] };
    if (sql.includes('FROM staff ORDER BY staff_code')) {
      return { rows: [{ staff_code: 'EMP001', name: 'Test Staff' }] };
    }
    if (sql.includes('FROM staff WHERE staff_code = $1')) {
      return { rows: [{
        staff_code: 'EMP001',
        name: 'Test Staff',
        starting_date: '2020-01-01',
      }] };
    }
    return { rows: [], rowCount: 1 };
  };

  const app = express();
  app.use(express.json());
  app.use('/api/auth', authRoutes);
  app.use('/api/staff', staffRoutes);
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const address = server.address();

  try {
    const login = await fetch(`http://127.0.0.1:${address.port}/api/auth/staff-login`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email: 'staff@example.com', password: '0123456789' }),
    });
    const loginBody = await login.json();
    assert.equal(login.status, 200);
    assert.ok(loginBody.token);
    assert.equal('portal_password_hash' in loginBody.staff, false);
    assert.equal('telegram_chat_id' in loginBody.staff, false);
    assert.ok(queryLog.some((sql) => sql.includes('portal_password_hash = $1')));

    const directory = await fetch(`http://127.0.0.1:${address.port}/api/staff/directory`);
    assert.equal(directory.status, 200);
    assert.deepEqual((await directory.json()).data, [{ staff_code: 'EMP001', name: 'Test Staff' }]);
    const directoryQuery = queryLog.find((sql) => sql.includes('FROM staff ORDER BY staff_code'));
    assert.doesNotMatch(directoryQuery, /salary|bank_number|phone|gmail/i);

    const card = await fetch(`http://127.0.0.1:${address.port}/api/staff/public/EMP001`);
    const cardBody = await card.json();
    assert.equal(card.status, 200);
    assert.equal(cardBody.data.staff_code, 'EMP001');
    assert.equal('salary' in cardBody.data, false);
    const cardQuery = queryLog.find((sql) => sql.includes('starting_date'));
    assert.doesNotMatch(cardQuery, /salary|bank_number|phone|gmail/i);

    const privateDirectory = await fetch(`http://127.0.0.1:${address.port}/api/staff`);
    assert.equal(privateDirectory.status, 401);
  } finally {
    if (server.listening) {
      await new Promise((resolve, reject) => {
        server.close((error) => error ? reject(error) : resolve());
      });
    }
    pool.query = originalQuery;
  }
});
