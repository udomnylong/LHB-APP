const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function createAppsScriptContext(userRows) {
  const headers = [
    'Username', 'Password', 'Name', 'Role', 'Email',
    'Department', 'Position', 'SessionToken', 'SessionTime',
  ];
  const rows = userRows.map((row) => [...row]);
  const cache = new Map();
  const sheet = {
    getLastColumn: () => headers.length,
    getLastRow: () => rows.length + 1,
    getRange(row, column, rowCount, columnCount) {
      return {
        getValues() {
          if (row === 1) return [headers.slice(column - 1, column - 1 + columnCount)];
          return rows.slice(row - 2, row - 2 + rowCount)
            .map((values) => values.slice(column - 1, column - 1 + columnCount));
        },
        setValue(value) {
          const target = rows[row - 2];
          target[column - 1] = value;
        },
      };
    },
  };
  const context = {
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => '' }) },
    SpreadsheetApp: { openById: () => ({ getSheetByName: () => sheet }) },
    CacheService: {
      getScriptCache: () => ({
        get: (key) => cache.get(key) || null,
        put: (key, value) => cache.set(key, value),
        remove: (key) => cache.delete(key),
      }),
    },
    LockService: {
      getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }),
    },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput(content) {
        return { getContent: () => content, setMimeType() { return this; } };
      },
    },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' },
      Charset: { UTF_8: 'utf8' },
      computeDigest: (_algorithm, value, encoding) =>
        [...crypto.createHash('sha256').update(value, encoding).digest()],
      base64EncodeWebSafe: (value) => Buffer.from(value).toString('base64url'),
      getUuid: () => crypto.randomUUID(),
    },
    Logger: { log: () => {} },
    Date,
  };
  vm.createContext(context);
  const source = fs.readFileSync(
    path.join(__dirname, '..', '..', 'Code.gs'),
    'utf8'
  );
  vm.runInContext(source, context);
  return { context, rows };
}

function postAuth(context, action, values) {
  const response = context.doPost({
    postData: { contents: JSON.stringify({ action, ...values }) },
  });
  return JSON.parse(response.getContent());
}

test('HR login uses Apps Script authentication instead of Cloud Run login', () => {
  const projectRoot = path.join(__dirname, '..', '..');
  const hrPage = fs.readFileSync(path.join(projectRoot, 'hr-system.html'), 'utf8');
  const loginStart = hrPage.indexOf('async function doLoginInner');
  const loginEnd = hrPage.indexOf("\ndocument.getElementById('loginPass')", loginStart);
  assert.notEqual(loginStart, -1);
  assert.notEqual(loginEnd, -1);
  const loginBody = hrPage.slice(loginStart, loginEnd);
  assert.match(loginBody, /appsScriptAuth\('authLogin'/);
  assert.doesNotMatch(loginBody, /LhbApi\.login\s*\(/);
  assert.doesNotMatch(loginBody, /u\s*===\s*['"]admin['"].*8285/);
});

test('Apps Script login verifies sheet credentials and stores only a session token hash', () => {
  const { context, rows } = createAppsScriptContext([
    ['alice', 'correct-password', 'Alice', 'hr', 'alice@example.com', 'HR', 'Officer', '', ''],
  ]);
  const result = postAuth(context, 'authLogin', { username: 'Alice', password: 'correct-password' });

  assert.equal(result.status, 'ok');
  assert.equal(result.user.username, 'alice');
  assert.equal(result.user.role, 'hr');
  assert.equal('password' in result.user, false);
  assert.ok(result.token);
  assert.notEqual(rows[0][7], result.token);
  assert.equal(rows[0][7].length, 43);
});

test('Apps Script login rejects wrong passwords and rate-limits repeated failures', () => {
  const { context } = createAppsScriptContext([
    ['alice', 'correct-password', 'Alice', 'hr', '', '', '', '', ''],
  ]);
  for (let attempt = 0; attempt < 10; attempt++) {
    const result = postAuth(context, 'authLogin', { username: 'alice', password: 'wrong' });
    assert.equal(result.code, 'INVALID_CREDENTIALS');
  }
  const limited = postAuth(context, 'authLogin', { username: 'alice', password: 'wrong' });
  assert.equal(limited.code, 'RATE_LIMITED');
});

test('Apps Script sessions refresh, prevent concurrent logins, and revoke on logout', () => {
  const { context } = createAppsScriptContext([
    ['alice', 'correct-password', 'Alice', 'hr', '', '', '', '', ''],
  ]);
  const login = postAuth(context, 'authLogin', { username: 'alice', password: 'correct-password' });
  assert.equal(login.status, 'ok');

  const concurrent = postAuth(context, 'authLogin', { username: 'alice', password: 'correct-password' });
  assert.equal(concurrent.code, 'CONCURRENT_SESSION');

  const session = postAuth(context, 'authSession', { token: login.token });
  assert.equal(session.status, 'ok');

  assert.equal(postAuth(context, 'authLogout', { token: login.token }).status, 'ok');
  assert.equal(postAuth(context, 'authSession', { token: login.token }).code, 'SESSION_INVALID');
});
