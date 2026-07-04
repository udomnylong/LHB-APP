// Shared Cloud Run API client for hr-system.html / staff-portal.html.
// Phase 2 of the Sheets -> Cloud SQL migration: only auth, staff, and
// check-in/out+attendance are migrated here. Everything else still goes
// through the Apps Script URL each HTML file already has.
//
// Override with window.LHB_API_BASE (set via a <script> tag before this file
// loads) for local dev/testing against a different backend.
(function (global) {
  const API_BASE = global.LHB_API_BASE || 'https://lhb-hr-api-860256256963.asia-southeast1.run.app';
  const TOKEN_KEY = 'lhbApiToken';

  function getToken() { return sessionStorage.getItem(TOKEN_KEY) || ''; }
  function setToken(t) { if (t) sessionStorage.setItem(TOKEN_KEY, t); else sessionStorage.removeItem(TOKEN_KEY); }

  async function apiFetch(path, options) {
    options = options || {};
    const headers = Object.assign({ 'Content-Type': 'application/json' }, options.headers || {});
    const token = getToken();
    if (token) headers.Authorization = 'Bearer ' + token;

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), options.timeoutMs || 15000);
    try {
      const res = await fetch(API_BASE + path, { ...options, headers, signal: controller.signal });
      const text = await res.text();
      let json;
      try { json = text ? JSON.parse(text) : {}; }
      catch (e) { json = { status: 'error', msg: 'Invalid JSON response from API' }; }
      json._httpStatus = res.status;
      json._ok = res.ok;
      return json;
    } catch (e) {
      return { status: 'error', msg: e.name === 'AbortError' ? 'Request timed out' : e.message, _ok: false, _httpStatus: 0 };
    } finally {
      clearTimeout(timer);
    }
  }

  // ── StaffInfo sheet-header <-> API snake_case field mapping ──
  // Keeps every existing hr-system.html / staff-portal.html call site (which reads/writes
  // row.ID, row.Name, row.ProjectName, ...) working unchanged against the new API.
  const STAFF_FIELD_MAP = {
    ID: 'staff_code', Name: 'name', NameLatin: 'name_latin', Sex: 'sex', LV: 'lv',
    Position: 'position', Department: 'department', ProjectName: 'project_name',
    DateOfBirth: 'date_of_birth', StartingDate: 'starting_date', ResignDate: 'resign_date',
    Salary: 'salary', Gmail: 'gmail', BankName: 'bank_name', BankNumber: 'bank_number',
    Photo: 'photo_url', Phone: 'phone', EmploymentStatus: 'employment_status',
    TelegramChatId: 'telegram_chat_id',
  };

  function staffRowToApi(row) {
    const out = {};
    Object.keys(STAFF_FIELD_MAP).forEach((sheetKey) => {
      if (row[sheetKey] !== undefined) out[STAFF_FIELD_MAP[sheetKey]] = row[sheetKey] === '' ? null : row[sheetKey];
    });
    return out;
  }

  function staffApiToRow(apiRow) {
    const out = {};
    Object.keys(STAFF_FIELD_MAP).forEach((sheetKey) => {
      const v = apiRow[STAFF_FIELD_MAP[sheetKey]];
      out[sheetKey] = v === null || v === undefined ? '' : String(v);
    });
    return out;
  }

  const LhbApi = {
    // Resources fully migrated off Apps Script as of Phase 2 — callers use this
    // to decide whether to route a given sheet name through LhbApi or the old URL.
    MIGRATED_SHEETS: ['User', 'StaffInfo', 'CheckIn', 'CheckOut', 'Attendance'],

    // ── Auth ──
    async login(username, password) {
      const r = await apiFetch('/api/auth/login', { method: 'POST', body: JSON.stringify({ username, password }) });
      if (r._ok && r.token) setToken(r.token);
      return r;
    },
    async logout() {
      const r = await apiFetch('/api/auth/logout', { method: 'POST' });
      setToken(null);
      return r;
    },
    // Best-effort logout for beforeunload/pagehide — keepalive fetch, response not awaited.
    logoutBeacon() {
      const token = getToken();
      if (!token) return;
      try {
        fetch(API_BASE + '/api/auth/logout', {
          method: 'POST', keepalive: true, headers: { Authorization: 'Bearer ' + token },
        });
      } catch (e) { /* best effort */ }
      setToken(null);
    },
    async getSession() {
      return apiFetch('/api/auth/session', { method: 'GET' });
    },
    hasToken() { return !!getToken(); },
    clearToken() { setToken(null); },

    // ── Admin: back up Cloud SQL data for the 5 migrated resources into the Sheet ──
    async backupToSheets() {
      return apiFetch('/api/admin/backup-to-sheets', { method: 'POST', timeoutMs: 60000 });
    },

    // ── Users (Settings screen list — no password field, see server/routes/users.js) ──
    async getUsers() {
      const r = await apiFetch('/api/users', { method: 'GET' });
      if (r._ok && Array.isArray(r.data)) {
        r.data = r.data.map((u) => ({
          Username: u.username, Name: u.name, Role: u.role,
          Email: u.email || '', Department: u.department || '', Position: u.position || '',
        }));
      }
      return r;
    },

    // ── Staff (sheet-shaped in/out — see staffRowToApi/staffApiToRow) ──
    async getStaff() {
      const r = await apiFetch('/api/staff', { method: 'GET' });
      if (r._ok && Array.isArray(r.data)) r.data = r.data.map(staffApiToRow);
      return r;
    },
    async createStaff(sheetRow) {
      return apiFetch('/api/staff', { method: 'POST', body: JSON.stringify(staffRowToApi(sheetRow)) });
    },
    async updateStaff(staffCode, sheetRow) {
      return apiFetch('/api/staff/' + encodeURIComponent(staffCode), { method: 'PUT', body: JSON.stringify(staffRowToApi(sheetRow)) });
    },
    async deleteStaff(staffCode) {
      return apiFetch('/api/staff/' + encodeURIComponent(staffCode), { method: 'DELETE' });
    },

    // ── Attendance (real-time check-in/out; NOT for manual/backfill admin edits —
    // those stay on Apps Script, see the migration plan) ──
    async checkIn({ staffCode, projectName, latitude, longitude, accuracy }) {
      return apiFetch('/api/checkins', {
        method: 'POST',
        body: JSON.stringify({ staff_code: staffCode, project_name: projectName, latitude, longitude, accuracy }),
      });
    },
    async checkOut({ staffCode, projectName, latitude, longitude, accuracy }) {
      return apiFetch('/api/checkouts', {
        method: 'POST',
        body: JSON.stringify({ staff_code: staffCode, project_name: projectName, latitude, longitude, accuracy }),
      });
    },
    async getAttendance(params) {
      const qs = new URLSearchParams(params || {}).toString();
      return apiFetch('/api/attendance' + (qs ? '?' + qs : ''), { method: 'GET' });
    },
    // Sheet-shaped Attendance rows (ID/Name/Position/Department/ProjectName/CheckIn/CheckOut/Late/Early/Status/Date)
    // for hr-system.html's dashboards, which still read the old Attendance sheet's column names.
    async getAttendanceSheetRows(params) {
      const r = await this.getAttendance(params);
      if (r._ok && Array.isArray(r.data)) {
        r.data = r.data.map((a) => ({
          ID: a.staff_code, Name: a.name || '', Position: a.staff_position || '', Department: a.staff_department || '',
          ProjectName: a.project_name || '', Date: a.date, CheckIn: a.check_in || '', CheckOut: a.check_out || '',
          Late: a.late || '', Early: a.early || '', Status: a.status || '',
        }));
      }
      return r;
    },
    async getCheckEventSheetRows(kind, params) {
      const qs = new URLSearchParams(params || {}).toString();
      const path = (kind === 'checkin' ? '/api/checkins' : '/api/checkouts') + (qs ? '?' + qs : '');
      const r = await apiFetch(path, { method: 'GET' });
      if (r._ok && Array.isArray(r.data)) {
        r.data = r.data.map((e) => ({
          ID: e.staff_code, Name: e.name || '', Gmail: e.gmail || '', ProjectName: e.project_name || '',
          Date: e.event_date, Time: (e.event_time || '').slice(0, 5), Timestamp: e.event_timestamp,
          Latitude: e.latitude != null ? String(e.latitude) : '', Longitude: e.longitude != null ? String(e.longitude) : '',
          Accuracy: e.accuracy != null ? String(e.accuracy) : '', LateEarly: e.late_early || '',
          Minutes: e.minutes != null ? String(e.minutes) : '0', Position: e.position || '', Department: e.department || '',
        }));
      }
      return r;
    },
  };

  global.LhbApi = LhbApi;
})(window);
