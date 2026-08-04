// Shared Cloud Run API client for hr-system.html / staff-portal.html.
// All 12 sheets are now migrated onto Cloud Run/Postgres (see LhbApi.MIGRATED_SHEETS).
// The Apps Script URL each HTML file still has is now used only for Google Drive photo
// upload/replace calls, not for any data read/write.
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

  // ── StaffLeave / StaffOT / Project sheet-header <-> API snake_case field mapping ──
  // Same purpose as STAFF_FIELD_MAP. Leave/OT also carry a RecordId (the real numeric
  // Postgres PK) so callers can key edit/delete on it instead of the old ID+Date /
  // ID-only composite matching (ID here is the *staff* ID, not a unique row ID).
  const LEAVE_FIELD_MAP = {
    ID: 'staff_code', TypeOfLeave: 'type_of_leave', StartDate: 'start_date',
    EndDate: 'end_date', Days: 'days', Reason: 'reason', Status: 'status',
  };
  const OT_FIELD_MAP = {
    ID: 'staff_code', Date: 'date', Hours: 'hours', TypeOfWork: 'type_of_work',
    Status: 'status', Remark: 'remark',
  };
  const PROJECT_FIELD_MAP = {
    ProjectID: 'project_id', ProjectName: 'project_name', Location: 'location',
    Latitude: 'latitude', Longitude: 'longitude', Radius: 'radius', Status: 'status',
  };
  const FOOD_FIELD_MAP = {
    Date: 'date', ID: 'staff_code', Name: 'name', Sex: 'sex', Position: 'position', ProjectName: 'project_name',
    Morning: 'morning', Lunch: 'lunch', Evening: 'evening', Total: 'total', UnitPrice: 'unit_price', TotalPrice: 'total_price',
    PhotoMorning: 'photo_morning_url', PhotoLunch: 'photo_lunch_url', PhotoEvening: 'photo_evening_url',
    Comment: 'comment', Remark: 'remark',
  };
  // Comment / WorkPlace sheets share the same shape.
  const REPORT_FIELD_MAP = {
    Date: 'date', Time: 'time', ID: 'staff_code', Department: 'department',
    ProjectName: 'project_name', Comment: 'comment', Photo: 'photo_url', Status: 'status',
  };
  const EVALUATE_FIELD_MAP = {
    RequestBy: 'request_by', StaffID: 'staff_code', StaffName: 'staff_name', DateEvaluate: 'date_evaluate',
    KPIScore: 'kpi_score', PreviousSalary: 'previous_salary', CurrentSalary: 'current_salary',
    ApprovedBy: 'approved_by', Remark: 'remark',
  };

  function mapRowToApi(fieldMap, row) {
    const out = {};
    Object.keys(fieldMap).forEach((sheetKey) => {
      if (row[sheetKey] !== undefined) out[fieldMap[sheetKey]] = row[sheetKey] === '' ? null : row[sheetKey];
    });
    return out;
  }
  function mapApiToRow(fieldMap, apiRow, extra) {
    const out = {};
    Object.keys(fieldMap).forEach((sheetKey) => {
      const v = apiRow[fieldMap[sheetKey]];
      out[sheetKey] = v === null || v === undefined ? '' : String(v);
    });
    if (apiRow.staff_name !== undefined) out.Name = apiRow.staff_name || '';
    if (extra) Object.assign(out, extra);
    return out;
  }

  const LhbApi = {
    // Resources fully migrated off Apps Script as of Phase 2 — callers use this
    // to decide whether to route a given sheet name through LhbApi or the old URL.
    MIGRATED_SHEETS: ['User', 'StaffInfo', 'CheckIn', 'CheckOut', 'Attendance', 'StaffLeave', 'StaffOT', 'Project', 'Food', 'Comment', 'Workplace', 'EvaluateStaff'],

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

    // ── Admin: back up Cloud SQL data for the 8 migrated resources into the Sheet ──
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
    // No auth needed — staff self-service profile photo update (see server/routes/staff.js).
    async updateStaffPhoto(staffCode, photoUrl) {
      return apiFetch('/api/staff/' + encodeURIComponent(staffCode) + '/photo', {
        method: 'PUT', body: JSON.stringify({ photoUrl }),
      });
    },

    // ── Leave (sheet-shaped in/out, see LEAVE_FIELD_MAP) ──
    async getLeave(params) {
      const qs = new URLSearchParams(params || {}).toString();
      const r = await apiFetch('/api/leave' + (qs ? '?' + qs : ''), { method: 'GET' });
      if (r._ok && Array.isArray(r.data)) r.data = r.data.map((row) => mapApiToRow(LEAVE_FIELD_MAP, row, { RecordId: String(row.id) }));
      return r;
    },
    async createLeave(sheetRow) {
      return apiFetch('/api/leave', { method: 'POST', body: JSON.stringify(mapRowToApi(LEAVE_FIELD_MAP, sheetRow)) });
    },
    async updateLeave(recordId, sheetRow) {
      return apiFetch('/api/leave/' + encodeURIComponent(recordId), { method: 'PUT', body: JSON.stringify(mapRowToApi(LEAVE_FIELD_MAP, sheetRow)) });
    },
    async deleteLeave(recordId) {
      return apiFetch('/api/leave/' + encodeURIComponent(recordId), { method: 'DELETE' });
    },

    // ── OT (sheet-shaped in/out, see OT_FIELD_MAP) ──
    async getOT(params) {
      const qs = new URLSearchParams(params || {}).toString();
      const r = await apiFetch('/api/ot' + (qs ? '?' + qs : ''), { method: 'GET' });
      if (r._ok && Array.isArray(r.data)) r.data = r.data.map((row) => mapApiToRow(OT_FIELD_MAP, row, { RecordId: String(row.id) }));
      return r;
    },
    async createOT(sheetRow) {
      return apiFetch('/api/ot', { method: 'POST', body: JSON.stringify(mapRowToApi(OT_FIELD_MAP, sheetRow)) });
    },
    async updateOT(recordId, sheetRow) {
      return apiFetch('/api/ot/' + encodeURIComponent(recordId), { method: 'PUT', body: JSON.stringify(mapRowToApi(OT_FIELD_MAP, sheetRow)) });
    },
    async deleteOT(recordId) {
      return apiFetch('/api/ot/' + encodeURIComponent(recordId), { method: 'DELETE' });
    },

    // ── Projects (sheet-shaped in/out, see PROJECT_FIELD_MAP; project_id is the natural key) ──
    async getProjects() {
      const r = await apiFetch('/api/projects', { method: 'GET' });
      if (r._ok && Array.isArray(r.data)) r.data = r.data.map((row) => mapApiToRow(PROJECT_FIELD_MAP, row));
      return r;
    },
    async createProject(sheetRow) {
      return apiFetch('/api/projects', { method: 'POST', body: JSON.stringify(mapRowToApi(PROJECT_FIELD_MAP, sheetRow)) });
    },
    async updateProject(projectId, sheetRow) {
      return apiFetch('/api/projects/' + encodeURIComponent(projectId), { method: 'PUT', body: JSON.stringify(mapRowToApi(PROJECT_FIELD_MAP, sheetRow)) });
    },
    async deleteProject(projectId) {
      return apiFetch('/api/projects/' + encodeURIComponent(projectId), { method: 'DELETE' });
    },

    // ── Food (sheet-shaped in/out, see FOOD_FIELD_MAP; photos stay on Google Drive —
    // callers pass the resulting Drive URL in PhotoMorning/PhotoLunch/PhotoEvening) ──
    async getFood(params) {
      const qs = new URLSearchParams(params || {}).toString();
      const r = await apiFetch('/api/food' + (qs ? '?' + qs : ''), { method: 'GET' });
      if (r._ok && Array.isArray(r.data)) r.data = r.data.map((row) => mapApiToRow(FOOD_FIELD_MAP, row, { RecordId: String(row.id) }));
      return r;
    },
    async createFood(sheetRow) {
      return apiFetch('/api/food', { method: 'POST', body: JSON.stringify(mapRowToApi(FOOD_FIELD_MAP, sheetRow)) });
    },
    async updateFood(recordId, sheetRow) {
      return apiFetch('/api/food/' + encodeURIComponent(recordId), { method: 'PUT', body: JSON.stringify(mapRowToApi(FOOD_FIELD_MAP, sheetRow)) });
    },
    async deleteFood(recordId) {
      return apiFetch('/api/food/' + encodeURIComponent(recordId), { method: 'DELETE' });
    },

    // ── Comment / WorkPlace reports (sheet-shaped in/out, see REPORT_FIELD_MAP) ──
    async getComments(params) {
      const qs = new URLSearchParams(params || {}).toString();
      const r = await apiFetch('/api/comments' + (qs ? '?' + qs : ''), { method: 'GET' });
      if (r._ok && Array.isArray(r.data)) r.data = r.data.map((row) => mapApiToRow(REPORT_FIELD_MAP, row, { RecordId: String(row.id) }));
      return r;
    },
    async createComment(sheetRow) {
      return apiFetch('/api/comments', { method: 'POST', body: JSON.stringify(mapRowToApi(REPORT_FIELD_MAP, sheetRow)) });
    },
    async updateComment(recordId, sheetRow) {
      return apiFetch('/api/comments/' + encodeURIComponent(recordId), { method: 'PUT', body: JSON.stringify(mapRowToApi(REPORT_FIELD_MAP, sheetRow)) });
    },
    async deleteComment(recordId) {
      return apiFetch('/api/comments/' + encodeURIComponent(recordId), { method: 'DELETE' });
    },
    async getWorkplace(params) {
      const qs = new URLSearchParams(params || {}).toString();
      const r = await apiFetch('/api/workplace' + (qs ? '?' + qs : ''), { method: 'GET' });
      if (r._ok && Array.isArray(r.data)) r.data = r.data.map((row) => mapApiToRow(REPORT_FIELD_MAP, row, { RecordId: String(row.id) }));
      return r;
    },
    async createWorkplace(sheetRow) {
      return apiFetch('/api/workplace', { method: 'POST', body: JSON.stringify(mapRowToApi(REPORT_FIELD_MAP, sheetRow)) });
    },
    async updateWorkplace(recordId, sheetRow) {
      return apiFetch('/api/workplace/' + encodeURIComponent(recordId), { method: 'PUT', body: JSON.stringify(mapRowToApi(REPORT_FIELD_MAP, sheetRow)) });
    },
    async deleteWorkplace(recordId) {
      return apiFetch('/api/workplace/' + encodeURIComponent(recordId), { method: 'DELETE' });
    },

    // ── Staff evaluations (sheet-shaped in/out, see EVALUATE_FIELD_MAP; the old
    // client-generated EvalNo label is gone — RecordId, the real Postgres id, is the key) ──
    async getEvaluations(params) {
      const qs = new URLSearchParams(params || {}).toString();
      const r = await apiFetch('/api/evaluations' + (qs ? '?' + qs : ''), { method: 'GET' });
      if (r._ok && Array.isArray(r.data)) r.data = r.data.map((row) => mapApiToRow(EVALUATE_FIELD_MAP, row, { RecordId: String(row.id) }));
      return r;
    },
    async createEvaluation(sheetRow) {
      return apiFetch('/api/evaluations', { method: 'POST', body: JSON.stringify(mapRowToApi(EVALUATE_FIELD_MAP, sheetRow)) });
    },
    async updateEvaluation(recordId, sheetRow) {
      return apiFetch('/api/evaluations/' + encodeURIComponent(recordId), { method: 'PUT', body: JSON.stringify(mapRowToApi(EVALUATE_FIELD_MAP, sheetRow)) });
    },
    async deleteEvaluation(recordId) {
      return apiFetch('/api/evaluations/' + encodeURIComponent(recordId), { method: 'DELETE' });
    },

    // ── Attendance (real-time check-in/out; manualCheckIn/manualCheckOut below cover
    // admin backfill/correction edits) ──
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
    // Admin backfill/correction of a specific date+time — requireAuth-gated server-side
    // (see server/routes/attendance.js). Used by hr-system.html's manual attendance edit
    // forms (tpSave/saveManualAtt), not by staff-portal.html's real-time check-in.
    async manualCheckIn({ staffCode, projectName, date, time }) {
      return apiFetch('/api/checkins/manual', {
        method: 'POST',
        body: JSON.stringify({ staff_code: staffCode, project_name: projectName, date, time }),
      });
    },
    async manualCheckOut({ staffCode, projectName, date, time }) {
      return apiFetch('/api/checkouts/manual', {
        method: 'POST',
        body: JSON.stringify({ staff_code: staffCode, project_name: projectName, date, time }),
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
