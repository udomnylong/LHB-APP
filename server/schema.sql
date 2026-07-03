-- LHB HR System — Cloud SQL for PostgreSQL schema
-- Phase 1 of the Sheets -> Cloud SQL migration (see plan: mossy-toasting-treasure).
-- Additive-only after this point — new columns get added in later phases, not altered/removed.

-- ============================================================
-- Auth / sessions  (replaces User sheet + SessionToken/SessionTime columns)
-- ============================================================
CREATE TABLE IF NOT EXISTS users (
  id             BIGSERIAL PRIMARY KEY,
  username       VARCHAR(100) NOT NULL UNIQUE,
  password_hash  TEXT,                 -- NULL until Phase 5 lazy-migrates the plaintext value
  password_plain TEXT,                 -- temporary carry-over from Sheets; cleared once password_hash is set (Phase 5)
  name           TEXT,
  role           VARCHAR(30) NOT NULL DEFAULT 'user',
  email          TEXT,
  department     TEXT,
  position       TEXT,
  created_at     TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at     TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS user_sessions (
  id             BIGSERIAL PRIMARY KEY,
  user_id        BIGINT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  session_token  TEXT NOT NULL UNIQUE,
  created_at     TIMESTAMPTZ NOT NULL DEFAULT now(),
  last_seen_at   TIMESTAMPTZ NOT NULL DEFAULT now(),
  expires_at     TIMESTAMPTZ NOT NULL,
  revoked_at     TIMESTAMPTZ
);
CREATE INDEX IF NOT EXISTS idx_user_sessions_active ON user_sessions (user_id, revoked_at);

-- ============================================================
-- Staff directory  (StaffInfo sheet)
-- ============================================================
CREATE TABLE IF NOT EXISTS staff (
  staff_code         VARCHAR(50) PRIMARY KEY,   -- Sheet 'ID' column, kept as natural key
  name               TEXT,
  name_latin         TEXT,
  sex                VARCHAR(10),
  lv                 TEXT,
  position           TEXT,
  department         TEXT,
  project_name       TEXT,
  date_of_birth      DATE,
  starting_date      DATE,
  resign_date        DATE,
  salary             NUMERIC(14,2),
  gmail              TEXT,
  bank_name          TEXT,
  bank_number        TEXT,
  photo_url          TEXT,               -- legacy Drive URL or new GCS URL, scheme-agnostic
  phone              TEXT,
  employment_status  TEXT,
  telegram_chat_id   TEXT,
  created_at         TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at         TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_staff_gmail ON staff (gmail);
CREATE INDEX IF NOT EXISTS idx_staff_phone ON staff (phone);

CREATE TABLE IF NOT EXISTS staff_otp (
  id          BIGSERIAL PRIMARY KEY,
  staff_code  VARCHAR(50) NOT NULL REFERENCES staff(staff_code) ON DELETE CASCADE,
  otp_code    VARCHAR(10) NOT NULL,
  expires_at  TIMESTAMPTZ NOT NULL,
  used_at     TIMESTAMPTZ,
  created_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_staff_otp_staff ON staff_otp (staff_code, expires_at);

-- ============================================================
-- Attendance summary  (Attendance sheet — one row per staff per day)
-- ============================================================
CREATE TABLE IF NOT EXISTS attendance (
  id          BIGSERIAL PRIMARY KEY,
  staff_code  VARCHAR(50) NOT NULL REFERENCES staff(staff_code) ON DELETE CASCADE,
  date        DATE NOT NULL,
  check_in    TIME,
  check_out   TIME,
  late        TEXT,
  early       TEXT,
  status      TEXT,
  created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (staff_code, date)
);

-- ============================================================
-- Raw check-in / check-out event log  (CheckIn / CheckOut sheets — many rows/day allowed)
-- ============================================================
CREATE TABLE IF NOT EXISTS check_ins (
  id               BIGSERIAL PRIMARY KEY,
  staff_code       VARCHAR(50) NOT NULL REFERENCES staff(staff_code) ON DELETE CASCADE,
  project_name     TEXT,
  event_date       DATE NOT NULL,
  event_time       TIME NOT NULL,
  event_timestamp  TIMESTAMPTZ NOT NULL,
  latitude         NUMERIC(9,6),
  longitude        NUMERIC(9,6),
  accuracy         NUMERIC,
  late_early       TEXT,
  minutes          INT,
  position         TEXT,
  department       TEXT,
  gps_verified     BOOLEAN,             -- set by Phase 5 server-side geofence re-check; NULL until then
  created_at       TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_check_ins_staff_date ON check_ins (staff_code, event_date);

CREATE TABLE IF NOT EXISTS check_outs (
  id               BIGSERIAL PRIMARY KEY,
  staff_code       VARCHAR(50) NOT NULL REFERENCES staff(staff_code) ON DELETE CASCADE,
  project_name     TEXT,
  event_date       DATE NOT NULL,
  event_time       TIME NOT NULL,
  event_timestamp  TIMESTAMPTZ NOT NULL,
  latitude         NUMERIC(9,6),
  longitude        NUMERIC(9,6),
  accuracy         NUMERIC,
  late_early       TEXT,
  minutes          INT,
  position         TEXT,
  department       TEXT,
  gps_verified     BOOLEAN,
  created_at       TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_check_outs_staff_date ON check_outs (staff_code, event_date);

-- ============================================================
-- Leave  (StaffLeave sheet)
-- ============================================================
CREATE TABLE IF NOT EXISTS staff_leave (
  id             BIGSERIAL PRIMARY KEY,
  staff_code     VARCHAR(50) NOT NULL REFERENCES staff(staff_code) ON DELETE CASCADE,
  type_of_leave  TEXT,
  start_date     DATE,
  end_date       DATE,
  days           NUMERIC(6,2),
  reason         TEXT,
  status         TEXT,
  created_at     TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at     TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_staff_leave_staff_start ON staff_leave (staff_code, start_date);

-- ============================================================
-- Projects  (Project sheet — GPS geofence source)
-- ============================================================
CREATE TABLE IF NOT EXISTS projects (
  project_id    VARCHAR(50) PRIMARY KEY,
  project_name  TEXT,
  location      TEXT,
  latitude      NUMERIC(9,6),
  longitude     NUMERIC(9,6),
  radius        NUMERIC(10,2),
  status        TEXT,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- ============================================================
-- Overtime  (StaffOT sheet)
-- ============================================================
CREATE TABLE IF NOT EXISTS staff_ot (
  id            BIGSERIAL PRIMARY KEY,
  staff_code    VARCHAR(50) NOT NULL REFERENCES staff(staff_code) ON DELETE CASCADE,
  date          DATE,
  hours         NUMERIC(6,2),
  time_from     TIME,
  time_to       TIME,
  type_of_work  TEXT,
  reason        TEXT,
  status        TEXT,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_staff_ot_staff_date ON staff_ot (staff_code, date);

-- ============================================================
-- Food  (Food sheet — kiosk, no login, staff_code nullable)
-- ============================================================
CREATE TABLE IF NOT EXISTS food_records (
  id                 BIGSERIAL PRIMARY KEY,
  date               DATE,
  staff_code         VARCHAR(50) REFERENCES staff(staff_code) ON DELETE SET NULL,
  name               TEXT,             -- snapshot at time of record, staff_code may be null
  sex                VARCHAR(10),
  position           TEXT,
  project_name       TEXT,
  morning            TEXT,
  lunch              TEXT,
  evening            TEXT,
  total              NUMERIC(6,2),
  unit_price         NUMERIC(10,2),
  total_price        NUMERIC(12,2),
  photo_morning_url  TEXT,
  photo_lunch_url    TEXT,
  photo_evening_url  TEXT,
  comment            TEXT,
  remark             TEXT,
  created_at         TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_food_records_staff_date ON food_records (staff_code, date);

-- ============================================================
-- Work place reports & comments  (WorkPlace / Comment sheets — same shape)
-- ============================================================
CREATE TABLE IF NOT EXISTS work_place_reports (
  id            BIGSERIAL PRIMARY KEY,
  date          DATE,
  time          TIME,
  staff_code    VARCHAR(50) REFERENCES staff(staff_code) ON DELETE SET NULL,
  department    TEXT,
  project_name  TEXT,
  comment       TEXT,
  photo_url     TEXT,
  status        TEXT,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_work_place_reports_staff_date ON work_place_reports (staff_code, date);

CREATE TABLE IF NOT EXISTS comments (
  id            BIGSERIAL PRIMARY KEY,
  date          DATE,
  time          TIME,
  staff_code    VARCHAR(50) REFERENCES staff(staff_code) ON DELETE SET NULL,
  department    TEXT,
  project_name  TEXT,
  comment       TEXT,
  photo_url     TEXT,
  status        TEXT,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_comments_staff_date ON comments (staff_code, date);

-- ============================================================
-- Staff evaluations  (EvaluateStaff sheet)
-- ============================================================
CREATE TABLE IF NOT EXISTS staff_evaluations (
  id               BIGSERIAL PRIMARY KEY,
  request_by       TEXT,
  staff_code       VARCHAR(50) REFERENCES staff(staff_code) ON DELETE SET NULL,
  staff_name       TEXT,               -- snapshot at time of evaluation
  date_evaluate    DATE,
  kpi_score        NUMERIC(6,2),
  previous_salary  NUMERIC(14,2),
  current_salary   NUMERIC(14,2),
  approved_by      TEXT,
  remark           TEXT,
  created_at       TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_staff_evaluations_staff ON staff_evaluations (staff_code);
