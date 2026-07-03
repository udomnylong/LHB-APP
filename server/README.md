# LHB HR — server

Phase 1 of the Sheets → Cloud SQL migration (see the migration plan). This directory currently holds the target schema and a one-time ETL script; the Cloud Run API (Phase 2+) lands here later.

## Setup

```
cd server
npm install
cp .env.example .env   # fill in SHEET_ID, GOOGLE_APPLICATION_CREDENTIALS, DATABASE_URL
```

You need:
- A Cloud SQL for PostgreSQL instance (or any reachable Postgres for local testing), reachable at `DATABASE_URL`.
- A GCP service account with the Sheets API enabled, JSON key saved locally, and **the Google Sheet shared with that service account's email as Viewer**.

## Run

```
npm run db:schema     # creates all tables (safe to re-run)
npm run migrate:etl   # pulls all 12 sheets into Postgres (safe to re-run — full refresh per table)
```

The ETL prints a per-table row count summary and a punch list of any orphaned rows (e.g. a CheckIn row whose staff ID doesn't exist in StaffInfo) that were skipped or nulled out — resolve those in the source sheet and re-run before moving on to Phase 2.
