# enCODE Student Calling CRM

A deployable, database-backed student outreach CRM with a shared student directory, **two individually authenticated callers**, live admin visibility, **append-only call history**, follow-ups, query resolution workflows and Excel exports.

## Start (Docker recommended)

1. Copy `.env.example` to `.env` and replace **all three passwords** with different strong random strings (at least 12 characters). Never commit `.env`.
2. Run `docker compose up -d --build`.
3. Open `http://localhost:8000` on the machine hosting Docker.
4. Sign in as `admin` with `CRM_ADMIN_PASSWORD`. Go to **Import students** and upload your CSV. Caller accounts: `caller1` / `CRM_CALLER1_PASSWORD` and `caller2` / `CRM_CALLER2_PASSWORD`.
5. Both caller accounts see the same imported students and can record unlimited new attempts. Admin sees both callers' calls and productivity.

**For multi-device / production use:** Docker Compose intentionally binds to `127.0.0.1` only. Put the app behind an **HTTPS** reverse proxy with authentication hardening and proper firewall; change the binding as appropriate. Set `CRM_COOKIE_SECURE=1` when HTTPS is in place. The application uses a persistent Docker volume; perform encrypted, scheduled backups of this volume and test restores.

## Run without Docker

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
# Edit .env, then export the three password variables and (optionally) CRM_DB_PATH.
export CRM_ADMIN_PASSWORD='your-strong-secret-here'
export CRM_CALLER1_PASSWORD='another-strong-secret-here'
export CRM_CALLER2_PASSWORD='third-strong-secret-here'
uvicorn app:app --host 127.0.0.1 --port 8000
```

## How the system behaves

- **Students:** Import a UTF-8 CSV with mandatory `student_id,name,phone`; optional `email,course,batch`. Duplicate student IDs are skipped, never silently overwritten. Caller logins share this directory. Student data is not seeded with real PII.
- **Calling attempts:** Each POST creates a *new* record with the authenticated caller and server timestamp. Both the API and SQLite triggers prevent updates/deletes to historic call rows. Neither caller nor admin receives an edit/delete call action.
- **Follow-ups:** Scheduling a follow-up creates a separate task linked to the originating call. Completion only updates follow-up state and adds an audit event; the call itself remains unchanged. The assigned caller or admin can complete it.
- **Queries:** Open/in-progress queries with descriptions can be created when logging a call. Or create a ticket separately. All ticket status transitions create immutable ticket-event history.
- **Dashboard:** Selected-period call attempts, distinct contacted students, connected attempts, due follow-ups, per-caller productivity and recent activity. Browser refreshes the dashboard every **15 seconds**, and the page reads live DB data. It is polling-based near-real-time, **not WebSocket push**.
- **Reports:** Excel downloads available for today, this week (Monday UTC), this month (UTC) or all time. Contains summary, call history, caller metrics, follow-ups and (admin-only) queries. Caller report exports show only their own calls/follow-ups; admin exports cover the team. Download manually; this version does **not** automatically email scheduled reports.
- **Access:** Admin-only student import; caller access to all student profiles/combined history; only the assigned caller can complete that follow-up. Admin has global overview. Credentials are generated via environment variables on first initialization. Session cookies are HttpOnly, SameSite=Strict, with CSRF tokens for mutations. Passwords are PBKDF2-hashed.

## Operational limitations / recommended next steps

This is an **operational prototype**, not a security-audited enterprise deployment. Before exposing over the internet or storing sensitive real student records, deploy HTTPS, enforce network/IP restrictions, perform a security review and data-protection assessment (including Indian DPDP obligations as applicable), add offsite backup/restore, monitor failed logins, and configure retention/access policies. For higher concurrency / multiple app replicas, migrate from SQLite to PostgreSQL. Password reset and account administration, automated scheduled email delivery, telephony integration, SMS/WhatsApp reminders, user-specific timezone reporting, and importing/updating existing student fields are **not included**. The frontend student selector loads the latest 500 records; use student search to open the profile and log a call for any student beyond that threshold.

## Data integrity notes

A SQLite trigger rejects **UPDATE and DELETE** on `calls` even with direct SQL. Immutability of call records is application-level and database-trigger-level, not cryptographic/WORM storage; a privileged operator with disk access can alter the entire database, so backups and access control remain essential. Ticket events and audit events also reject update/delete. Student master records are deliberately import-insert-only. All server timestamps use UTC.

## Quick checks

```bash
python -m pytest -q
```

Test coverage verifies multiple attempts on one student, distinct/total counts, caller-specific identity, DB trigger immutability, follow-up permissions, and Excel file generation.
