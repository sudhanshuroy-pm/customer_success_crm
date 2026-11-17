"""Student calling CRM: shared directory, append-only call attempts, follow-ups, tickets, Excel reporting."""
from __future__ import annotations

import csv
import io
import json
import os
import re
import secrets
import sqlite3
import threading
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from hashlib import pbkdf2_hmac, sha256
from hmac import compare_digest
from pathlib import Path
from typing import Any

from fastapi import Depends, FastAPI, HTTPException, Request, Response
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel, Field
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
DB_PATH = Path(os.getenv('CRM_DB_PATH', str(BASE / 'data' / 'crm.sqlite3')))
COOKIE_SECURE = os.getenv('CRM_COOKIE_SECURE', '0') == '1'
SESSION_HOURS = int(os.getenv('CRM_SESSION_HOURS', '12'))
app = FastAPI(title='Student Calling CRM', docs_url=None, redoc_url=None)
LOCK = threading.RLock()


def now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec='seconds')


def hash_password(password: str, salt: str | None = None) -> str:
    salt = salt or secrets.token_hex(16)
    digest = pbkdf2_hmac('sha256', password.encode(), bytes.fromhex(salt), 260000).hex()
    return f'{salt}${digest}'


def verify_password(password: str, saved: str) -> bool:
    try:
        salt, _ = saved.split('$', 1)
        return compare_digest(hash_password(password, salt), saved)
    except (ValueError, TypeError):
        return False


@contextmanager
def db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(DB_PATH), timeout=20, isolation_level=None)
    con.row_factory = sqlite3.Row
    con.execute('PRAGMA foreign_keys=ON')
    con.execute('PRAGMA busy_timeout=20000')
    try:
        yield con
    finally:
        con.close()


def initialize():
    with LOCK, db() as con:
        con.executescript('''
        PRAGMA journal_mode=WAL;
        CREATE TABLE IF NOT EXISTS users (
          id INTEGER PRIMARY KEY, username TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
          full_name TEXT NOT NULL, role TEXT NOT NULL CHECK(role IN ('admin','caller')),
          active INTEGER NOT NULL DEFAULT 1);
        CREATE TABLE IF NOT EXISTS sessions (
          token_hash TEXT PRIMARY KEY, user_id INTEGER NOT NULL REFERENCES users(id),
          csrf_token TEXT NOT NULL, expires_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS students (
          id INTEGER PRIMARY KEY, external_id TEXT UNIQUE NOT NULL, name TEXT NOT NULL,
          phone TEXT NOT NULL, email TEXT DEFAULT '', course TEXT DEFAULT '', batch TEXT DEFAULT '',
          created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS calls (
          id INTEGER PRIMARY KEY, student_id INTEGER NOT NULL REFERENCES students(id),
          caller_id INTEGER NOT NULL REFERENCES users(id),
          called_at TEXT NOT NULL, outcome TEXT NOT NULL, student_response TEXT NOT NULL,
          notes TEXT NOT NULL DEFAULT '', query_status TEXT NOT NULL DEFAULT 'No query',
          query_details TEXT NOT NULL DEFAULT '');
        CREATE TRIGGER IF NOT EXISTS calls_no_update BEFORE UPDATE ON calls
          BEGIN SELECT RAISE(ABORT, 'Call history is immutable'); END;
        CREATE TRIGGER IF NOT EXISTS calls_no_delete BEFORE DELETE ON calls
          BEGIN SELECT RAISE(ABORT, 'Call history cannot be deleted'); END;
        CREATE TABLE IF NOT EXISTS followups (
          id INTEGER PRIMARY KEY, source_call_id INTEGER NOT NULL REFERENCES calls(id),
          student_id INTEGER NOT NULL REFERENCES students(id),
          assigned_user_id INTEGER NOT NULL REFERENCES users(id),
          due_at TEXT NOT NULL, completed_at TEXT, completion_note TEXT DEFAULT '');
        CREATE TABLE IF NOT EXISTS tickets (
          id INTEGER PRIMARY KEY, student_id INTEGER NOT NULL REFERENCES students(id),
          opened_by INTEGER NOT NULL REFERENCES users(id), opened_at TEXT NOT NULL,
          title TEXT NOT NULL, description TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'Open');
        CREATE TABLE IF NOT EXISTS ticket_events (
          id INTEGER PRIMARY KEY, ticket_id INTEGER NOT NULL REFERENCES tickets(id),
          actor_id INTEGER NOT NULL REFERENCES users(id), at TEXT NOT NULL,
          status TEXT NOT NULL, note TEXT NOT NULL);
        CREATE TRIGGER IF NOT EXISTS ticket_events_no_update BEFORE UPDATE ON ticket_events
          BEGIN SELECT RAISE(ABORT, 'Ticket event history is immutable'); END;
        CREATE TRIGGER IF NOT EXISTS ticket_events_no_delete BEFORE DELETE ON ticket_events
          BEGIN SELECT RAISE(ABORT, 'Ticket event history cannot be deleted'); END;
        CREATE TABLE IF NOT EXISTS audit_events (
          id INTEGER PRIMARY KEY, actor_id INTEGER REFERENCES users(id),
          at TEXT NOT NULL, action TEXT NOT NULL, details TEXT NOT NULL);
        CREATE TRIGGER IF NOT EXISTS audit_no_update BEFORE UPDATE ON audit_events
          BEGIN SELECT RAISE(ABORT, 'Audit trail is immutable'); END;
        CREATE TRIGGER IF NOT EXISTS audit_no_delete BEFORE DELETE ON audit_events
          BEGIN SELECT RAISE(ABORT, 'Audit trail cannot be deleted'); END;
        CREATE INDEX IF NOT EXISTS calls_student_date ON calls(student_id, called_at);
        CREATE INDEX IF NOT EXISTS calls_date ON calls(called_at);
        CREATE INDEX IF NOT EXISTS followups_owner_due ON followups(assigned_user_id,due_at);
        ''')
        if con.execute('SELECT COUNT(*) FROM users').fetchone()[0] == 0:
            accounts = [('admin', 'CRM_ADMIN_PASSWORD', 'CRM Administrator', 'admin'),
                        ('caller1', 'CRM_CALLER1_PASSWORD', 'Caller One', 'caller'),
                        ('caller2', 'CRM_CALLER2_PASSWORD', 'Caller Two', 'caller')]
            for username, envvar, name, role in accounts:
                password = os.getenv(envvar)
                if not password:
                    raise RuntimeError(f'Set {envvar} before first startup (12+ characters recommended).')
                if len(password) < 12:
                    raise RuntimeError(f'{envvar} must have at least 12 characters.')
                con.execute('INSERT INTO users(username,password_hash,full_name,role) VALUES (?,?,?,?)',
                            (username, hash_password(password), name, role))


@app.on_event('startup')
def on_startup():
    initialize()


def audit(con, user_id: int, action: str, details: dict):
    con.execute('INSERT INTO audit_events(actor_id,at,action,details) VALUES(?,?,?,?)',
                (user_id, now(), action, json.dumps(details, separators=(',', ':'))))


def current_user(request: Request):
    token = request.cookies.get('crm_session', '')
    if not token:
        raise HTTPException(401, 'Please log in')
    with db() as con:
        row = con.execute('''SELECT u.id,u.username,u.full_name,u.role,u.active,s.csrf_token
           FROM sessions s JOIN users u ON u.id=s.user_id
           WHERE s.token_hash=? AND s.expires_at>?''', (sha256(token.encode()).hexdigest(), now())).fetchone()
        if not row or not row['active']:
            raise HTTPException(401, 'Session expired; please log in')
        if request.method in ('POST', 'PUT', 'PATCH', 'DELETE'):
            origin = request.headers.get('origin')
            host = request.headers.get('host', '')
            if origin and not origin.endswith('://' + host):
                raise HTTPException(403, 'Origin mismatch')
            if not compare_digest(request.headers.get('x-csrf-token', ''), row['csrf_token']):
                raise HTTPException(403, 'CSRF token missing or invalid')
        return dict(row)


def admin_only(user=Depends(current_user)):
    if user['role'] != 'admin':
        raise HTTPException(403, 'Admin access required')
    return user


class Login(BaseModel):
    username: str
    password: str


@app.post('/api/login')
def login(body: Login, response: Response):
    with db() as con:
        user = con.execute('SELECT * FROM users WHERE username=? AND active=1', (body.username.strip(),)).fetchone()
        if not user or not verify_password(body.password, user['password_hash']):
            raise HTTPException(401, 'Invalid username or password')
        token = secrets.token_urlsafe(40)
        csrf = secrets.token_urlsafe(32)
        con.execute('DELETE FROM sessions WHERE expires_at<=?', (now(),))
        con.execute('INSERT INTO sessions VALUES (?,?,?,?)',
                    (sha256(token.encode()).hexdigest(), user['id'], csrf,
                     (datetime.now(timezone.utc) + timedelta(hours=SESSION_HOURS)).isoformat(timespec='seconds')))
        audit(con, user['id'], 'login', {})
        response.set_cookie('crm_session', token, httponly=True, secure=COOKIE_SECURE,
                            samesite='strict', max_age=SESSION_HOURS * 3600, path='/')
        return {'user': {'id': user['id'], 'username': user['username'], 'name': user['full_name'], 'role': user['role']},
                'csrf': csrf}


@app.post('/api/logout')
def logout(request: Request, response: Response, user=Depends(current_user)):
    with db() as con:
        con.execute('DELETE FROM sessions WHERE token_hash=?',
                    (sha256(request.cookies['crm_session'].encode()).hexdigest(),))
        audit(con, user['id'], 'logout', {})
    response.delete_cookie('crm_session', path='/')
    return {'ok': True}


@app.get('/api/me')
def me(user=Depends(current_user)):
    return {'user': {'id': user['id'], 'username': user['username'], 'name': user['full_name'], 'role': user['role']},
            'csrf': user['csrf_token']}


@app.get('/api/users')
def users(user=Depends(current_user)):
    with db() as con:
        return [dict(r) for r in con.execute('SELECT id,username,full_name,role FROM users WHERE active=1 ORDER BY id')]


@app.get('/api/students')
def students(q: str = '', limit: int = 200, offset: int = 0, user=Depends(current_user)):
    limit, offset = min(max(limit, 1), 500), max(offset, 0)
    term = f'%{q.strip()}%'
    with db() as con:
        data = con.execute('''SELECT s.*, (SELECT COUNT(*) FROM calls WHERE student_id=s.id) AS call_count,
          (SELECT MAX(called_at) FROM calls WHERE student_id=s.id) AS last_called
          FROM students s WHERE (s.name LIKE ? OR s.phone LIKE ? OR s.external_id LIKE ? OR s.course LIKE ?)
          ORDER BY s.id DESC LIMIT ? OFFSET ?''', (term, term, term, term, limit, offset)).fetchall()
        count = con.execute('''SELECT COUNT(*) FROM students WHERE name LIKE ? OR phone LIKE ? OR external_id LIKE ? OR course LIKE ?''',
                            (term, term, term, term)).fetchone()[0]
        return {'items': [dict(r) for r in data], 'total': count}


@app.get('/api/students/{student_id}')
def student_detail(student_id: int, user=Depends(current_user)):
    with db() as con:
        s = con.execute('SELECT * FROM students WHERE id=?', (student_id,)).fetchone()
        if not s:
            raise HTTPException(404, 'Student not found')
        calls = con.execute('''SELECT c.*,u.full_name AS caller_name FROM calls c JOIN users u ON u.id=c.caller_id
                              WHERE c.student_id=? ORDER BY c.id DESC''', (student_id,)).fetchall()
        followups = con.execute('''SELECT f.*,u.full_name AS assigned_to FROM followups f
                                  JOIN users u ON u.id=f.assigned_user_id
                                  WHERE f.student_id=? ORDER BY f.due_at''', (student_id,)).fetchall()
        tickets = con.execute('SELECT * FROM tickets WHERE student_id=? ORDER BY id DESC', (student_id,)).fetchall()
        return {'student': dict(s), 'calls': [dict(r) for r in calls],
                'followups': [dict(r) for r in followups], 'tickets': [dict(r) for r in tickets]}


class ImportBody(BaseModel):
    csv_text: str = Field(min_length=5, max_length=8_000_000)


@app.post('/api/students/import')
def import_students(body: ImportBody, user=Depends(admin_only)):
    reader = csv.DictReader(io.StringIO(body.csv_text.lstrip('\ufeff')))
    required = {'student_id', 'name', 'phone'}
    if not reader.fieldnames or not required.issubset({h.strip().lower() for h in reader.fieldnames}):
        raise HTTPException(400, 'CSV requires headers: student_id,name,phone (optional email,course,batch)')
    rows = []
    for number, raw in enumerate(reader, 2):
        r = {str(k).strip().lower(): (str(v).strip() if v is not None else '') for k, v in raw.items() if k is not None}
        if not r.get('student_id') or not r.get('name') or not r.get('phone'):
            raise HTTPException(400, f'Missing student_id, name or phone at CSV row {number}')
        rows.append((r['student_id'][:80],r['name'][:160],r['phone'][:40],r.get('email','')[:160],
                     r.get('course','')[:160],r.get('batch','')[:100],now()))
    if len(rows) > 20000:
        raise HTTPException(400, 'Maximum 20,000 students per import')
    with LOCK, db() as con:
        con.execute('BEGIN IMMEDIATE')
        before = con.total_changes
        con.executemany('''INSERT OR IGNORE INTO students(external_id,name,phone,email,course,batch,created_at)
                           VALUES (?,?,?,?,?,?,?)''', rows)
        added = con.total_changes - before
        audit(con, user['id'], 'student_import', {'added': added, 'skipped_existing': len(rows)-added})
        con.execute('COMMIT')
    return {'added': added, 'skipped_existing': len(rows)-added, 'total_rows': len(rows)}


OUTCOMES = {'Connected', 'No answer', 'Busy', 'Switched off', 'Wrong number', 'Callback requested', 'Not interested'}
RESPONSES = {'Interested', 'Not interested', 'Needs information', 'Query raised', 'No response', 'Other'}
QUERY_STATUSES = {'No query', 'Open', 'In progress', 'Resolved'}


class CallInput(BaseModel):
    student_id: int
    outcome: str
    student_response: str
    notes: str = Field(default='', max_length=2000)
    query_status: str = 'No query'
    query_details: str = Field(default='', max_length=2000)
    followup_at: str | None = None


def parse_date(value: str) -> str:
    try:
        parsed = datetime.fromisoformat(value.replace('Z', '+00:00'))
        if parsed.tzinfo is None:
            raise ValueError('Timezone required')
        return parsed.astimezone(timezone.utc).isoformat(timespec='seconds')
    except ValueError:
        raise HTTPException(400, 'Invalid date/time; use an ISO datetime with timezone')


@app.post('/api/calls')
def create_call(body: CallInput, user=Depends(current_user)):
    if body.outcome not in OUTCOMES or body.student_response not in RESPONSES or body.query_status not in QUERY_STATUSES:
        raise HTTPException(400, 'Invalid call outcome, response or query status')
    followup_at = parse_date(body.followup_at) if body.followup_at else None
    with LOCK, db() as con:
        con.execute('BEGIN IMMEDIATE')
        if not con.execute('SELECT 1 FROM students WHERE id=?', (body.student_id,)).fetchone():
            raise HTTPException(404, 'Student not found')
        cur = con.execute('''INSERT INTO calls(student_id,caller_id,called_at,outcome,student_response,notes,query_status,query_details)
            VALUES (?,?,?,?,?,?,?,?)''', (body.student_id, user['id'], now(), body.outcome, body.student_response,
                                          body.notes, body.query_status, body.query_details))
        cid = cur.lastrowid
        if followup_at:
            con.execute('''INSERT INTO followups(source_call_id,student_id,assigned_user_id,due_at)
                           VALUES (?,?,?,?)''', (cid, body.student_id, user['id'], followup_at))
        if body.query_status in ('Open','In progress') and body.query_details.strip():
            t = con.execute('''INSERT INTO tickets(student_id,opened_by,opened_at,title,description,status)
                               VALUES (?,?,?,?,?,?)''', (body.student_id, user['id'], now(),
                                                         'Call query', body.query_details, body.query_status))
            con.execute('INSERT INTO ticket_events(ticket_id,actor_id,at,status,note) VALUES(?,?,?,?,?)',
                        (t.lastrowid,user['id'],now(),body.query_status,'Created from call'))
        audit(con, user['id'], 'call_created', {'call_id': cid, 'student_id': body.student_id})
        con.execute('COMMIT')
        return {'id': cid, 'ok': True}


@app.get('/api/calls')
def list_calls(limit: int = 200, caller_id: int | None = None, user=Depends(current_user)):
    limit = min(max(limit, 1), 1000)
    if user['role'] != 'admin' and caller_id and caller_id != user['id']:
        raise HTTPException(403, 'Cannot filter other caller records')
    with db() as con:
        where = 'WHERE c.caller_id=?' if caller_id else ''
        rows = con.execute(f'''SELECT c.*,s.name AS student_name,s.external_id,u.full_name AS caller_name
          FROM calls c JOIN students s ON s.id=c.student_id JOIN users u ON u.id=c.caller_id
          {where} ORDER BY c.id DESC LIMIT ?''', ((caller_id,limit) if caller_id else (limit,))).fetchall()
        return [dict(r) for r in rows]


@app.get('/api/followups')
def list_followups(status: str = 'pending', user=Depends(current_user)):
    if status not in ('pending','completed','all'):
        raise HTTPException(400, 'Invalid status')
    where = ['1=1']
    params = []
    if user['role'] != 'admin':
        where.append('f.assigned_user_id=?');params.append(user['id'])
    if status != 'all':
        where.append('f.completed_at IS NULL' if status == 'pending' else 'f.completed_at IS NOT NULL')
    with db() as con:
        rows = con.execute(f'''SELECT f.*,s.name AS student_name,s.phone,s.external_id,
          u.full_name AS assigned_to FROM followups f JOIN students s ON s.id=f.student_id
          JOIN users u ON u.id=f.assigned_user_id WHERE {' AND '.join(where)} ORDER BY f.due_at ASC''',params).fetchall()
        return [dict(r) for r in rows]


class Completion(BaseModel):
    note: str = Field(default='', max_length=1000)


@app.post('/api/followups/{fid}/complete')
def complete_followup(fid: int, body: Completion, user=Depends(current_user)):
    with LOCK, db() as con:
        con.execute('BEGIN IMMEDIATE')
        f = con.execute('SELECT * FROM followups WHERE id=?', (fid,)).fetchone()
        if not f:
            raise HTTPException(404, 'Follow-up not found')
        if user['role'] != 'admin' and f['assigned_user_id'] != user['id']:
            raise HTTPException(403, 'Only assigned caller or admin can complete')
        if f['completed_at']:
            raise HTTPException(409, 'Already completed')
        con.execute('UPDATE followups SET completed_at=?,completion_note=? WHERE id=?', (now(),body.note,fid))
        audit(con,user['id'],'followup_completed',{'followup_id':fid,'note':body.note})
        con.execute('COMMIT')
        return {'ok': True}


class TicketInput(BaseModel):
    student_id: int
    title: str = Field(min_length=2,max_length=160)
    description: str = Field(min_length=2,max_length=2000)


@app.post('/api/tickets')
def create_ticket(body: TicketInput, user=Depends(current_user)):
    with LOCK, db() as con:
        if not con.execute('SELECT 1 FROM students WHERE id=?', (body.student_id,)).fetchone():
            raise HTTPException(404, 'Student not found')
        t = con.execute('INSERT INTO tickets(student_id,opened_by,opened_at,title,description) VALUES(?,?,?,?,?)',
                        (body.student_id,user['id'],now(),body.title,body.description))
        con.execute('INSERT INTO ticket_events(ticket_id,actor_id,at,status,note) VALUES(?,?,?,?,?)',
                    (t.lastrowid,user['id'],now(),'Open','Ticket created'))
        audit(con,user['id'],'ticket_created',{'ticket_id':t.lastrowid})
        return {'id':t.lastrowid}


@app.get('/api/tickets')
def tickets(user=Depends(current_user)):
    with db() as con:
        rows=con.execute('''SELECT t.*,s.name AS student_name,s.external_id,u.full_name AS opened_by_name
            FROM tickets t JOIN students s ON s.id=t.student_id JOIN users u ON u.id=t.opened_by
            ORDER BY CASE WHEN t.status='Resolved' THEN 1 ELSE 0 END,t.id DESC''').fetchall()
        return [dict(r) for r in rows]


class TicketEvent(BaseModel):
    status: str
    note: str = Field(min_length=2,max_length=2000)


@app.get('/api/tickets/{tid}/events')
def ticket_history(tid: int,user=Depends(current_user)):
    with db() as con:
        if not con.execute('SELECT 1 FROM tickets WHERE id=?',(tid,)).fetchone():
            raise HTTPException(404,'Ticket not found')
        return [dict(r) for r in con.execute('''SELECT e.*,u.full_name AS actor_name FROM ticket_events e
          JOIN users u ON u.id=e.actor_id WHERE ticket_id=? ORDER BY e.id''',(tid,))]


@app.post('/api/tickets/{tid}/events')
def update_ticket(tid: int,body: TicketEvent,user=Depends(current_user)):
    if body.status not in ('Open','In progress','Resolved'):
        raise HTTPException(400,'Invalid status')
    with LOCK,db() as con:
        con.execute('BEGIN IMMEDIATE')
        if not con.execute('SELECT 1 FROM tickets WHERE id=?',(tid,)).fetchone():
            raise HTTPException(404,'Ticket not found')
        con.execute('INSERT INTO ticket_events(ticket_id,actor_id,at,status,note) VALUES(?,?,?,?,?)',
                    (tid,user['id'],now(),body.status,body.note))
        con.execute('UPDATE tickets SET status=? WHERE id=?',(body.status,tid))
        audit(con,user['id'],'ticket_event',{'ticket_id':tid,'status':body.status})
        con.execute('COMMIT')
        return {'ok':True}


def date_start(period):
    t = datetime.now(timezone.utc)
    if period == 'today': start=t.replace(hour=0,minute=0,second=0,microsecond=0)
    elif period == 'week': start=(t - timedelta(days=t.weekday())).replace(hour=0,minute=0,second=0,microsecond=0)
    elif period == 'month': start=t.replace(day=1,hour=0,minute=0,second=0,microsecond=0)
    elif period == 'all': return '0000'
    else: raise HTTPException(400,'Period must be today, week, month or all')
    return start.isoformat(timespec='seconds')


@app.get('/api/dashboard')
def dashboard(period: str = 'today',user=Depends(current_user)):
    start = date_start(period)
    with db() as con:
        scope = '' if user['role']=='admin' else ' AND caller_id=?'
        args = (start,) if not scope else (start,user['id'])
        main = con.execute(f'''SELECT COUNT(*) AS total_calls,COUNT(DISTINCT student_id) AS unique_students,
            SUM(CASE WHEN outcome='Connected' THEN 1 ELSE 0 END) AS connected_calls
            FROM calls WHERE called_at>=? {scope}''',args).fetchone()
        outstanding = con.execute('SELECT COUNT(*) FROM tickets WHERE status!=?',('Resolved',)).fetchone()[0]
        fargs = (now(),) if user['role']=='admin' else (now(),user['id'])
        due = con.execute('''SELECT COUNT(*) FROM followups WHERE completed_at IS NULL AND due_at<=?'''+
                          ('' if user['role']=='admin' else ' AND assigned_user_id=?'),fargs).fetchone()[0]
        per_caller = [dict(r) for r in con.execute('''SELECT u.full_name AS caller,u.id AS caller_id,
            COUNT(c.id) AS calls,COUNT(DISTINCT c.student_id) AS unique_students
            FROM users u LEFT JOIN calls c ON c.caller_id=u.id AND c.called_at>=?
            WHERE u.role='caller' GROUP BY u.id ORDER BY u.id''',(start,))]
        recent = [dict(r) for r in con.execute('''SELECT c.id,c.called_at,c.outcome,s.name AS student,u.full_name AS caller
            FROM calls c JOIN students s ON s.id=c.student_id JOIN users u ON u.id=c.caller_id
            WHERE c.called_at>=?'''+scope+' ORDER BY c.id DESC LIMIT 12',args)]
        return {'period':period,'total_calls':main['total_calls'],'unique_students':main['unique_students'],
                'connected_calls':main['connected_calls'] or 0,'open_queries':outstanding,
                'overdue_followups':due,'per_caller':per_caller,'recent_calls':recent,
                'total_students':con.execute('SELECT COUNT(*) FROM students').fetchone()[0],
                'refreshed_at':now()}


def make_sheet(wb, title, headings, data):
    ws = wb.create_sheet(title)
    ws.append(headings)
    for r in data:
        ws.append([str(v) if v is not None else '' for v in r])
    ws.freeze_panes='A2'
    ws.auto_filter.ref=ws.dimensions
    for cell in ws[1]:
        cell.fill=PatternFill('solid',fgColor='162A45')
        cell.font=Font(color='FFFFFF',bold=True)
        cell.alignment=Alignment(wrap_text=True)
    for i, col in enumerate(ws.columns,1):
        values=[len(str(c.value or '')) for c in list(col)[:300]]
        ws.column_dimensions[get_column_letter(i)].width=min(max(max(values or [10])+2,13),46)
    return ws


@app.get('/api/reports.xlsx')
def download_report(period: str = 'today',user=Depends(current_user)):
    start=date_start(period)
    with db() as con:
        scope='' if user['role']=='admin' else ' AND c.caller_id=?'
        args=(start,) if not scope else (start,user['id'])
        calls=con.execute('''SELECT c.id,c.called_at,s.external_id,s.name,s.phone,s.course,
          u.full_name,c.outcome,c.student_response,c.query_status,c.query_details,c.notes
          FROM calls c JOIN students s ON s.id=c.student_id JOIN users u ON u.id=c.caller_id
          WHERE c.called_at>=?'''+scope+' ORDER BY c.id',args).fetchall()
        flscope='' if user['role']=='admin' else ' AND f.assigned_user_id=?'
        flargs=() if user['role']=='admin' else (user['id'],)
        followups=con.execute('''SELECT f.id,s.external_id,s.name,u.full_name,f.due_at,f.completed_at,f.completion_note
          FROM followups f JOIN students s ON s.id=f.student_id JOIN users u ON u.id=f.assigned_user_id
          WHERE 1=1'''+flscope+' ORDER BY f.due_at',flargs).fetchall()
        tickets=con.execute('''SELECT t.id,s.external_id,s.name,t.title,t.description,t.status,t.opened_at
          FROM tickets t JOIN students s ON s.id=t.student_id ORDER BY t.id''').fetchall()
        per_caller=con.execute('''SELECT u.full_name,COUNT(c.id),COUNT(DISTINCT c.student_id),
            SUM(CASE WHEN c.outcome='Connected' THEN 1 ELSE 0 END)
          FROM users u LEFT JOIN calls c ON c.caller_id=u.id AND c.called_at>=?
          WHERE u.role='caller' GROUP BY u.id''',(start,)).fetchall()
    wb=Workbook();summary=wb.active;wb.remove(summary)
    unique=len({r[2] for r in calls})
    make_sheet(wb,'Summary',['Metric','Value'],[
        ['Period',period],['Exported (UTC)',now()],['Total call attempts',len(calls)],
        ['Unique students contacted',unique],['Connected attempts',sum(1 for r in calls if r[7]=='Connected')]])
    make_sheet(wb,'Call History',['Call ID','Timestamp UTC','Student ID','Student Name','Phone','Course','Caller',
                                  'Outcome','Student Response','Query Status','Query Details','Notes'],calls)
    make_sheet(wb,'Caller Metrics',['Caller','Total Calls','Unique Students','Connected Attempts'],per_caller)
    make_sheet(wb,'Follow-ups',['ID','Student ID','Student','Assigned Caller','Due UTC','Completed UTC','Completion Note'],followups)
    if user['role']=='admin':
        make_sheet(wb,'Queries',['ID','Student ID','Student','Title','Description','Current Status','Opened UTC'],tickets)
    out=io.BytesIO();wb.save(out);out.seek(0)
    return StreamingResponse(out,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
       headers={'Content-Disposition':f'attachment; filename="crm_{period}_{datetime.now(timezone.utc).strftime("%Y%m%d")}.xlsx"',
                'Cache-Control':'no-store'})


@app.get('/')
def home():
    return FileResponse(BASE / 'static' / 'index.html',headers={'Cache-Control':'no-store'})


@app.get('/static/{file_name}')
def static(file_name: str):
    if file_name not in {'app.js','style.css'}:
        raise HTTPException(404)
    return FileResponse(BASE/'static'/file_name)
