import io
import os
import sqlite3
from pathlib import Path

import httpx
import pytest

import app as crm


@pytest.mark.asyncio
async def test_end_to_end(tmp_path, monkeypatch):
    monkeypatch.setattr(crm, 'DB_PATH', tmp_path / 'test.sqlite3')
    monkeypatch.setenv('CRM_ADMIN_PASSWORD', 'AdminStrongPassword!123')
    monkeypatch.setenv('CRM_CALLER1_PASSWORD', 'CallerOnePassword!123')
    monkeypatch.setenv('CRM_CALLER2_PASSWORD', 'CallerTwoPassword!123')
    crm.initialize()
    async with httpx.AsyncClient(transport=httpx.ASGITransport(app=crm.app), base_url='http://testserver') as admin, \
               httpx.AsyncClient(transport=httpx.ASGITransport(app=crm.app), base_url='http://testserver') as caller1, \
               httpx.AsyncClient(transport=httpx.ASGITransport(app=crm.app), base_url='http://testserver') as caller2:
        async def sign_in(client, user, pwd):
            r = await client.post('/api/login', json={'username':user,'password':pwd})
            assert r.status_code == 200, r.text
            client.headers['X-CSRF-Token'] = r.json()['csrf']
        await sign_in(admin, 'admin', 'AdminStrongPassword!123')
        await sign_in(caller1, 'caller1', 'CallerOnePassword!123')
        await sign_in(caller2, 'caller2', 'CallerTwoPassword!123')

        csv_text='student_id,name,phone,course\n001,Ananya,9000000001,Design\n002,Rohan,9000000002,Technology\n'
        r=await admin.post('/api/students/import',json={'csv_text':csv_text})
        assert r.status_code == 200 and r.json()['added']==2
        assert (await admin.post('/api/students/import',json={'csv_text':csv_text})).json()['skipped_existing']==2
        assert (await caller1.post('/api/students/import',json={'csv_text':csv_text})).status_code==403
        students=(await caller1.get('/api/students')).json()['items']
        student_id=next(x['id'] for x in students if x['external_id']=='001')
        assert len((await caller2.get('/api/students')).json()['items'])==2

        for client, outcome in ((caller1, 'Connected'), (caller2, 'No answer')):
            r=await client.post('/api/calls',json={'student_id':student_id,'outcome':outcome,
                'student_response':'Needs information','query_status':'No query','notes':'Separate historical attempt',
                'followup_at':'2025-12-01T10:00:00Z'})
            assert r.status_code==200, r.text
        d=(await admin.get('/api/dashboard?period=all')).json()
        assert d['total_calls']==2 and d['unique_students']==1 and d['connected_calls']==1
        assert [x['calls'] for x in d['per_caller']]==[1,1]
        assert len((await admin.get('/api/students/'+str(student_id))).json()['calls'])==2
        with sqlite3.connect(crm.DB_PATH) as con:
            for statement in ('DELETE FROM calls WHERE id=1', "UPDATE calls SET notes='overwritten' WHERE id=1"):
                with pytest.raises(sqlite3.IntegrityError):
                    con.execute(statement)
        followups=(await admin.get('/api/followups')).json()
        assert len(followups)==2
        assert (await caller2.post(f'/api/followups/{followups[0]["id"]}/complete',json={'note':'done'})).status_code==403
        assert (await caller1.post(f'/api/followups/{followups[0]["id"]}/complete',json={'note':'done'})).status_code==200
        assert len((await admin.get('/api/students/'+str(student_id))).json()['calls'])==2

        t=await caller1.post('/api/tickets',json={'student_id':student_id,'title':'Certificate query','description':'Needs certificate'})
        assert t.status_code==200
        tid=t.json()['id']
        r=await caller2.post(f'/api/tickets/{tid}/events',json={'status':'Resolved','note':'Certificate issued'})
        assert r.status_code==200
        assert len((await admin.get(f'/api/tickets/{tid}/events')).json())==2
        with sqlite3.connect(crm.DB_PATH) as con:
            with pytest.raises(sqlite3.IntegrityError):
                con.execute('DELETE FROM ticket_events WHERE id=1')

        x=await admin.get('/api/reports.xlsx?period=all')
        assert x.status_code==200 and x.content[:2]==b'PK'
        from openpyxl import load_workbook
        wb=load_workbook(io.BytesIO(x.content), read_only=True)
        assert 'Call History' in wb.sheetnames and 'Caller Metrics' in wb.sheetnames and 'Queries' in wb.sheetnames
        assert wb['Call History'].max_row==3
        own=await caller1.get('/api/reports.xlsx?period=all')
        own_wb=load_workbook(io.BytesIO(own.content),read_only=True)
        assert own_wb['Call History'].max_row==2
        assert 'Queries' not in own_wb.sheetnames
        assert (await caller1.get('/api/dashboard?period=all')).json()['total_calls']==1
        assert (await admin.get('/')).status_code==200
