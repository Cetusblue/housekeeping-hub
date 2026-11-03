"""Independent Linen Call Log storage. No inventory or top-up side effects."""
from datetime import datetime
from zoneinfo import ZoneInfo
from db import get_conn, DB_TYPE

SGT = ZoneInfo('Asia/Singapore')

def _now():
    return datetime.now(SGT).strftime('%Y-%m-%d %H:%M:%S')

def _dict(row):
    return dict(row)

def init_linen_call_log():
    conn = get_conn()
    try:
        cur = conn.cursor()
        if DB_TYPE == 'postgres':
            cur.execute('''CREATE TABLE IF NOT EXISTS linen_call_logs (
                id BIGSERIAL PRIMARY KEY, created_at TEXT NOT NULL,
                tower TEXT NOT NULL, level TEXT NOT NULL, zone TEXT NOT NULL,
                created_by TEXT NOT NULL, deleted_at TEXT, deleted_by TEXT)''')
            cur.execute('''CREATE TABLE IF NOT EXISTS linen_call_log_lines (
                id BIGSERIAL PRIMARY KEY, call_id BIGINT NOT NULL REFERENCES linen_call_logs(id),
                item_key TEXT NOT NULL, item_name TEXT NOT NULL,
                request_qty INTEGER NOT NULL CHECK(request_qty >= 0),
                delivered_qty INTEGER NOT NULL CHECK(delivered_qty >= 0))''')
        else:
            cur.execute('''CREATE TABLE IF NOT EXISTS linen_call_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT, created_at TEXT NOT NULL,
                tower TEXT NOT NULL, level TEXT NOT NULL, zone TEXT NOT NULL,
                created_by TEXT NOT NULL, deleted_at TEXT, deleted_by TEXT)''')
            cur.execute('''CREATE TABLE IF NOT EXISTS linen_call_log_lines (
                id INTEGER PRIMARY KEY AUTOINCREMENT, call_id INTEGER NOT NULL,
                item_key TEXT NOT NULL, item_name TEXT NOT NULL,
                request_qty INTEGER NOT NULL CHECK(request_qty >= 0),
                delivered_qty INTEGER NOT NULL CHECK(delivered_qty >= 0))''')
        conn.commit()
    finally:
        conn.close()

def create_call(tower, level, zone, username, lines):
    lines = [x for x in lines if x['request_qty'] or x['delivered_qty']]
    if not lines:
        raise ValueError('Enter at least one request or delivery quantity.')
    conn = get_conn()
    try:
        cur = conn.cursor()
        values = (_now(), tower, level, zone, username)
        if DB_TYPE == 'postgres':
            cur.execute('INSERT INTO linen_call_logs (created_at,tower,level,zone,created_by) VALUES (%s,%s,%s,%s,%s) RETURNING id',values)
            call_id = cur.fetchone()['id']
            placeholder = '%s'
        else:
            cur.execute('INSERT INTO linen_call_logs (created_at,tower,level,zone,created_by) VALUES (?,?,?,?,?)',values)
            call_id = cur.lastrowid
            placeholder = '?'
        for line in lines:
            cur.execute('INSERT INTO linen_call_log_lines (call_id,item_key,item_name,request_qty,delivered_qty) VALUES ('+','.join([placeholder]*5)+')',
                        (call_id,line['item_key'],line['item_name'],line['request_qty'],line['delivered_qty']))
        conn.commit()
        return int(call_id)
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def list_calls(limit=None, month=None, include_deleted=False):
    conn = get_conn()
    try:
        cur = conn.cursor()
        sql = 'SELECT * FROM linen_call_logs WHERE 1=1'
        params=[]
        if not include_deleted: sql += ' AND deleted_at IS NULL'
        if month:
            sql += ' AND substr(created_at,1,7) = '+('%s' if DB_TYPE=='postgres' else '?')
            params.append(month)
        sql += ' ORDER BY id DESC'
        if limit:
            sql += ' LIMIT '+str(int(limit))
        cur.execute(sql,tuple(params))
        return [_dict(x) for x in cur.fetchall()]
    finally: conn.close()

def get_call(call_id):
    conn=get_conn()
    try:
        cur=conn.cursor()
        p='%s' if DB_TYPE=='postgres' else '?'
        cur.execute('SELECT * FROM linen_call_logs WHERE id='+p,(call_id,))
        row=cur.fetchone()
        if not row: return None,None
        cur.execute('SELECT * FROM linen_call_log_lines WHERE call_id='+p+' ORDER BY id',(call_id,))
        return _dict(row),[_dict(x) for x in cur.fetchall()]
    finally: conn.close()

def delete_call(call_id, username):
    conn=get_conn()
    try:
        cur=conn.cursor()
        p='%s' if DB_TYPE=='postgres' else '?'
        cur.execute('UPDATE linen_call_logs SET deleted_at='+p+',deleted_by='+p+' WHERE id='+p+' AND deleted_at IS NULL',(_now(),username,call_id))
        count=cur.rowcount
        conn.commit()
        return count==1
    except Exception:
        conn.rollback();raise
    finally:conn.close()
