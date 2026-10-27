"""Telegram notification delivery for OPS HUB toilet requests.

Token is read from environment variables or Streamlit secrets. Recipient selection uses toilet_routing.py; Demo-zone toilets route to
their assigned User Master group regardless of shift.
"""
import json
import os
from urllib import request, parse, error
import streamlit as st
from master_loader import _open_workbook
from toilet_routing import select_recipients


def setting(key, default=''):
    """Prefer deployment environment variables, fall back to Streamlit secrets."""
    value = os.environ.get(key)
    if value is not None:
        return str(value).strip()
    try:
        return str(st.secrets.get(key, default) or default).strip()
    except Exception:
        return default


def _eligible_users():
    wb = _open_workbook()
    try:
        rows = wb['User Master'].iter_rows(values_only=True)
        headers = [str(v or '').strip().lower() for v in next(rows)]
        if 'telegram_chat_id' not in headers:
            raise ValueError('Add telegram_chat_id column to User Master')
        users = []
        for row in rows:
            d = dict(zip(headers, row))
            if str(d.get('active') or '').strip().upper() not in ('Y', 'YES', 'TRUE', '1'):
                continue
            chat_id = str(d.get('telegram_chat_id') or '').strip()
            if chat_id.endswith('.0'):
                chat_id = chat_id[:-2]
            if not chat_id:
                continue
            users.append({'username': str(d.get('username') or '').strip(),
                          'group': str(d.get('group') or '').strip().upper(),
                          'chat_id': chat_id})
        return users
    finally:
        wb.close()


def send(toilet, request_type, details, when):
    """Backward-compatible single-issue notification."""
    return send_many(toilet, [(request_type, details)], when)


def send_many(toilet, issues, when):
    if setting('TOILET_TELEGRAM_MODE', 'off').lower() != 'on':
        return {'status': 'DISABLED'}
    token = setting('TELEGRAM_BOT_TOKEN')
    if not token:
        return {'status': 'CONFIG_MISSING'}
    try:
        people = _eligible_users()
        selected = select_recipients(toilet, when, people)
    except Exception as exc:
        return {'status': 'FAILED', 'reason': type(exc).__name__}
    if not selected:
        return {'status': 'NO_RECIPIENT'}
    delivered = 0
    failed = 0
    missing = 0
    for user in selected:
        if not user['chat_id']:
            missing += 1
            continue
        issue_lines = []
        for item, detail in issues:
            issue_lines.append(f'• {item}' + (f' — {detail}' if detail else ''))
        msg = ('OPS HUB | Toilet Service Request\n'
               f"Location: {toilet.get('toilet_name', '')}\n"
               f"Tower: {toilet.get('tower', '')} | Level: {toilet.get('level', '')}\n"
               f'{len(issues)} issue(s) reported:\n' + '\n'.join(issue_lines) + '\n'
               + f'Time: {when:%d %b %Y, %I:%M %p} (SGT)')
        payload = parse.urlencode({'chat_id': user['chat_id'], 'text': msg}).encode()
        req = request.Request(f'https://api.telegram.org/bot{token}/sendMessage',
                              data=payload, method='POST')
        try:
            with request.urlopen(req, timeout=8) as response:
                result = json.loads(response.read())
            delivered += bool(result.get('ok'))
            failed += not bool(result.get('ok'))
        except (error.URLError, TimeoutError, ValueError):
            failed += 1
    if delivered and not failed and not missing:
        status = 'ACCEPTED'
    elif delivered:
        status = 'PARTIAL'
    elif missing and not failed:
        status = 'NO_CHAT_ID'
    else:
        status = 'FAILED'
    return {'status': status, 'accepted': delivered, 'failed': failed, 'missing_chat_id': missing}
