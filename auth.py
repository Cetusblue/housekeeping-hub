import hashlib
import secrets
from datetime import datetime, timedelta

from db import get_conn, ph, now_iso


def hash_password(raw_password: str) -> str:
    return hashlib.sha256(raw_password.encode("utf-8")).hexdigest()


def verify_password(raw_password: str, stored_hash: str) -> bool:
    return hash_password(raw_password) == stored_hash


def _user_dict(row):
    return {
        "user_id": row["user_id"],
        "username": row["username"],
        "display_name": row["display_name"],
        "role": row["role"],
        "team_code": row["team_code"],
    }


def authenticate(username: str, raw_password: str):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute(f"""
        SELECT user_id, username, display_name, password_hash, role, team_code, active
        FROM users
        WHERE username = {ph()} AND active = 'Y'
    """, (username,))

    row = cur.fetchone()
    conn.close()

    if not row or row["active"] != "Y":
        return None

    if not verify_password(raw_password, row["password_hash"]):
        return None

    return _user_dict(row)


def create_persistent_session(user_id: int, days_valid: int = 30) -> str:
    """Create a random browser token and store only its SHA-256 hash."""
    raw_token = secrets.token_urlsafe(32)
    token_hash = hashlib.sha256(raw_token.encode("utf-8")).hexdigest()
    expires_at = (datetime.now() + timedelta(days=days_valid)).strftime("%Y-%m-%d %H:%M:%S")

    conn = get_conn()
    cur = conn.cursor()

    # Keep the table tidy whenever a new session is created.
    cur.execute(f"DELETE FROM auth_sessions WHERE expires_at < {ph()}", (now_iso(),))
    cur.execute(f"""
        INSERT INTO auth_sessions (user_id, token_hash, created_at, expires_at)
        VALUES ({ph()}, {ph()}, {ph()}, {ph()})
    """, (int(user_id), token_hash, now_iso(), expires_at))

    conn.commit()
    conn.close()
    return raw_token


def authenticate_persistent_session(raw_token: str):
    if not raw_token:
        return None

    token_hash = hashlib.sha256(raw_token.encode("utf-8")).hexdigest()

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(f"""
        SELECT u.user_id, u.username, u.display_name, u.role, u.team_code, u.active
        FROM auth_sessions s
        JOIN users u ON u.user_id = s.user_id
        WHERE s.token_hash = {ph()}
          AND s.expires_at >= {ph()}
          AND u.active = 'Y'
        LIMIT 1
    """, (token_hash, now_iso()))

    row = cur.fetchone()
    conn.close()
    return _user_dict(row) if row else None


def revoke_persistent_session(raw_token: str):
    if not raw_token:
        return

    token_hash = hashlib.sha256(raw_token.encode("utf-8")).hexdigest()
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(f"DELETE FROM auth_sessions WHERE token_hash = {ph()}", (token_hash,))
    conn.commit()
    conn.close()
