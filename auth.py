import hashlib
from db import get_conn, ph


def hash_password(raw_password: str) -> str:
    return hashlib.sha256(raw_password.encode("utf-8")).hexdigest()


def verify_password(raw_password: str, stored_hash: str) -> bool:
    return hash_password(raw_password) == stored_hash


def authenticate(username: str, raw_password: str):
    conn = get_conn()
    cur = conn.cursor()

    query = f"""
        SELECT user_id, username, display_name, password_hash, role, team_code, active
        FROM users
        WHERE username = {ph()} AND active = 'Y'
    """
    cur.execute(query, (username,))

    row = cur.fetchone()
    conn.close()

    if not row:
        return None

    if row["active"] != "Y":
        return None

    if not verify_password(raw_password, row["password_hash"]):
        return None

    return {
        "user_id": row["user_id"],
        "username": row["username"],
        "display_name": row["display_name"],
        "role": row["role"],
        "team_code": row["team_code"],
    }