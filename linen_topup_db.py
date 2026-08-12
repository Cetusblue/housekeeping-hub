from db import get_conn
import hashlib
import secrets

from db import get_conn


def create_manual_topup(location_id, created_by, lines):
    """
    Creates one Manual Top Up transaction.

    lines:
    [
        {
            "item_no": "...",
            "quantity": 10
        },
        ...
    ]
    """

    positive_lines = [
        line
        for line in lines
        if int(line.get("quantity") or 0) > 0
    ]

    if not positive_lines:
        raise ValueError("No top-up quantities entered.")

    conn = get_conn()
    cur = conn.cursor()

    try:
        cur.execute("""
            INSERT INTO linen_topups (
                location_id,
                created_by
            )
            VALUES (%s, %s)
            RETURNING topup_id
        """, (
            location_id,
            created_by
        ))

        topup_id = cur.fetchone()["topup_id"]

        for line in positive_lines:
            cur.execute("""
                INSERT INTO linen_topup_lines (
                    topup_id,
                    item_no,
                    quantity
                )
                VALUES (%s, %s, %s)
            """, (
                topup_id,
                line["item_no"],
                int(line["quantity"])
            ))

        conn.commit()
        return int(topup_id)

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()


def get_recent_manual_topups(limit=10):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            t.topup_id,
            t.location_id,
            t.created_by,
            t.created_at,
            SUM(l.quantity) AS total_quantity
        FROM linen_topups t
        JOIN linen_topup_lines l
            ON l.topup_id = t.topup_id
        GROUP BY
            t.topup_id,
            t.location_id,
            t.created_by,
            t.created_at
        ORDER BY t.created_at DESC
        LIMIT %s
    """, (int(limit),))

    rows = [dict(r) for r in cur.fetchall()]
    conn.close()

    return rows

def _hash_qr_token(raw_token):
    return hashlib.sha256(
        raw_token.encode("utf-8")
    ).hexdigest()


def create_qr_token_for_location(location_id):
    """
    Creates/replaces the QR token for one Manual Top Up location.

    Returns the raw token ONCE so it can be encoded into the QR.
    Only the token hash is stored in the database.
    """

    raw_token = secrets.token_urlsafe(24)
    token_hash = _hash_qr_token(raw_token)

    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO linen_topup_qr_tokens (
            location_id,
            token_hash,
            active
        )
        VALUES (%s, %s, TRUE)
        ON CONFLICT (location_id)
        DO UPDATE SET
            token_hash = EXCLUDED.token_hash,
            active = TRUE,
            created_at = CURRENT_TIMESTAMP
    """, (
        location_id,
        token_hash,
    ))

    conn.commit()
    conn.close()

    return raw_token


def resolve_qr_token(raw_token):
    """
    Returns location_id if token is valid and active.
    Returns None otherwise.
    """

    if not raw_token:
        return None

    token_hash = _hash_qr_token(raw_token)

    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT location_id
        FROM linen_topup_qr_tokens
        WHERE token_hash = %s
          AND active = TRUE
        LIMIT 1
    """, (token_hash,))

    row = cur.fetchone()
    conn.close()

    if not row:
        return None

    return row["location_id"]