from db import get_conn

from db import get_conn


def create_linen_cycle(cycle_name, created_by):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO linen_cycles (
            cycle_name,
            status,
            created_by
        )
        VALUES (
            %s,
            'DRAFT',
            %s
        )
    """, (cycle_name, created_by))

    conn.commit()
    conn.close()


def get_linen_cycles():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            id,
            cycle_name,
            status,
            created_by,
            created_at,
            started_at,
            completed_at
        FROM linen_cycles
        ORDER BY created_at DESC
    """)

    rows = cur.fetchall()

    conn.close()

    return rows

def get_linen_cycle(cycle_id):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            id,
            cycle_name,
            status,
            created_by,
            created_at,
            started_at,
            completed_at
        FROM linen_cycles
        WHERE id = %s
    """, (cycle_id,))

    row = cur.fetchone()

    conn.close()

    return row

def get_cycle_reps(cycle_id):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM linen_cycle_reps
        WHERE cycle_id = %s
        ORDER BY CAST(REPLACE(rep_username, 'LINREP', '') AS INTEGER)
    """, (cycle_id,))

    rows = cur.fetchall()

    conn.close()

    return rows

def save_cycle_reps(cycle_id, rep_rows):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        DELETE FROM linen_cycle_reps
        WHERE cycle_id = %s
    """, (cycle_id,))

    for row in rep_rows:
        cur.execute("""
            INSERT INTO linen_cycle_reps (
                cycle_id,
                rep_username,
                display_name
            )
            VALUES (%s, %s, %s)
        """, (
            cycle_id,
            row["rep_username"],
            row["display_name"],
        ))

    conn.commit()
    conn.close()

def get_cycle_assignments(cycle_id):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM linen_location_assignments
        WHERE cycle_id = %s
        ORDER BY assigned_to, location_id
    """, (cycle_id,))

    rows = cur.fetchall()

    conn.close()

    return [dict(r) for r in rows]

def save_cycle_assignments(
    cycle_id,
    assignments
):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        DELETE FROM linen_location_assignments
        WHERE cycle_id = %s
    """, (cycle_id,))

    for row in assignments:

        cur.execute("""
            INSERT INTO linen_location_assignments (
                cycle_id,
                location_id,
                assigned_to,
                assigned_type
            )
            VALUES (%s, %s, %s, %s)
        """, (
            cycle_id,
            row["location_id"],
            row["assigned_to"],
            row["assigned_type"]
        ))

    conn.commit()
    conn.close()

def start_linen_cycle(cycle_id):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        UPDATE linen_cycles
        SET status = 'ACTIVE',
            started_at = CURRENT_TIMESTAMP
        WHERE id = %s
          AND status = 'DRAFT'
    """, (cycle_id,))

    conn.commit()
    conn.close()

def get_active_linen_cycle():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM linen_cycles
        WHERE status = 'ACTIVE'
        ORDER BY id DESC
        LIMIT 1
    """)

    row = cur.fetchone()

    conn.close()

    return row

def get_assignments_for_user(
    cycle_id,
    username
):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM linen_location_assignments
        WHERE cycle_id = %s
        AND assigned_to = %s
        ORDER BY location_id
    """, (
        cycle_id,
        username
    ))

    rows = cur.fetchall()

    conn.close()

    return [dict(r) for r in rows]

def get_submission(cycle_id, location_id):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM linen_submissions
        WHERE cycle_id = %s
          AND location_id = %s
        LIMIT 1
    """, (
        cycle_id,
        location_id
    ))

    row = cur.fetchone()

    conn.close()

    return dict(row) if row else None

def save_submission_draft(
    cycle_id,
    location_id,
    submitted_by,
    lines,
    bundle_lines=None
):
    if bundle_lines is None:
        bundle_lines = []
    
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM linen_submissions
        WHERE cycle_id = %s
          AND location_id = %s
        LIMIT 1
    """, (
        cycle_id,
        location_id
    ))

    submission = cur.fetchone()

    if submission:
        submission_id = submission["id"]

        cur.execute("""
            DELETE FROM linen_submission_lines
            WHERE submission_id = %s
        """, (submission_id,))

        cur.execute("""
            DELETE FROM linen_submission_bundle_lines
            WHERE submission_id = %s
        """, (submission_id,))

        cur.execute("""
            UPDATE linen_submissions
            SET
                submitted_by = %s,
                status = CASE
                    WHEN status = 'SUBMITTED' THEN 'SUBMITTED'
                    ELSE 'PENDING'
                END,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = %s
        """, (
            submitted_by,
            submission_id
        ))

    else:
        cur.execute("""
            INSERT INTO linen_submissions (
                cycle_id,
                location_id,
                submitted_by,
                status
            )
            VALUES (%s, %s, %s, 'PENDING')
            RETURNING id
        """, (
            cycle_id,
            location_id,
            submitted_by
        ))

        submission_id = cur.fetchone()["id"]

    for line in lines:
        cur.execute("""
            INSERT INTO linen_submission_lines (
                submission_id,
                item_no,
                quantity
            )
            VALUES (%s, %s, %s)
        """, (
            submission_id,
            line["item_no"],
            int(line["quantity"] or 0)
        ))

    for bundle_line in bundle_lines:
        qty = int(
            bundle_line.get("quantity") or 0
        )

        if qty < 0:
            raise ValueError(
                "Bundle quantity cannot be negative."
            )

        cur.execute("""
            INSERT INTO linen_submission_bundle_lines (
                submission_id,
                bundle_id,
                quantity
            )
            VALUES (%s, %s, %s)
        """, (
            submission_id,
            str(bundle_line["bundle_id"]),
            qty,
        ))

    conn.commit()
    conn.close()

def get_submission_lines(submission_id):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM linen_submission_lines
        WHERE submission_id = %s
    """, (submission_id,))

    rows = cur.fetchall()

    conn.close()

    return [dict(r) for r in rows]

def get_submission_bundle_lines(submission_id):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            bundle_id,
            quantity
        FROM linen_submission_bundle_lines
        WHERE submission_id = %s
    """, (submission_id,))

    rows = cur.fetchall()

    conn.close()

    return [dict(r) for r in rows]

def submit_submission(
    cycle_id,
    location_id,
    submitted_by
):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        UPDATE linen_submissions
        SET
            status = 'SUBMITTED',
            submitted_by = %s,
            submitted_at = CURRENT_TIMESTAMP,
            updated_at = CURRENT_TIMESTAMP
        WHERE cycle_id = %s
          AND location_id = %s
    """, (
        submitted_by,
        cycle_id,
        location_id
    ))

    conn.commit()
    conn.close()

def get_submission_status_map(cycle_id):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT location_id, status
        FROM linen_submissions
        WHERE cycle_id = %s
    """, (cycle_id,))

    rows = cur.fetchall()
    conn.close()

    return {
        row["location_id"]: row["status"]
        for row in rows
    }

def complete_linen_cycle(cycle_id):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        UPDATE linen_cycles
        SET
            status = 'COMPLETED',
            completed_at = CURRENT_TIMESTAMP
        WHERE id = %s
          AND status = 'ACTIVE'
    """, (cycle_id,))

    conn.commit()
    conn.close()

def get_submitted_location_count(cycle_id):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT COUNT(DISTINCT location_id) AS cnt
        FROM linen_submissions
        WHERE cycle_id = %s
          AND status = 'SUBMITTED'
    """, (cycle_id,))

    row = cur.fetchone()
    conn.close()

    return int(row["cnt"] or 0)

def get_cycle_submission_lines(cycle_id):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            s.location_id,
            s.submitted_by,
            s.status,
            s.submitted_at,
            l.item_no,
            l.quantity
        FROM linen_submissions s
        JOIN linen_submission_lines l
            ON s.id = l.submission_id
        WHERE s.cycle_id = %s
          AND s.status = 'SUBMITTED'
        ORDER BY s.location_id, l.item_no
    """, (cycle_id,))

    rows = cur.fetchall()
    conn.close()

    return [dict(r) for r in rows]

def force_complete_linen_cycle(cycle_id):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        UPDATE linen_cycles
        SET status = 'COMPLETED',
            completed_at = CURRENT_TIMESTAMP
        WHERE id = %s
    """, (cycle_id,))

    conn.commit()
    conn.close()

def get_linen_rep_names(cycle_id):
    rows = get_cycle_reps(cycle_id)

    return {
        row["rep_username"]: row["display_name"]
        for row in rows
    }