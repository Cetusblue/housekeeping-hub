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
        ORDER BY rep_username
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