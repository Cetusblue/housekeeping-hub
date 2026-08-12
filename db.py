import os
import sqlite3
from datetime import datetime
DB_TYPE = os.getenv("DB_TYPE", "sqlite").lower()

def now_iso():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "housekeeping_hub.db")


def get_conn():
    if DB_TYPE == "postgres":
        import psycopg2
        from psycopg2.extras import RealDictCursor

        return psycopg2.connect(
            os.getenv("DATABASE_URL"),
            cursor_factory=RealDictCursor
        )

    os.makedirs(DATA_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    if DB_TYPE == "postgres":
        init_db_postgres()
    else:
        init_db_sqlite()

def init_db_sqlite():
    conn = get_conn()
    cur = conn.cursor()

    # ---------------------------
    # Users
    # ---------------------------
    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        user_id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE,
        display_name TEXT NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL CHECK(role IN ('TEAM', 'STORE', 'LINREP', 'LINTEAM', 'LINSUP', 'ADMIN', 'BOSS')),
        team_code TEXT NOT NULL,
        active TEXT NOT NULL DEFAULT 'Y' CHECK(active IN ('Y', 'N'))
    )
    """)

    # ---------------------------
    # Persistent login sessions
    # ---------------------------
    cur.execute("""
    CREATE TABLE IF NOT EXISTS auth_sessions (
        session_id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        token_hash TEXT NOT NULL UNIQUE,
        created_at TEXT NOT NULL,
        expires_at TEXT NOT NULL,
        FOREIGN KEY(user_id) REFERENCES users(user_id)
    )
    """)

    cur.execute("""
    CREATE INDEX IF NOT EXISTS idx_auth_sessions_expiry
    ON auth_sessions(expires_at)
    """)

    # ---------------------------
    # Orders (header)
    # ---------------------------
    cur.execute("""
    CREATE TABLE IF NOT EXISTS orders (
        order_id INTEGER PRIMARY KEY AUTOINCREMENT,
        team_code TEXT NOT NULL,
        template_day TEXT NOT NULL CHECK(template_day IN ('TUE', 'FRI', 'OT', 'ANNEX_TUE', 'ANNEX_FRI')),
        run_date TEXT NOT NULL,
        status TEXT NOT NULL CHECK(status IN ('PENDING', 'PARTIALLY_ISSUED', 'ISSUED', 'CANCELLED', 'CLOSED')),
        created_by TEXT NOT NULL,
        created_at TEXT NOT NULL,
        issued_by TEXT,
        issued_at TEXT,
        cancelled_at TEXT,
        cancelled_by TEXT,
        cancel_reason TEXT,
        closed_reason TEXT,
        closed_by TEXT,
        closed_at TEXT
    )
    """)

    cur.execute("""
    CREATE UNIQUE INDEX IF NOT EXISTS idx_orders_unique_group_day_date
    ON orders(team_code, template_day, run_date)
    """)

    # ---------------------------
    # Order lines
    # ---------------------------
    cur.execute("""
    CREATE TABLE IF NOT EXISTS order_lines (
        line_id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER NOT NULL,
        item_no INTEGER NOT NULL,
        item_name TEXT NOT NULL,
        qty_requested INTEGER NOT NULL DEFAULT 0,
        qty_issued INTEGER NOT NULL DEFAULT 0,
        UNIQUE(order_id, item_no),
        FOREIGN KEY(order_id) REFERENCES orders(order_id)
    )
    """)

    # ---------------------------
    # Stock movements
    # ---------------------------
    cur.execute("""
    CREATE TABLE IF NOT EXISTS stock_movements (
        movement_id INTEGER PRIMARY KEY AUTOINCREMENT,
        item_name TEXT NOT NULL,
        movement_type TEXT NOT NULL CHECK(movement_type IN ('IN', 'OUT')),
        qty INTEGER NOT NULL CHECK(qty > 0),
        issued_to TEXT,
        source_type TEXT NOT NULL CHECK(source_type IN ('STOCK_IN', 'ADHOC', 'ORDER')),
        source_id INTEGER,
        created_by TEXT NOT NULL,
        created_at TEXT NOT NULL
    )
    """)

    # ---------------------------
    # Glo Gel audits (header)
    # ---------------------------
    cur.execute("""
    CREATE TABLE IF NOT EXISTS audits (
        audit_id INTEGER PRIMARY KEY AUTOINCREMENT,
        audit_date TEXT NOT NULL,
        auditor_name TEXT NOT NULL,
        staff_name TEXT NOT NULL,
        location_name TEXT NOT NULL,
        tower TEXT NOT NULL,
        zone TEXT,
        template_group TEXT NOT NULL CHECK(template_group IN ('AC', 'B')),
        remarks TEXT,
        status TEXT NOT NULL DEFAULT 'DRAFT' CHECK(status IN ('DRAFT', 'COMPLETED')),
        created_by TEXT NOT NULL,
        created_at TEXT NOT NULL
    )
    """)

    # ---------------------------
    # Glo Gel audit results (detail)
    # ---------------------------
    cur.execute("""
    CREATE TABLE IF NOT EXISTS audit_results (
        result_id INTEGER PRIMARY KEY AUTOINCREMENT,
        audit_id INTEGER NOT NULL,
        surface_id TEXT,
        surface_name TEXT NOT NULL,
        result TEXT NOT NULL CHECK(result IN ('C', 'NC', 'NA')),
        is_additional TEXT NOT NULL DEFAULT 'N' CHECK(is_additional IN ('Y', 'N')),
        area_group TEXT,
        display_order INTEGER,
        FOREIGN KEY(audit_id) REFERENCES audits(audit_id)
    )
    """)

def init_db_postgres():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        user_id SERIAL PRIMARY KEY,
        username TEXT NOT NULL UNIQUE,
        display_name TEXT NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL CHECK(role IN ('TEAM', 'STORE', 'LINREP', 'LINTEAM', 'LINSUP', 'ADMIN', 'BOSS')),
        team_code TEXT NOT NULL,
        active TEXT NOT NULL DEFAULT 'Y' CHECK(active IN ('Y', 'N'))
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS auth_sessions (
        session_id SERIAL PRIMARY KEY,
        user_id INTEGER NOT NULL REFERENCES users(user_id),
        token_hash TEXT NOT NULL UNIQUE,
        created_at TEXT NOT NULL,
        expires_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE INDEX IF NOT EXISTS idx_auth_sessions_expiry
    ON auth_sessions(expires_at)
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS orders (
        order_id SERIAL PRIMARY KEY,
        team_code TEXT NOT NULL,
        template_day TEXT NOT NULL CHECK(template_day IN ('TUE', 'FRI', 'OT', 'ANNEX_TUE', 'ANNEX_FRI')),
        run_date TEXT NOT NULL,
        status TEXT NOT NULL CHECK(status IN ('PENDING', 'PARTIALLY_ISSUED', 'ISSUED', 'CANCELLED', 'CLOSED')),
        created_by TEXT NOT NULL,
        created_at TEXT NOT NULL,
        issued_by TEXT,
        issued_at TEXT,
        cancelled_at TEXT,
        cancelled_by TEXT,
        cancel_reason TEXT,
        closed_reason TEXT,
        closed_by TEXT,
        closed_at TEXT
    )
    """)

    # Keep existing deployments compatible with newer order end states.
    cur.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS cancelled_at TEXT")
    cur.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS cancelled_by TEXT")
    cur.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS cancel_reason TEXT")
    cur.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS closed_reason TEXT")
    cur.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS closed_by TEXT")
    cur.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS closed_at TEXT")

    # Replace the original status CHECK so CLOSED and CANCELLED are accepted.
    cur.execute("""
        DO $$
        DECLARE
            constraint_name TEXT;
        BEGIN
            SELECT conname INTO constraint_name
            FROM pg_constraint
            WHERE conrelid = 'orders'::regclass
              AND contype = 'c'
              AND pg_get_constraintdef(oid) ILIKE '%status%'
              AND pg_get_constraintdef(oid) ILIKE '%PARTIALLY_ISSUED%'
            LIMIT 1;

            IF constraint_name IS NOT NULL THEN
                EXECUTE format('ALTER TABLE orders DROP CONSTRAINT %I', constraint_name);
            END IF;
        END $$;
    """)
    cur.execute("""
        ALTER TABLE orders
        ADD CONSTRAINT orders_status_check
        CHECK (status IN ('PENDING', 'PARTIALLY_ISSUED', 'ISSUED', 'CANCELLED', 'CLOSED'))
    """)

    cur.execute("""
    CREATE UNIQUE INDEX IF NOT EXISTS idx_orders_unique_group_day_date
    ON orders(team_code, template_day, run_date)
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS order_lines (
        line_id SERIAL PRIMARY KEY,
        order_id INTEGER NOT NULL,
        item_no INTEGER NOT NULL,
        item_name TEXT NOT NULL,
        qty_requested INTEGER NOT NULL DEFAULT 0,
        qty_issued INTEGER NOT NULL DEFAULT 0,
        UNIQUE(order_id, item_no),
        FOREIGN KEY(order_id) REFERENCES orders(order_id)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS stock_movements (
        movement_id SERIAL PRIMARY KEY,
        item_name TEXT NOT NULL,
        movement_type TEXT NOT NULL CHECK(movement_type IN ('IN', 'OUT')),
        qty INTEGER NOT NULL CHECK(qty > 0),
        issued_to TEXT,
        source_type TEXT NOT NULL CHECK(source_type IN ('STOCK_IN', 'ADHOC', 'ORDER')),
        source_id INTEGER,
        created_by TEXT NOT NULL,
        created_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS audits (
        audit_id SERIAL PRIMARY KEY,
        audit_date TEXT NOT NULL,
        auditor_name TEXT NOT NULL,
        staff_name TEXT NOT NULL,
        location_name TEXT NOT NULL,
        tower TEXT NOT NULL,
        zone TEXT,
        template_group TEXT NOT NULL CHECK(template_group IN ('AC', 'B')),
        remarks TEXT,
        status TEXT NOT NULL DEFAULT 'DRAFT' CHECK(status IN ('DRAFT', 'COMPLETED')),
        created_by TEXT NOT NULL,
        created_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS audit_results (
        result_id SERIAL PRIMARY KEY,
        audit_id INTEGER NOT NULL,
        surface_id TEXT,
        surface_name TEXT NOT NULL,
        result TEXT NOT NULL CHECK(result IN ('C', 'NC', 'NA')),
        is_additional TEXT NOT NULL DEFAULT 'N' CHECK(is_additional IN ('Y', 'N')),
        area_group TEXT,
        display_order INTEGER,
        FOREIGN KEY(audit_id) REFERENCES audits(audit_id)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS linen_cycles (
        id SERIAL PRIMARY KEY,
        cycle_name TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'DRAFT',
        created_by TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        started_at TIMESTAMP,
        completed_at TIMESTAMP
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS linen_cycle_reps (
        id SERIAL PRIMARY KEY,
        cycle_id INTEGER REFERENCES linen_cycles(id),
        rep_username TEXT NOT NULL,
        display_name TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS linen_location_assignments (
        id SERIAL PRIMARY KEY,
        cycle_id INTEGER REFERENCES linen_cycles(id),
        location_id TEXT NOT NULL,
        assigned_to TEXT NOT NULL,
        assigned_type TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE INDEX IF NOT EXISTS idx_linen_assignments_cycle
    ON linen_location_assignments(cycle_id)
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS linen_submissions (
        id SERIAL PRIMARY KEY,
        cycle_id INTEGER REFERENCES linen_cycles(id),
        location_id TEXT NOT NULL,
        submitted_by TEXT,
        status TEXT DEFAULT 'PENDING',
        submitted_at TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cur.execute("""
    CREATE INDEX IF NOT EXISTS idx_linen_submissions_cycle
    ON linen_submissions(cycle_id)
    """)

    cur.execute("""
    CREATE INDEX IF NOT EXISTS idx_linen_submissions_location
    ON linen_submissions(location_id)
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS linen_submission_lines (
        id SERIAL PRIMARY KEY,
        submission_id INTEGER NOT NULL REFERENCES linen_submissions(id),
        item_no TEXT NOT NULL,
        quantity INTEGER NOT NULL DEFAULT 0
    )
    """)

    # ---------------------------
    # Linen Manual Top Up
    # ---------------------------
    cur.execute("""
    CREATE TABLE IF NOT EXISTS linen_topups (
        topup_id SERIAL PRIMARY KEY,
        location_id TEXT NOT NULL,
        created_by TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cur.execute("""
    CREATE INDEX IF NOT EXISTS idx_linen_topups_location
    ON linen_topups(location_id)
    """)

    cur.execute("""
    CREATE INDEX IF NOT EXISTS idx_linen_topups_created_at
    ON linen_topups(created_at)
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS linen_topup_lines (
        id SERIAL PRIMARY KEY,
        topup_id INTEGER NOT NULL REFERENCES linen_topups(topup_id),
        item_no TEXT NOT NULL,
        quantity INTEGER NOT NULL CHECK(quantity > 0)
    )
    """)

    cur.execute("""
    CREATE INDEX IF NOT EXISTS idx_linen_topup_lines_topup
    ON linen_topup_lines(topup_id)
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS linen_topup_qr_tokens (
        location_id TEXT PRIMARY KEY,
        token_hash TEXT NOT NULL UNIQUE,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        active BOOLEAN NOT NULL DEFAULT TRUE
    )
    """)

    conn.commit()
    conn.close()

def seed_minimal_data():
    """
    Sync users from Master Lists workbook if possible.
    If workbook is unavailable / invalid, fall back to minimal demo users.
    """
    try:
        from master_loader import sync_users_from_workbook
        sync_users_from_workbook()
        return
    except Exception as e:
        print(f"[seed_minimal_data] Workbook sync failed, falling back to demo users: {e}")

    # fallback demo seed
    from auth import hash_password

    conn = get_conn()
    cur = conn.cursor()

    demo_users = [
        ("HenryC", "Henry", "1234", "TEAM", "B1-4", "Y"),
        ("TomA", "Tom", "1234", "TEAM", "B1-4", "Y"),
        ("DickA", "Dick", "1234", "TEAM", "B1-4", "Y"),
        ("HarryB", "Harry", "1234", "TEAM", "B5-10", "Y"),
        ("SandraB", "Sandra", "1234", "TEAM", "B5-10", "Y"),
        ("Store1", "Storeman", "1234", "STORE", "STORE", "Y"),
        ("Orville", "Orville", "1234", "ADMIN", "ADMIN", "Y"),
        ("Boss1", "Boss", "1234", "BOSS", "BOSS", "Y"),
    ]

    for username, display_name, raw_password, role, team_code, active in demo_users:
        if DB_TYPE == "postgres":
            cur.execute("SELECT 1 FROM users WHERE username = %s", (username,))
        else:
            cur.execute("SELECT 1 FROM users WHERE username = ?", (username,))
        exists = cur.fetchone()

        if not exists:
            if DB_TYPE == "postgres":
                cur.execute("""
                    INSERT INTO users (
                        username,
                        display_name,
                        password_hash,
                        role,
                        team_code,
                        active
                    )
                    VALUES (%s, %s, %s, %s, %s, %s)
                """, (
                    username,
                    display_name,
                    hash_password(raw_password),
                    role,
                    team_code,
                    active,
                ))
            else:
                cur.execute("""
                    INSERT INTO users (
                        username,
                        display_name,
                        password_hash,
                        role,
                        team_code,
                        active
                    )
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (
                    username,
                    display_name,
                    hash_password(raw_password),
                    role,
                    team_code,
                    active,
                ))

    conn.commit()
    conn.close()

def ph():
    return "%s" if DB_TYPE == "postgres" else "?"

