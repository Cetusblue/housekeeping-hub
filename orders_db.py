from datetime import datetime, date
from db import get_conn, ph, DB_TYPE, now_iso
from templates import get_template_items
from stock_db import get_current_stock

# ---------------------------
# Helpers
# ---------------------------
def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def create_order_issue_movement(item_name: str, qty: int, issued_to: str, created_by: str, order_id: int):
    """
    Records an ORDER-based stock-out movement.
    """
    if qty <= 0:
        return

    conn = get_conn()
    cur = conn.cursor()

    cur.execute(f"""
        INSERT INTO stock_movements (
            item_name,
            movement_type,
            qty,
            issued_to,
            source_type,
            source_id,
            created_by,
            created_at
        )
        VALUES ({ph()}, 'OUT', {ph()}, {ph()}, 'ORDER', {ph()}, {ph()}, {ph()})
    """, (
        item_name,
        int(qty),
        issued_to,
        int(order_id),
        created_by,
        now_iso()
    ))

    conn.commit()
    conn.close()

# ---------------------------
# Orders: lookup / listing
# ---------------------------
def get_order_by_unique(team_code: str, template_day: str, run_date: str):
    """
    Returns the existing order for a team + template_day + run_date, or None.
    Used to prevent duplicate orders for the same group and issue date.
    """
    conn = get_conn()
    cur = conn.cursor()

    cur.execute(f"""
        SELECT order_id, team_code, template_day, run_date, status,
               created_by, created_at, issued_by, issued_at,
               closed_reason, closed_by, closed_at
        FROM orders
        WHERE team_code = {ph()} AND template_day = {ph()} AND run_date = {ph()} AND status != 'CANCELLED'
        LIMIT 1
    """, (team_code, template_day, run_date))

    row = cur.fetchone()
    conn.close()

    if not row:
        return None

    return dict(row)


def get_order(order_id: int):
    conn = get_conn()
    cur = conn.cursor()

    query = f"""
        SELECT order_id, team_code, template_day, run_date, status,
               created_by, created_at, issued_by, issued_at,
               closed_reason, closed_by, closed_at
        FROM orders
        WHERE order_id = {ph()}
    """
    cur.execute(query, (order_id,))

    row = cur.fetchone()
    conn.close()

    if not row:
        return None

    return dict(row)


def _recent_order_cutoff() -> str:
    """Returns the first day of the previous calendar month as YYYY-MM-DD."""
    today = date.today()

    if today.month == 1:
        cutoff = date(today.year - 1, 12, 1)
    else:
        cutoff = date(today.year, today.month - 1, 1)

    return cutoff.isoformat()


def list_orders_for_team(team_code: str):
    """Lists only orders whose run_date is in the current or previous month."""
    conn = get_conn()
    cur = conn.cursor()

    query = f"""
        SELECT order_id, team_code, template_day, run_date, status,
               created_by, created_at, issued_by, issued_at,
               closed_by, closed_at, closed_reason
        FROM orders
        WHERE team_code = {ph()}
          AND run_date >= {ph()}
        ORDER BY run_date DESC, order_id DESC
    """
    cur.execute(query, (team_code, _recent_order_cutoff()))

    rows = cur.fetchall()
    conn.close()

    return [dict(r) for r in rows]


def list_orders_for_store(status=None, recent_only: bool = True):
    """
    Lists orders for Store/Admin.

    Store should use recent_only=True to show the current and previous month.
    Admin can use recent_only=False to retain full history.
    """
    conn = get_conn()
    cur = conn.cursor()

    where_clauses = []
    params = []

    if status:
        where_clauses.append(f"status = {ph()}")
        params.append(status)

    if recent_only:
        where_clauses.append(f"run_date >= {ph()}")
        params.append(_recent_order_cutoff())

    where_sql = ""
    if where_clauses:
        where_sql = "WHERE " + " AND ".join(where_clauses)

    query = f"""
        SELECT order_id, team_code, template_day, run_date, status,
               created_by, created_at, issued_by, issued_at,
               closed_by, closed_at, closed_reason
        FROM orders
        {where_sql}
        ORDER BY
            CASE status
                WHEN 'PENDING' THEN 1
                WHEN 'PARTIALLY_ISSUED' THEN 2
                WHEN 'ISSUED' THEN 3
                WHEN 'CLOSED' THEN 4
                WHEN 'CANCELLED' THEN 5
                ELSE 99
            END,
            run_date DESC,
            order_id DESC
    """

    cur.execute(query, tuple(params))

    rows = cur.fetchall()
    conn.close()

    return [dict(r) for r in rows]


# ---------------------------
# Order creation
# ---------------------------
def create_order(team_code: str, template_day: str, run_date: str, created_by: str) -> int:
    """
    Creates an order header and seeds order_lines from the template list.
    """
    conn = get_conn()
    cur = conn.cursor()

    created_at = now_iso()

    if DB_TYPE == "postgres":
        cur.execute(f"""
            INSERT INTO orders (
                team_code,
                template_day,
                run_date,
                status,
                created_by,
                created_at
            )
            VALUES ({ph()}, {ph()}, {ph()}, 'PENDING', {ph()}, {ph()})
            RETURNING order_id
        """, (team_code, template_day, run_date, created_by, created_at))
        order_id = cur.fetchone()["order_id"]
    else:
        cur.execute(f"""
            INSERT INTO orders (
                team_code,
                template_day,
                run_date,
                status,
                created_by,
                created_at
            )
            VALUES ({ph()}, {ph()}, {ph()}, 'PENDING', {ph()}, {ph()})
        """, (team_code, template_day, run_date, created_by, created_at))
        order_id = cur.lastrowid

    # Seed order lines from template
    items = get_template_items(template_day)

    for it in items:
        # Support tuples like: (item_no, item_name)
        if isinstance(it, dict):
            item_no = int(it["item_no"])
            item_name = str(it["item_name"])
        else:
            item_no = int(it[0])
            item_name = str(it[1])

        cur.execute(f"""
            INSERT INTO order_lines (
                order_id,
                item_no,
                item_name,
                qty_requested,
                qty_issued
            )
            VALUES ({ph()}, {ph()}, {ph()}, 0, 0)
        """, (order_id, item_no, item_name))

    conn.commit()
    conn.close()

    return order_id

def cancel_order(order_id, cancelled_by, cancel_reason):
    from db import get_conn

    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        UPDATE orders
        SET status = 'CANCELLED',
            cancelled_at = NOW(),
            cancelled_by = %s,
            cancel_reason = %s
        WHERE order_id = %s
          AND status = 'PENDING'
    """, (cancelled_by, cancel_reason, order_id))

    rows_updated = cur.rowcount

    conn.commit()
    cur.close()
    conn.close()

    return rows_updated

def close_order(order_id: int, closed_by: str, closed_reason: str):
    """
    Closes a partially issued order without issuing the remaining quantities.

    The requested and issued quantities are preserved for audit purposes.
    Only PARTIALLY_ISSUED orders may be closed.
    """
    reason = (closed_reason or "").strip()
    if not reason:
        raise ValueError("A closure reason is required.")

    conn = get_conn()
    cur = conn.cursor()

    cur.execute(f"""
        UPDATE orders
        SET status = 'CLOSED',
            closed_reason = {ph()},
            closed_by = {ph()},
            closed_at = {ph()}
        WHERE order_id = {ph()}
          AND status = 'PARTIALLY_ISSUED'
    """, (
        reason,
        closed_by,
        now_iso(),
        int(order_id),
    ))

    rows_updated = cur.rowcount
    conn.commit()
    conn.close()
    return rows_updated


# ---------------------------
# Order lines
# ---------------------------
def get_order_lines(order_id: int, template_day=None):
    """
    Returns order lines for an order.
    template_day is accepted for compatibility but not needed now.
    """
    conn = get_conn()
    cur = conn.cursor()

    query = f"""
        SELECT line_id, order_id, item_no, item_name, qty_requested, qty_issued
        FROM order_lines
        WHERE order_id = {ph()}
        ORDER BY item_no
    """
    cur.execute(query, (order_id,))

    rows = cur.fetchall()
    conn.close()

    return [dict(r) for r in rows]


def update_order_lines(order_id: int, rows_to_save: list[dict], updated_by: str = ""):
    """
    Updates request / issued quantities and recalculates order status.

    Also writes ORDER stock movements for any positive delta in qty_issued.
    Blocks save if any positive delta would cause stock to go below zero.
    Expected row format:
    {
        "line_id": ...,
        "qty_requested": ...,
        "qty_issued": ...
    }
    """
    conn = get_conn()
    cur = conn.cursor()

    # Get order header first
    query = f"SELECT * FROM orders WHERE order_id = {ph()}"
    cur.execute(query, (order_id,))
    order_row = cur.fetchone()

    if not order_row:
        conn.close()
        raise ValueError(f"Order {order_id} not found.")

    team_code = order_row["team_code"]

    # Get original line values before update
    query = f"""
        SELECT line_id, order_id, item_no, item_name, qty_requested, qty_issued
        FROM order_lines
        WHERE order_id = {ph()}
    """
    cur.execute(query, (order_id,))
    original_lines = cur.fetchall()

    original_lookup = {int(r["line_id"]): dict(r) for r in original_lines}

    # -----------------------
    # First pass: validate inputs and accumulate positive deltas by item
    # -----------------------
    positive_deltas_by_item = {}

    validated_rows = []

    for row in rows_to_save:
        lid = int(row["line_id"])
        new_req = int(row["qty_requested"])
        new_iss = int(row["qty_issued"])

        if lid not in original_lookup:
            continue

        orig = original_lookup[lid]
        prev_iss = int(orig["qty_issued"] or 0)
        item_name = orig["item_name"]

        # basic validation
        if new_req < 0:
            conn.close()
            raise ValueError(f"{item_name}: requested quantity cannot be negative.")

        if new_iss < 0:
            conn.close()
            raise ValueError(f"{item_name}: issued quantity cannot be negative.")

        # do not allow issued > requested on order lines
        if new_iss > new_req:
            conn.close()
            raise ValueError(f"{item_name}: issued quantity cannot exceed requested quantity.")

        delta_issued = new_iss - prev_iss

        if delta_issued > 0:
            positive_deltas_by_item[item_name] = positive_deltas_by_item.get(item_name, 0) + delta_issued

        validated_rows.append({
            "line_id": lid,
            "item_name": item_name,
            "qty_requested": new_req,
            "qty_issued": new_iss,
            "prev_issued": prev_iss,
            "delta_issued": delta_issued,
        })

    # -----------------------
    # Stock validation: block if any positive delta exceeds current stock
    # -----------------------
    stock_errors = []
    for item_name, total_positive_delta in positive_deltas_by_item.items():
        current_stock = int(get_current_stock(item_name) or 0)
        if total_positive_delta > current_stock:
            stock_errors.append(
                f"{item_name}: stock {current_stock}, attempted additional issue {total_positive_delta}"
            )

    if stock_errors:
        conn.close()
        raise ValueError("Insufficient stock:\n" + "\n".join(stock_errors))

    # -----------------------
    # Second pass: apply updates and write movements
    # -----------------------
    for row in validated_rows:
        lid = row["line_id"]
        item_name = row["item_name"]
        new_req = row["qty_requested"]
        new_iss = row["qty_issued"]
        delta_issued = row["delta_issued"]

        # only positive delta should create ORDER stock movements
        if delta_issued > 0:
            cur.execute(f"""
                INSERT INTO stock_movements (
                    item_name,
                    movement_type,
                    qty,
                    issued_to,
                    source_type,
                    source_id,
                    created_by,
                    created_at
                )
                VALUES ({ph()}, 'OUT', {ph()}, {ph()}, 'ORDER', {ph()}, {ph()}, {ph()})
            """, (
                item_name,
                delta_issued,
                team_code,
                int(order_id),
                updated_by or "SYSTEM",
                now_iso()
            ))

        cur.execute(f"""
            UPDATE order_lines
            SET qty_requested = {ph()}, qty_issued = {ph()}
            WHERE order_id = {ph()} AND line_id = {ph()}
        """, (
            new_req,
            new_iss,
            order_id,
            lid,
        ))

    # -----------------------
    # Recalculate status after saving
    # -----------------------
    query = f"""
        SELECT line_id, order_id, item_no, item_name, qty_requested, qty_issued
        FROM order_lines
        WHERE order_id = {ph()}
    """
    cur.execute(query, (order_id,))
    lines = cur.fetchall()

    if not lines:
        status = "PENDING"
    else:
        any_issued = any(int(r["qty_issued"]) > 0 for r in lines)
        all_fully_issued = all(int(r["qty_issued"]) >= int(r["qty_requested"]) for r in lines)

        if not any_issued:
            status = "PENDING"
        elif all_fully_issued:
            status = "ISSUED"
        else:
            status = "PARTIALLY_ISSUED"

    if status in ("PARTIALLY_ISSUED", "ISSUED"):
        cur.execute(f"""
            UPDATE orders
            SET status = {ph()},
                issued_by = {ph()},
                issued_at = {ph()}
            WHERE order_id = {ph()}
        """, (
            status,
            updated_by or None,
            now_iso(),
            order_id
        ))
    else:
        cur.execute(f"""
            UPDATE orders
            SET status = {ph()}
            WHERE order_id = {ph()}
        """, (
            status,
            order_id
        ))
    conn.commit()
    conn.close()

# ---------------------------
# Optional helper (legacy compatibility)
# ---------------------------
def mark_issued(order_id: int, issued_by: str):
    """
    Optional compatibility function.
    In the newer model, Save Issued auto-updates status.
    This function can still be used if you need it somewhere.
    """
    conn = get_conn()
    cur = conn.cursor()

    cur.execute(f"""
        UPDATE orders
        SET status = 'ISSUED',
            issued_by = {ph()},
            issued_at = {ph()}
        WHERE order_id = {ph()}
    """, (issued_by, now_iso(), order_id))

    conn.commit()
    conn.close()