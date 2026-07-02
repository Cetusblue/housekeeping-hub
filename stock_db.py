from datetime import datetime
from db import get_conn, ph
from templates import get_template_items
from master_loader import get_item_master_lookup


# ---------------------------
# Helpers
# ---------------------------
def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _all_template_items():
    """
    Returns a de-duplicated list of all template items from TUE/FRI/OT.
    Output format:
        [{"item_name": "..."}]
    """
    seen = set()
    items = []

    for template_day in ("TUE", "FRI", "OT"):
        for it in get_template_items(template_day):
            if isinstance(it, dict):
                item_name = str(it["item_name"])
            else:
                item_name = str(it[1])

            if item_name not in seen:
                seen.add(item_name)
                items.append({"item_name": item_name})

    return items


# ---------------------------
# Stock movement creation
# ---------------------------
def create_stock_in(item_name: str, qty: int, created_by: str, created_at=None):
    """
    Records a stock-in movement.
    """
    if qty <= 0:
        raise ValueError("Stock-in quantity must be greater than 0.")

    conn = get_conn()
    cur = conn.cursor()

    query = f"""
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
        VALUES ({ph()}, 'IN', {ph()}, NULL, 'STOCK_IN', NULL, {ph()}, {ph()})
    """ 
    cur.execute(query, (
        item_name,
        int(qty),
        created_by,
        created_at or now_iso()
    ))

    conn.commit()
    conn.close()


def create_adhoc_issue(item_name: str, qty: int, issued_to: str, created_by: str, created_at=None):
    """
    Records an adhoc stock-out movement.
    Enforces no negative stock.
    """
    if qty <= 0:
        raise ValueError("Issue quantity must be greater than 0.")

    current = get_current_stock(item_name)
    if current < qty:
        raise ValueError(
            f"Insufficient stock for {item_name}. Current: {current}, requested issue: {qty}"
        )

    conn = get_conn()
    cur = conn.cursor()

    query = f"""
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
        VALUES ({ph()}, 'OUT', {ph()}, {ph()}, 'ADHOC', NULL, {ph()}, {ph()})
    """
    cur.execute(query, (
        item_name,
        int(qty),
        issued_to,
        created_by,
        created_at or now_iso()
    ))

    conn.commit()
    conn.close()


def create_adhoc_issue_batch(issue_rows: list[dict], issued_to: str, created_by: str, created_at=None):
    """
    Creates a batch of adhoc issue movements.
    Strict mode:
      - if any row would go negative, block the whole batch.
    Expects:
      [{"item_name": "...", "qty": 2}, ...]
    """
    ok, errors = validate_issue_quantities(issue_rows)
    if not ok:
        raise ValueError("Insufficient stock:\n" + "\n".join(errors))

    conn = get_conn()
    cur = conn.cursor()

    for row in issue_rows:
        item_name = row["item_name"]
        qty = int(row["qty"])

        if qty <= 0:
            continue

        query = f"""
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
            VALUES ({ph()}, 'OUT', {ph()}, {ph()}, 'ADHOC', NULL, {ph()}, {ph()})
        """ 
        cur.execute(query, (
            item_name,
            qty,
            issued_to,
            created_by,
            created_at or now_iso()
        ))

    conn.commit()
    conn.close()


# ---------------------------
# Current stock calculations
# ---------------------------
def get_current_stock(item_name: str) -> int:
    """
    Returns current stock balance for one item.
    Balance = total IN - total OUT
    """
    conn = get_conn()
    cur = conn.cursor()

    query = f"""
        SELECT
            COALESCE(SUM(CASE WHEN movement_type = 'IN' THEN qty ELSE 0 END), 0) AS total_in,
            COALESCE(SUM(CASE WHEN movement_type = 'OUT' THEN qty ELSE 0 END), 0) AS total_out
        FROM stock_movements
        WHERE item_name = {ph()}
          AND COALESCE(is_voided, FALSE) = FALSE
    """
    cur.execute(query, (item_name,))

    row = cur.fetchone()
    conn.close()

    total_in = int(row["total_in"] or 0)
    total_out = int(row["total_out"] or 0)
    return total_in - total_out


def get_inventory_rows():
    """
    Returns inventory-enabled items with current stock.
    Uses Item Master Inventory = Y.
    """
    item_lookup = get_item_master_lookup()
    rows = []

    for item_name, info in item_lookup.items():
        if str(info.get("Inventory", "")).upper() != "Y":
            continue

        rows.append({
            "item_name": item_name,
            "category": info.get("category", "Others"),
            "display_order": info.get("display_order", 9999),
            "current_stock": get_current_stock(item_name),
        })

    rows.sort(key=lambda x: (
        x.get("category", "Others"),
        int(x.get("display_order", 9999)),
        x.get("item_name", "")
    ))

    return rows


# ---------------------------
# Validation helpers
# ---------------------------
def validate_issue_quantities(issue_rows: list[dict]):
    """
    Validates a batch of adhoc issue rows before committing anything.
    Expects:
        [{"item_name": "...", "qty": 3}, ...]
    Returns:
        (True, [])
        or
        (False, [error messages])
    """
    errors = []

    for row in issue_rows:
        item_name = row["item_name"]
        qty = int(row["qty"])

        if qty <= 0:
            continue

        current = get_current_stock(item_name)
        if current < qty:
            errors.append(
                f"{item_name}: current stock {current}, attempted issue {qty}"
            )

    return (len(errors) == 0, errors)


# ---------------------------
# Movement history
# ---------------------------
def get_stock_movements_for_item(item_name: str, date_from=None, date_to=None):
    """
    Returns movement history for one item, sorted oldest first.
    """
    conn = get_conn()
    cur = conn.cursor()

    query = f"""
        SELECT movement_id, item_name, movement_type, qty, issued_to,
               source_type, source_id, created_by, created_at
        FROM stock_movements
        WHERE item_name = {ph()}
          AND COALESCE(is_voided, FALSE) = FALSE
    """
    params = [item_name]

    if date_from:
        query += f" AND created_at::timestamp >= {ph()}::date"
        params.append(date_from)

    if date_to:
        query += f" AND created_at::timestamp < {ph()}::date"
        params.append(date_to)

    query += " ORDER BY created_at ASC, movement_id ASC"

    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()

    return [dict(r) for r in rows]

def search_stock_movements(item_name="", date_from=None, date_to=None):
    from db import get_conn, ph

    conn = get_conn()
    cur = conn.cursor()

    query = f"""
        SELECT
            movement_id,
            item_name,
            movement_type,
            qty,
            issued_to,
            source_type,
            source_id,
            created_by,
            created_at,
            COALESCE(is_voided, FALSE) AS is_voided
        FROM stock_movements
        WHERE 1=1
    """

    params = []

    if item_name.strip():
        query += f" AND item_name ILIKE {ph()}"
        params.append(f"%{item_name.strip()}%")

    if date_from:
        query += f" AND created_at::date >= {ph()}"
        params.append(str(date_from))

    if date_to:
        query += f" AND created_at::date <= {ph()}"
        params.append(str(date_to))

    query += " ORDER BY created_at DESC, movement_id DESC LIMIT 100"

    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()

    return [dict(r) for r in rows]

def void_stock_movement(movement_id: int, voided_by: str, void_reason: str):
    from db import get_conn, ph

    conn = get_conn()
    cur = conn.cursor()

    query = f"""
        UPDATE stock_movements
        SET is_voided = TRUE,
            voided_at = CURRENT_TIMESTAMP,
            voided_by = {ph()},
            void_reason = {ph()}
        WHERE movement_id = {ph()}
          AND COALESCE(is_voided, FALSE) = FALSE
    """

    cur.execute(query, (voided_by, void_reason, movement_id))
    rows_updated = cur.rowcount

    conn.commit()
    conn.close()

    return rows_updated

def get_stock_card_rows(item_name: str, date_from=None, date_to=None):
    """
    Returns stock card rows with running balance for a single item.
    Output columns:
      Date, Stock In, Stock Out, Balance, Issued To
    """
    movements = get_stock_movements_for_item(item_name, date_from=date_from, date_to=date_to)

    from datetime import datetime

    month_key = None

    if date_from:
        month_key = str(date_from)[:7]
    elif date_to:
        month_key = str(date_to)[:7]

    opening_balance = 0

    if month_key:
        override = get_opening_balance_override(
            item_name,
            month_key
        )

        if override is not None:
            opening_balance = override
        else:
            conn = get_conn()
            cur = conn.cursor()

            query = f"""
                SELECT
                    COALESCE(
                        SUM(
                            CASE
                                WHEN movement_type = 'IN'
                                THEN qty
                                ELSE -qty
                            END
                        ),
                        0
                    ) AS opening_balance
                FROM stock_movements
                WHERE item_name = {ph()}
                AND COALESCE(is_voided, FALSE) = FALSE
                AND created_at::timestamp < {ph()}::date
            """

            cur.execute(
                query,
                (
                    item_name,
                    f"{month_key}-01"
                )
            )

            row = cur.fetchone()
            opening_balance = int(row["opening_balance"] or 0)

            conn.close()

    balance = opening_balance
    rows = []

    if month_key:
        opening_date = datetime.strptime(
            f"{month_key}-01",
            "%Y-%m-%d"
        )

        rows.append({
            "Date": f"{opening_date.day}-{opening_date.strftime('%b-%Y')}",
            "Stock In": "",
            "Stock Out": "",
            "Balance": balance,
            "Issued To": "",
            "Remarks": "Opening Balance"
        })

    for m in movements:
        qty_in = int(m["qty"]) if m["movement_type"] == "IN" else 0
        qty_out = int(m["qty"]) if m["movement_type"] == "OUT" else 0

        balance += qty_in
        balance -= qty_out

        movement_date = datetime.strptime(
        m["created_at"][:10],
        "%Y-%m-%d"
    )

        rows.append({
            "Date": f"{movement_date.day}-{movement_date.strftime('%b-%Y')}",
            "Stock In": qty_in if qty_in > 0 else "",
            "Stock Out": qty_out if qty_out > 0 else "",
            "Balance": balance,
            "Issued To": m["issued_to"] or "",
            "Remarks": ""
        })

    return rows

def get_opening_balance_override(item_name: str, month_key: str):
    conn = get_conn()
    cur = conn.cursor()

    query = f"""
        SELECT opening_balance
        FROM stock_opening_balances
        WHERE item_name = {ph()}
          AND month_key = {ph()}
    """

    cur.execute(query, (item_name, month_key))
    row = cur.fetchone()
    conn.close()

    if not row:
        return None

    return int(row["opening_balance"])


def save_opening_balance_override(item_name: str, month_key: str, opening_balance: int):
    conn = get_conn()
    cur = conn.cursor()

    query = f"""
        INSERT INTO stock_opening_balances (
            item_name,
            month_key,
            opening_balance
        )
        VALUES ({ph()}, {ph()}, {ph()})
        ON CONFLICT (item_name, month_key)
        DO UPDATE SET opening_balance = EXCLUDED.opening_balance
    """

    cur.execute(query, (item_name, month_key, int(opening_balance)))

    conn.commit()
    conn.close()

def update_stock_movement_date(movement_id: int, new_date: str):
    conn = get_conn()
    cur = conn.cursor()

    query = f"""
        UPDATE stock_movements
        SET created_at = {ph()}
        WHERE movement_id = {ph()}
          AND COALESCE(is_voided, FALSE) = FALSE
    """

    cur.execute(query, (new_date, movement_id))
    rows_updated = cur.rowcount

    conn.commit()
    conn.close()

    return rows_updated