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

def get_manual_topup_month_rows(year, month):
    """
    Returns individual Manual Top Up item quantities for one month.

    Aggregation into report tabs / report item mappings is deliberately
    handled outside SQL because those mappings live in Master Lists.xlsx.
    """

    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            t.location_id,
            t.created_at,
            l.item_no,
            l.quantity
        FROM linen_topups t
        JOIN linen_topup_lines l
            ON l.topup_id = t.topup_id
        WHERE EXTRACT(YEAR FROM t.created_at) = %s
          AND EXTRACT(MONTH FROM t.created_at) = %s
        ORDER BY
            t.created_at,
            t.topup_id,
            l.id
    """, (
        int(year),
        int(month),
    ))

    rows = [dict(row) for row in cur.fetchall()]
    conn.close()

    return rows

def build_manual_topup_month_report(year, month):
    from master_loader import (
        load_linen_master_rows,
        get_linen_location_map,
    )

    transaction_rows = get_manual_topup_month_rows(
        year,
        month,
    )

    linen_items = load_linen_master_rows()
    location_map = get_linen_location_map()

    # --------------------------------
    # Item No -> report mapping
    # --------------------------------
    item_report_map = {}

    for item in linen_items:
        report_name = str(
            item.get("man_report_mapping") or ""
        ).strip()

        report_no = item.get("man_report_no")

        if not report_name or report_no is None:
            continue

        item_report_map[str(item["item_no"])] = {
            "report_no": int(report_no),
            "report_name": report_name,
        }

    # --------------------------------
    # Fixed report rows
    # --------------------------------
    fixed_rows = {
        2: "CURTAIN",
        3: "CURTAIN SHOWER",
        9: "LAB COAT",
        36: "MATRESS PROTECTOR",
        37: "BLUE TOWEL",
    }

    # --------------------------------
    # Build canonical 1-39 report list
    # --------------------------------
    report_lines = {}

    for mapping in item_report_map.values():
        report_no = mapping["report_no"]
        report_name = mapping["report_name"]

        # Several Linen Master items are allowed to share
        # the same report number, but they must agree on its name.
        if report_no in report_lines:
            if report_lines[report_no] != report_name:
                raise ValueError(
                    f"Manual report no. {report_no} has conflicting "
                    f"names: '{report_lines[report_no]}' and "
                    f"'{report_name}'."
                )
        else:
            report_lines[report_no] = report_name

    for report_no, report_name in fixed_rows.items():
        if report_no in report_lines:
            raise ValueError(
                f"Hard-coded report no. {report_no} conflicts "
                f"with Linen Master mapping."
            )

        report_lines[report_no] = report_name

    missing_numbers = [
        n for n in range(1, 40)
        if n not in report_lines
    ]

    if missing_numbers:
        raise ValueError(
            "Manual Top Up report is missing report number(s): "
            + ", ".join(str(n) for n in missing_numbers)
        )

    # --------------------------------
    # Output:
    #
    # {
    #   "L1 Outpatient DI": {
    #       "tower": "A",
    #       "rows": {
    #           1: {
    #               "name": "...",
    #               "days": {1: 10, 2: 5, ...}
    #           }
    #       }
    #   }
    # }
    # --------------------------------
    report = {}

    # First create every report tab from Location Master,
    # even if there were no top-ups that month.
    for location_id, location in location_map.items():

        if str(
            location.get("manual_topup") or ""
        ).strip().upper() != "Y":
            continue

        sheet_name = str(
            location.get("manual_report_mapping") or ""
        ).strip()

        if not sheet_name:
            continue

        tower = str(
            location.get("tower") or ""
        ).strip().upper()

        if sheet_name not in report:
            report[sheet_name] = {
                "tower": tower,
                "rows": {
                    report_no: {
                        "name": report_lines[report_no],
                        "days": {},
                    }
                    for report_no in range(1, 40)
                },
            }

        else:
            existing_tower = report[sheet_name]["tower"]

            if (
                existing_tower
                and tower
                and existing_tower != tower
            ):
                raise ValueError(
                    f"Report tab '{sheet_name}' contains locations "
                    f"from more than one tower."
                )

    # --------------------------------
    # Aggregate transactions
    # --------------------------------
    for transaction in transaction_rows:

        location_id = transaction["location_id"]
        item_no = str(transaction["item_no"])

        location = location_map.get(location_id)

        if not location:
            continue

        sheet_name = str(
            location.get("manual_report_mapping") or ""
        ).strip()

        if not sheet_name or sheet_name not in report:
            continue

        item_mapping = item_report_map.get(item_no)

        # Linen items intentionally not mapped into this report
        # are ignored.
        if not item_mapping:
            continue

        report_no = item_mapping["report_no"]

        created_at = transaction["created_at"]

        if hasattr(created_at, "day"):
            day = created_at.day
        else:
            day = int(str(created_at)[8:10])

        quantity = int(
            transaction.get("quantity") or 0
        )

        current = report[
            sheet_name
        ]["rows"][
            report_no
        ]["days"].get(day, 0)

        report[
            sheet_name
        ]["rows"][
            report_no
        ]["days"][day] = current + quantity

    return report

def purge_old_manual_topups():
    """
    Retains approximately the most recent 3 years
    of Manual Top Up transaction data.
    """

    conn = get_conn()
    cur = conn.cursor()

    try:
        # Delete child rows first because linen_topup_lines
        # references linen_topups.
        cur.execute("""
            DELETE FROM linen_topup_lines
            WHERE topup_id IN (
                SELECT topup_id
                FROM linen_topups
                WHERE created_at < CURRENT_TIMESTAMP - INTERVAL '3 years'
            )
        """)

        cur.execute("""
            DELETE FROM linen_topups
            WHERE created_at < CURRENT_TIMESTAMP - INTERVAL '3 years'
        """)

        conn.commit()

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()