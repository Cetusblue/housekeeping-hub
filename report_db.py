from db import get_conn, ph
from report_config import get_report_lines, get_report_line_id_for_item
from master_loader import load_destinations_rows


# ---------------------------
# Period helpers
# ---------------------------
def get_half_year_months(period_code: str):
    """
    H1 = Jan to Jun
    H2 = Jul to Dec
    """
    period_code = period_code.upper()

    if period_code == "H1":
        return [1, 2, 3, 4, 5, 6]
    elif period_code == "H2":
        return [7, 8, 9, 10, 11, 12]
    else:
        raise ValueError("period_code must be 'H1' or 'H2'")


# ---------------------------
# Destination / bucket helpers
# ---------------------------
def _get_destination_group_lookup():
    """
    Returns:
        {
            "Annex": "ANX Blk",
            "Laundry Bay": "Others",
            ...
        }
    """
    rows = load_destinations_rows()
    return {
        r["destination_name"].strip(): r["report_group"]
        for r in rows
        if r["active"] == "Y"
    }


def get_order_bucket(team_code: str, template_day: str):
    """
    Determines report sheet bucket for normal order-issued movements.
    """
    if template_day in ("ANNEX_TUE", "ANNEX_FRI"):
        return "ANX Blk"

    # follow workbook/live naming as source of truth
    if team_code in ("A-B2-B1", "A1-3", "A4-7"):
        return "Tower A"

    if team_code in ("B1-4", "B5-10", "B11-16"):
        return "Tower B"

    if team_code in ("C1-12",):
        return "Tower C"

    return None


# ---------------------------
# Main report builder
# ---------------------------
def get_half_year_report_data(year: int, period_code: str):
    """
    Builds half-year monthly report data.

    Includes:
    - ORDER-issued stock movements
    - ADHOC issue stock movements

    Returns:
    {
        "Tower ABC": [...],
        "Tower A": [...],
        "Tower B": [...],
        "Tower C": [...],
        "ANX Blk": [...],
        "Others": [...],
    }
    """
    months = get_half_year_months(period_code)
    destination_group_lookup = _get_destination_group_lookup()

    conn = get_conn()
    cur = conn.cursor()

    # Pull all OUT movements for the selected year
    query = f"""
        SELECT
            sm.item_name,
            sm.qty,
            sm.created_at,
            sm.source_type,
            sm.source_id,
            sm.issued_to,
            o.team_code,
            o.template_day
        FROM stock_movements sm
        LEFT JOIN orders o
            ON sm.source_type = 'ORDER'
           AND sm.source_id = o.order_id
        WHERE sm.movement_type = 'OUT'
          AND COALESCE(sm.is_voided, FALSE) = FALSE
          AND EXTRACT(YEAR FROM sm.created_at::timestamp) = {ph()}
    """
    cur.execute(query, (year,))

    rows = cur.fetchall()
    conn.close()

    # Filter to selected half-year months
    filtered_rows = []
    for row in rows:
        created_at = row["created_at"]  # YYYY-MM-DD HH:MM:SS
        month = int(created_at[5:7])
        if month in months:
            filtered_rows.append(dict(row))

    report_lines = get_report_lines()

    bucket_names = ["Tower A", "Tower B", "Tower C", "ANX Blk", "Others", "Tower ABC"]

    # bucket -> report_line_id -> month -> qty
    bucket_monthly_totals = {}
    for bucket in bucket_names:
        bucket_monthly_totals[bucket] = {}
        for line in report_lines:
            rid = line["report_line_id"]
            bucket_monthly_totals[bucket][rid] = {m: 0 for m in months}

    # Fill tower-specific buckets
    for row in filtered_rows:
        item_name = row["item_name"]
        qty = int(row["qty"] or 0)
        month = int(row["created_at"][5:7])
        source_type = row["source_type"]

        report_line_id = get_report_line_id_for_item(item_name)
        if not report_line_id:
            continue

        bucket = None

        if source_type == "ORDER":
            team_code = row["team_code"]
            template_day = row["template_day"]
            bucket = get_order_bucket(team_code, template_day)

        elif source_type in ("ADHOC", "STOCK_ISSUE"):
            issued_to = (row["issued_to"] or "").strip()
            bucket = destination_group_lookup.get(issued_to, "Others")

        if not bucket:
            continue

        if bucket not in bucket_monthly_totals:
            continue

        bucket_monthly_totals[bucket][report_line_id][month] += qty

    # Build Tower ABC as sum of A + B + C + ANX Blk
    # (Others is intentionally NOT included in Tower ABC)
    for line in report_lines:
        rid = line["report_line_id"]
        for month in months:
            bucket_monthly_totals["Tower ABC"][rid][month] = (
                bucket_monthly_totals["Tower A"][rid][month]
                + bucket_monthly_totals["Tower B"][rid][month]
                + bucket_monthly_totals["Tower C"][rid][month]
                + bucket_monthly_totals["ANX Blk"][rid][month]
                + bucket_monthly_totals["Others"][rid][month]
            )

    # Convert into ordered row structures with static A:F fields
    result = {}
    for bucket in bucket_names:
        result[bucket] = []
        for line in report_lines:
            rid = line["report_line_id"]
            result[bucket].append({
                "report_line_id": line["report_line_id"],     # A
                "for_column_b": line["for_column_b"],         # B
                "report_line_name": line["report_line_name"], # C
                "report_uom": line["report_uom"],             # D
                "for_column_e": line["for_column_e"],         # E
                "for_column_f": line["for_column_f"],         # F
                "monthly_qty": bucket_monthly_totals[bucket][rid],  # G:L
            })

    return result