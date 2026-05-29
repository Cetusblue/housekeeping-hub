from db import get_conn, ph
from master_loader import load_item_master_rows

from datetime import datetime, timedelta, date
from orders_db import update_order_lines

def most_recent_sunday():
    today = date.today()
    days_since_sunday = (today.weekday() + 1) % 7
    return today - timedelta(days=days_since_sunday)

def get_packing_list_data(mode: str = "Tuesday"):
    """
    Returns:
    {
        "items": [
            {
                "item_name": ...,
                "stock": ...,
                "total_requested": ...,
                "total_issued": ...,
                "total_outstanding": ...,
                "low_stock": ...,
                "teams": {
                    "B1-4": {
                        "requested": ...,
                        "issued": ...,
                        "outstanding": ...
                    },
                    ...
                }
            }
        ],
        "team_order": [...]
    }
    """

    conn = get_conn()
    cur = conn.cursor()

    # filter by packing mode
    if mode == "TUE":
        template_days = ("TUE",)
    elif mode == "FRI":
        template_days = ("FRI",)
    elif mode == "ANNEX_TUE":
        template_days = ("ANNEX_TUE",)
    elif mode == "ANNEX_FRI":
        template_days = ("ANNEX_FRI",)
    elif mode == "TUE_COMBINED":
        template_days = ("TUE", "ANNEX_TUE")
    elif mode == "FRI_COMBINED":
        template_days = ("FRI", "ANNEX_FRI")
    elif mode == "OT":
        template_days = ("OT",)
    else:
        template_days = ()

    placeholders = ",".join([ph()] * len(template_days))

    query = f"""
        SELECT
            ol.item_name,
            o.team_code,
            SUM(ol.qty_requested) AS requested,
            SUM(ol.qty_issued) AS issued,
            SUM(ol.qty_requested - ol.qty_issued) AS outstanding
        FROM order_lines ol
        JOIN orders o ON ol.order_id = o.order_id
        WHERE o.status IN ('PENDING', 'PARTIALLY_ISSUED')
          AND o.run_date >= {ph()}
    """

    cutoff_date = most_recent_sunday().isoformat()
    params = [cutoff_date]

    if template_days:
        query += f"\n AND o.template_day IN ({placeholders})"
        params.extend(template_days)

    query += """
        GROUP BY ol.item_name, o.team_code
        ORDER BY ol.item_name, o.team_code
    """

    cur.execute(query, params)
    rows = cur.fetchall()

    # stock lookup
    cur.execute("""
        SELECT
            item_name,
            SUM(CASE WHEN movement_type = 'IN' THEN qty ELSE 0 END) -
            SUM(CASE WHEN movement_type = 'OUT' THEN qty ELSE 0 END) AS stock
        FROM stock_movements
        WHERE COALESCE(is_voided, FALSE) = FALSE
        GROUP BY item_name
    """)
    stock_rows = cur.fetchall()
    conn.close()

    stock_lookup = {r["item_name"]: int(r["stock"] or 0) for r in stock_rows}

    # build structure
    item_map = {}
    team_set = set()

    for r in rows:
        item = r["item_name"]
        team = r["team_code"]
        requested = int(r["requested"] or 0)
        issued = int(r["issued"] or 0)
        outstanding = int(r["outstanding"] or 0)

        if requested <= 0:
            continue

        team_set.add(team)

        if item not in item_map:
            item_map[item] = {
                "item_name": item,
                "stock": stock_lookup.get(item, 0),
                "teams": {},
            }

        item_map[item]["teams"][team] = {
            "requested": requested,
            "issued": issued,
            "outstanding": outstanding,
        }

    # enrich totals
    for item_name, item_data in item_map.items():
        total_requested = sum(v["requested"] for v in item_data["teams"].values())
        total_issued = sum(v["issued"] for v in item_data["teams"].values())
        total_outstanding = sum(v["outstanding"] for v in item_data["teams"].values())
        low_stock = item_data["stock"] < total_outstanding

        item_data["total_requested"] = total_requested
        item_data["total_issued"] = total_issued
        item_data["total_outstanding"] = total_outstanding
        item_data["low_stock"] = low_stock

    # team order
    team_order = sorted(team_set)

    # sort items by Item Master display_order
    item_master = load_item_master_rows()
    order_map = {i["item_name"]: i["display_order"] for i in item_master}

    items = list(item_map.values())
    items.sort(key=lambda x: order_map.get(x["item_name"], 9999))

    return {
        "items": items,
        "team_order": team_order,
    }

def _now_iso():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_team_item_order_lines(mode: str, team_code: str, item_name: str):
    """
    Returns open order lines for a given mode/team/item, oldest order first.
    Used for FIFO distribution.
    """
    conn = get_conn()
    cur = conn.cursor()

    if mode == "Tuesday":
        template_days = ("TUE",)
    elif mode == "Friday":
        template_days = ("FRI",)
    elif mode == "ANNEX_TUE":
        template_days = ("ANNEX_TUE",)
    elif mode == "ANNEX_FRI":
        template_days = ("ANNEX_FRI",)
    elif mode == "OT":
        template_days = ("OT",)
    else:
        template_days = ()

    placeholders = ",".join([ph()] * len(template_days))

    query = f"""
        SELECT
            o.order_id,
            o.created_at,
            ol.line_id,
            ol.item_name,
            ol.qty_requested,
            ol.qty_issued
        FROM orders o
        JOIN order_lines ol
            ON o.order_id = ol.order_id
        WHERE o.status IN ('PENDING', 'PARTIALLY_ISSUED')
          AND o.run_date >= {ph()}
          AND o.team_code = {ph()}
          AND ol.item_name = {ph()}
    """
    cutoff_date = most_recent_sunday().isoformat()
    params = [cutoff_date, team_code, item_name]

    if template_days:
        query += f"\n AND o.template_day IN ({placeholders})"
        params.extend(template_days)

    qquery += """
    ORDER BY o.created_at ASC, o.order_id ASC, ol.line_id ASC
    """

    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()

    return [dict(r) for r in rows]


def save_packing_board_issued(mode: str, edited_rows: list[dict], updated_by: str):
    """
    Saves issued quantities from packing board using FIFO auto-distribution.

    edited_rows is the current packing board table, containing:
    - Item
    - Team Req
    - Team Iss
    for each visible team
    """

    # 1. Build desired totals per item/team from edited grid
    desired_by_item_team = {}
    visible_teams = set()

    for row in edited_rows:
        item_name = row["Item"]
        desired_by_item_team[item_name] = {}

        for key, value in row.items():
            if key.endswith(" Req"):
                team = key[:-4]
                visible_teams.add(team)

        for team in visible_teams:
            iss_col = f"{team} Iss"
            req_col = f"{team} Req"

            desired_iss = int(row.get(iss_col, 0) or 0)
            requested = int(row.get(req_col, 0) or 0)

            # clamp desired issued to requested total
            if desired_iss > requested:
                desired_iss = requested
            if desired_iss < 0:
                desired_iss = 0

            desired_by_item_team[item_name][team] = desired_iss

    # 2. For each item/team, distribute desired cumulative issued across open order lines FIFO
    grouped_updates = {}  # order_id -> list[{line_id, qty_requested, qty_issued}]

    for item_name, team_map in desired_by_item_team.items():
        for team_code, desired_total_issued in team_map.items():
            fifo_lines = get_team_item_order_lines(mode, team_code, item_name)

            if not fifo_lines:
                continue

            remaining_to_allocate = desired_total_issued

            for line in fifo_lines:
                line_id = int(line["line_id"])
                order_id = int(line["order_id"])
                qty_requested = int(line["qty_requested"] or 0)

                new_issued_for_line = min(qty_requested, remaining_to_allocate)
                remaining_to_allocate -= new_issued_for_line

                grouped_updates.setdefault(order_id, []).append({
                    "line_id": line_id,
                    "qty_requested": qty_requested,
                    "qty_issued": new_issued_for_line,
                })

            # if remaining_to_allocate > 0, it just means desired grid value exceeded
            # total request across open lines; already effectively clamped by distribution

    # 3. Save order by order using existing delta-aware order updater
    for order_id, rows_to_save in grouped_updates.items():
        update_order_lines(order_id, rows_to_save, updated_by=updated_by)

    return True

    