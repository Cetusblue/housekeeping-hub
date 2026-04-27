from master_loader import load_item_master_rows


def get_template_items(template_day: str):
    """
    Returns order template items from Item Master in Master Lists.xlsx.

    Output format expected by orders_db.py:
        [
            (display_order, item_name),
            ...
        ]

    template_day:
    - TUE
    - FRI
    - OT
    """
    template_day = template_day.upper()
    rows = load_item_master_rows()

    items = []

    for r in rows:
        if template_day in ("TUE", "ANNEX_TUE") and r["Tue"] == "Y":
            items.append((r["display_order"], r["item_name"]))

        elif template_day in ("FRI", "ANNEX_FRI") and r["Fri"] == "Y":
            items.append((r["display_order"], r["item_name"]))

        elif template_day == "OT" and r["OT"] == "Y":
            items.append((r["display_order"], r["item_name"]))

    return items