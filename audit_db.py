import os

from db import get_conn, now_iso, ph, DB_TYPE

from master_loader import load_audit_locations_rows, load_audit_surfaces_rows

DB_TYPE = os.getenv("DB_TYPE", "sqlite").lower()

def ph():
    return "%s" if DB_TYPE == "postgres" else "?"

def get_completed_audits_grouped_by_tower(date_from=None, date_to=None):
    """
    Returns completed audits grouped into:
    {
        "A": [audit_dict, ...],
        "B": [audit_dict, ...],
        "C": [audit_dict, ...],
    }

    Each audit_dict contains:
    - header
    - standard_results
    - additional_results
    """
    conn = get_conn()
    cur = conn.cursor()

    query = f"""
        SELECT *
        FROM audits
        WHERE status = 'COMPLETED'
    """

    params = []

    if date_from:
        query += f" AND audit_date >= {ph()}"
        params.append(str(date_from))

    if date_to:
        query += f" AND audit_date <= {ph()}"
        params.append(str(date_to))

    query += """
        ORDER BY date(audit_date) ASC, audit_id ASC
    """

    cur.execute(query, params)
    audit_rows = [dict(r) for r in cur.fetchall()]

    grouped = {"A": [], "B": [], "C": []}

    for audit in audit_rows:
        tower = (audit.get("tower") or "").upper()
        if tower not in grouped:
            continue

        query = f"""
            SELECT *
            FROM audit_results
            WHERE audit_id = {ph()}
            ORDER BY is_additional ASC, display_order ASC, result_id ASC
        """
        cur.execute(query, (audit["audit_id"],))
        result_rows = [dict(r) for r in cur.fetchall()]

        standard_results = [r for r in result_rows if r["is_additional"] == "N"]
        additional_results = [r for r in result_rows if r["is_additional"] == "Y"]

        grouped[tower].append({
            "audit": audit,
            "standard_results": standard_results,
            "additional_results": additional_results,
        })

    conn.close()
    return grouped


def get_surface_template_for_tower(tower: str):
    """
    Returns the master surface structure for a tower's template,
    including area grouping and display order.
    """
    template_group = get_template_for_tower(tower)
    if not template_group:
        return []
    return get_surfaces_for_template(template_group)

# ---------------------------
# Location filtering (by user group)
# ---------------------------
def get_visible_locations_for_user(user: dict):
    team = user.get("team_code", "")
    key = f"for_{team}"

    rows = load_audit_locations_rows()

    visible = []
    for r in rows:
        if r["active"] != "Y":
            continue
        if key in r and r[key] == "Y":
            visible.append(r)

    visible.sort(key=lambda x: x["display_order"])
    return visible

def get_visible_zones_for_user(user: dict):
    """
    Returns unique visible zones for the user, sorted.
    Blank zones are ignored.
    """
    locations = get_visible_locations_for_user(user)

    zones = []
    seen = set()

    for loc in locations:
        zone = (loc.get("zone") or "").strip()
        if not zone:
            continue
        if zone not in seen:
            seen.add(zone)
            zones.append(zone)

    return zones


def get_visible_locations_for_user_and_zone(user: dict, selected_zone: str):
    """
    Returns visible locations for the user filtered by selected zone.
    """
    selected_zone = (selected_zone or "").strip()
    locations = get_visible_locations_for_user(user)

    filtered = [
        loc for loc in locations
        if (loc.get("zone") or "").strip() == selected_zone
    ]

    filtered.sort(key=lambda x: x["display_order"])
    return filtered

# ---------------------------
# Template detection
# ---------------------------
def get_template_for_tower(tower: str):
    tower = (tower or "").upper()

    if tower in ("A", "C"):
        return "AC"
    elif tower == "B":
        return "B"
    return None


# ---------------------------
# Load surfaces for template
# ---------------------------
def get_surfaces_for_template(template_group: str):
    surfaces = load_audit_surfaces_rows()
    result = []

    for s in surfaces:
        if s["active"] != "Y":
            continue

        if template_group == "AC" and s["in_ac"] == "Y":
            result.append({
                "surface_id": s["surface_id"],
                "surface_name": s["surface_name"],
                "area_group": s["ac_area_group"],
                "display_order": int(s["ac_display_order"] or 0),
            })

        elif template_group == "B" and s["in_b"] == "Y":
            result.append({
                "surface_id": s["surface_id"],
                "surface_name": s["surface_name"],
                "area_group": s["b_area_group"],
                "display_order": int(s["b_display_order"] or 0),
            })

    result.sort(key=lambda x: x["display_order"])
    return result


# ---------------------------
# Audit list
# ---------------------------
def list_audits_for_user(user: dict):
    conn = get_conn()
    cur = conn.cursor()

    role = user.get("role", "")
    username = user.get("username", "")

    if role == "ADMIN":
        cur.execute("""
            SELECT audit_id, audit_date, auditor_name, staff_name, location_name,
                   tower, zone, template_group, remarks, status, created_by, created_at
            FROM audits
            ORDER BY date(audit_date) DESC, audit_id DESC
        """)
    else:
        query = f"""
            SELECT audit_id, audit_date, auditor_name, staff_name, location_name,
                   tower, zone, template_group, remarks, status, created_by, created_at
            FROM audits
            WHERE created_by = {ph()}
            ORDER BY date(audit_date) DESC, audit_id DESC
        """
        cur.execute(query, (username,))

    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


# ---------------------------
# Create audit header only
# ---------------------------
def create_audit_header(
    audit_date: str,
    auditor_name: str,
    staff_name: str,
    location: dict,
    template_group: str,
    remarks: str,
    created_by: str,
):
    conn = get_conn()
    cur = conn.cursor()

    created_at = now_iso()

    location_name = location.get("location_name")
    tower = location.get("tower")
    zone = location.get("zone")

    try:
        if DB_TYPE == "postgres":
            query = f"""
                INSERT INTO audits (
                    audit_date,
                    auditor_name,
                    staff_name,
                    location_name,
                    tower,
                    zone,
                    template_group,
                    remarks,
                    status,
                    created_by,
                    created_at
                )
                VALUES ({ph()}, {ph()}, {ph()}, {ph()}, {ph()}, {ph()}, {ph()}, {ph()}, 'DRAFT', {ph()}, {ph()})
                RETURNING audit_id
            """
        else:
            query = f"""
                INSERT INTO audits (
                    audit_date,
                    auditor_name,
                    staff_name,
                    location_name,
                    tower,
                    zone,
                    template_group,
                    remarks,
                    status,
                    created_by,
                    created_at
                )
                VALUES ({ph()}, {ph()}, {ph()}, {ph()}, {ph()}, {ph()}, {ph()}, {ph()}, 'DRAFT', {ph()}, {ph()})
            """

        cur.execute(query, (
            audit_date,
            auditor_name,
            staff_name,
            location_name,
            tower,
            zone,
            template_group,
            remarks,
            created_by,
            created_at,
        ))

        if DB_TYPE == "postgres":
            row = cur.fetchone()
            audit_id = row["audit_id"]
        else:
            audit_id = cur.lastrowid

        conn.commit()
        return int(audit_id)

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()
# ---------------------------
# Read audit header
# ---------------------------
def get_audit_header(audit_id: int):
    conn = get_conn()
    cur = conn.cursor()

    query = f"""
        SELECT *
        FROM audits
        WHERE audit_id = {ph()}
    """
    cur.execute(query, (audit_id,))
    row = cur.fetchone()
    conn.close()

    return dict(row) if row else None


# ---------------------------
# Read audit results
# ---------------------------
def get_audit_results(audit_id: int):
    conn = get_conn()
    cur = conn.cursor()

    query = f"""
        SELECT *
        FROM audit_results
        WHERE audit_id = {ph()}
        ORDER BY is_additional ASC, display_order ASC, result_id ASC
    """
    cur.execute(query, (audit_id,))
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


# ---------------------------
# Save audit detail
# ---------------------------
def save_audit_detail(
    audit_id: int,
    audit_date: str,
    auditor_name: str,
    staff_name: str,
    remarks: str,
    standard_results: list,
    additional_results: list,
    final_status: str = "DRAFT",
):
    if final_status not in ("DRAFT", "COMPLETED"):
        raise ValueError("Invalid audit status.")

    conn = get_conn()
    cur = conn.cursor()

    # update header
    query = f"""
        UPDATE audits
        SET audit_date = {ph()},
            auditor_name = {ph()},
            staff_name = {ph()},
            remarks = {ph()},
            status = {ph()}
        WHERE audit_id = {ph()}
    """
    cur.execute(query, (
        audit_date,
        auditor_name,
        staff_name,
        remarks,
        final_status,
        audit_id
    ))

    # replace detail rows
    query = f"""DELETE FROM audit_results WHERE audit_id = {ph()}"""
    cur.execute(query, (audit_id,))

    for r in standard_results:
        query = f"""
            INSERT INTO audit_results (
                audit_id,
                surface_id,
                surface_name,
                result,
                is_additional,
                area_group,
                display_order
            )
            VALUES ({ph()}, {ph()}, {ph()}, {ph()}, 'N', {ph()}, {ph()})
        """
        cur.execute(query, (
            audit_id,
            r["surface_id"],
            r["surface_name"],
            r["result"],
            r["area_group"],
            r["display_order"],
        ))

    for r in additional_results:
        surface_name = (r.get("surface_name") or "").strip()
        result = r.get("result") or "NA"

        if not surface_name:
            continue

        query = f"""
            INSERT INTO audit_results (
                audit_id,
                surface_id,
                surface_name,
                result,
                is_additional
            )
            VALUES ({ph()}, NULL, {ph()}, {ph()}, 'Y')
        """ 
        cur.execute(query, (
            audit_id,
            surface_name,
            result,
        ))

    conn.commit()
    conn.close()

