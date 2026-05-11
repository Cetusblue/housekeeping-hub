import streamlit as st
from datetime import date

from audit_db import (
    get_visible_locations_for_user,
    get_visible_zones_for_user,
    get_visible_locations_for_user_and_zone,
    get_template_for_tower,
    get_surfaces_for_template,
    list_audits_for_user,
    create_audit_header,
    get_audit_header,
    get_audit_results,
    save_audit_detail,
    get_completed_audits_grouped_by_tower,
    get_surface_template_for_tower,
)

from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

from master_loader import load_destinations_rows
from admin_tools import reset_orders_only, reset_orders_and_movements

from packing_db import get_packing_list_data, save_packing_board_issued

from packing_db import get_packing_list_data
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO
from datetime import datetime

from report_db import get_half_year_report_data

from report_db import get_half_year_report_data
from master_loader import get_item_master_lookup

from stock_db import (
    create_stock_in,
    get_inventory_rows,
    create_adhoc_issue_batch,
    get_stock_card_rows,
)
from io import BytesIO
from openpyxl import Workbook
from db import init_db, seed_minimal_data
from auth import authenticate
from run_dates import compute_run_date
from orders_db import (
    list_orders_for_store,
    list_orders_for_team,
    get_order_by_unique,
    get_order,
    create_order,
    get_order_lines,
    update_order_lines,
)


st.set_page_config(page_title="Housekeeping Hub", layout="centered")


# ---------------------------
# App bootstrap
# ---------------------------
def ensure_bootstrap():
    init_db()
    seed_minimal_data()


# ---------------------------
# Session helpers
# ---------------------------
def logout():
    st.session_state.pop("user", None)
    st.session_state["page"] = "login"


def require_login():
    if "user" not in st.session_state or not st.session_state["user"]:
        st.session_state["page"] = "login"
        st.stop()


# ---------------------------
# UI helpers
# ---------------------------
def format_status(status: str) -> str:
    if status == "PENDING":
        return "🟡 PENDING"
    if status == "PARTIALLY_ISSUED":
        return "🟠 PARTIALLY ISSUED"
    if status == "ISSUED":
        return "🟢 ISSUED"
    return status


# ---------------------------
# Login page
# ---------------------------
def page_login():
    st.title("Ops Hub")
    st.markdown(
        "<p style='font-size:14px; color:gray;'>"
        "Housekeeping & Linen Operations System"
        "</p>",
        unsafe_allow_html=True
    )
    st.markdown(
        "<p style='font-size:18px; font-weight:600;'>Login</p>",
        unsafe_allow_html=True
    )

    username = st.text_input("Username", placeholder="e.g. HenryC")
    pin = st.text_input("PIN / Password", type="password", placeholder="Enter password")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Log in", use_container_width=True):
            user = authenticate(username.strip(), pin.strip())
            if user:
                st.session_state["user"] = user
                st.session_state["page"] = "home"
                st.rerun()
            else:
                st.error("Invalid username or password.")
    with col2:
        if st.button("Clear", use_container_width=True):
            st.rerun()
    
    st.divider()
    st.markdown("""
    ### Announcements

    13/5/2026 Update
    - Changed Stock Monthly Report format, '0's are now blanks 
    - Added grouped ordering layout  
    - Added Inventory categories  
    - Added new Tuesday items
    - Revised app header
    """)


# ---------------------------
# Home page (role-based)
# ---------------------------
def page_home():
    require_login()
    user = st.session_state["user"]
    role = user["role"]

    st.title("Home")
    st.write(f"Logged in as **{user['username']}** ({role})")

    # TEAM
    if role == "TEAM":
        if st.button("Orders", use_container_width=True):
            st.session_state["page"] = "orders"
            st.rerun()

        if st.button("Glo Gel Audit", use_container_width=True):
            st.session_state["page"] = "glo_gel_audits"
            st.rerun()

        st.button("Logout", use_container_width=True, on_click=logout)
        return

    # STORE / ADMIN
    if role in ("STORE", "ADMIN"):
        if st.button("Orders", use_container_width=True):
            st.session_state["page"] = "orders"
            st.rerun()

        if st.button("Issue Stock", use_container_width=True):
            st.session_state["page"] = "issue_stock"
            st.rerun()

        if st.button("Stock In", use_container_width=True):
            st.session_state["page"] = "stock_in"
            st.rerun()

        if st.button("Inventory", use_container_width=True):
            st.session_state["page"] = "inventory"
            st.rerun()

        if st.button("Monthly Report", use_container_width=True):
            st.session_state["page"] = "monthly_report"
            st.rerun()

        if st.button("Stock Card Export", use_container_width=True):
            st.session_state["page"] = "stock_card"
            st.rerun()

        if role == "ADMIN":
            if st.button("System Tools", use_container_width=True):
                st.session_state["page"] = "system_tools"
                st.rerun()

            if st.button("Glo Gel Audit", use_container_width=True):
                st.session_state["page"] = "glo_gel_audits"
                st.rerun()

        st.button("Logout", use_container_width=True, on_click=logout)
        return

    # BOSS
    if role == "BOSS":
        if st.button("Monthly Report", use_container_width=True):
            st.session_state["page"] = "monthly_report"
            st.rerun()

        if st.button("Stock Card Export", use_container_width=True):
            st.session_state["page"] = "stock_card"
            st.rerun()

        st.button("Logout", use_container_width=True, on_click=logout)
        return


# ---------------------------
# Orders page router
# ---------------------------
def page_orders():
    require_login()
    user = st.session_state["user"]
    role = user["role"]

    st.title("Orders")

    if role == "STORE":
        page_orders_store(user)
        return

    if role == "TEAM":
        page_orders_team(user)
        return

    if role == "ADMIN":
        page_orders_admin(user)
        return

    st.error("Access denied.")


# ---------------------------
# Orders page: STORE
# ---------------------------
def page_orders_store(user):
    st.caption("Storeman view: pending orders first.")

    status_filter = st.selectbox(
        "Filter",
        ["PENDING", "PARTIALLY_ISSUED", "ISSUED", "ALL"],
        index=0,
    )
    status = None if status_filter == "ALL" else status_filter

    orders = list_orders_for_store(status=status)

    if not orders:
        st.info("No orders found.")
    else:
        for o in orders:
            header = f"{format_status(o['status'])} | {o['team_code']} | {o['template_day']} | Issue {o['run_date']} | by {o['created_by']}"

            with st.expander(header, expanded=False):
                st.write(f"Submitted: **{o['created_at']}**")
                if o.get("issued_at"):
                    st.write(f"Last issued: **{o['issued_at']}** by **{o.get('issued_by', '-') }**")

                if st.button(
                    "Open Order",
                    key=f"store_open_{o['order_id']}",
                    use_container_width=True,
                ):
                    st.session_state["active_order_id"] = o["order_id"]
                    st.session_state["page"] = "order_detail"
                    st.rerun()

    st.divider()
    st.subheader("Packing List Export")

    c1, c2 = st.columns(2)

    with c1:
        if st.button("Generate Tuesday Packing List", use_container_width=True):
            try:
                st.session_state["packing_output"] = generate_packing_list_excel(mode="TUE_COMBINED")
                st.rerun()
            except Exception as e:
                st.error(str(e))

    with c2:
        if st.button("Generate Friday Packing List", use_container_width=True):
            try:
                st.session_state["packing_output"] = generate_packing_list_excel(mode="FRI_COMBINED")
                st.rerun()
            except Exception as e:
                st.error(str(e))

    if "packing_output" in st.session_state:
        st.download_button(
            label="Download Packing List",
            data=st.session_state["packing_output"],
            file_name="packing_list.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Back", use_container_width=True):
            st.session_state["page"] = "home"
            st.rerun()
    with c2:
        st.button("Logout", use_container_width=True, on_click=logout)


# ---------------------------
# Orders page: TEAM
# ---------------------------
def page_orders_team(user):
    st.caption("Team view: one order per issue date. Only creator can edit pending.")

    orders = list_orders_for_team(user["team_code"])

    if not orders:
        st.info("No orders yet.")
    else:
        for o in orders:
            header = f"{format_status(o['status'])} | {o['template_day']} | Issue {o['run_date']} | by {o['created_by']}"

            with st.expander(header, expanded=False):
                st.write(f"Group: **{o['team_code']}**")
                st.write(f"Created: **{o['created_at']}**")

                if o.get("issued_at"):
                    st.write(f"Last issued: **{o['issued_at']}** by **{o.get('issued_by', '-') }**")

                is_creator = (o["created_by"] == user["username"])
                can_edit = (o["status"] == "PENDING") and is_creator

                c1, c2 = st.columns([1, 2])
                with c1:
                    if st.button(
                        "Edit Order" if can_edit else "View Order",
                        key=f"team_open_{o['order_id']}",
                        use_container_width=True,
                    ):
                        st.session_state["active_order_id"] = o["order_id"]
                        st.session_state["page"] = "order_detail"
                        st.rerun()
                with c2:
                    if not is_creator and o["status"] == "PENDING":
                        st.caption("Edit locked: only the creator can edit this pending order.")
                    elif o["status"] in ("PARTIALLY_ISSUED", "ISSUED"):
                        st.caption("This order is no longer editable by team users.")


    st.divider()
    st.subheader("Create / Open Order")

    # Row 1
    c1, c2 = st.columns(2)

    with c1:
        if st.button("Tuesday Order", use_container_width=True):
            _open_or_create_order("TUE")

    with c2:
        if st.button("Friday Order", use_container_width=True):
            _open_or_create_order("FRI")

    # Row 2
    c3, c4 = st.columns(2)

    with c3:
        if user["team_code"] == "AA1-3":
            if st.button("Annex - Tuesday", use_container_width=True):
                _open_or_create_order("ANNEX_TUE")

    with c4:
        if user["team_code"] == "AA1-3":
            if st.button("Annex - Friday", use_container_width=True):
                _open_or_create_order("ANNEX_FRI")

    # OT stays separate (optional row 3)
    if user["team_code"] == "B1-4":
        st.divider()
        if st.button("OT Order", use_container_width=True):
            _open_or_create_order("OT")

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Back", use_container_width=True):
            st.session_state["page"] = "home"
            st.rerun()
    with c2:
        st.button("Logout", use_container_width=True, on_click=logout)


# ---------------------------
# Orders page: ADMIN
# ---------------------------
def page_orders_admin(user):
    st.caption("Admin view: browse all orders.")

    orders = list_orders_for_store(status=None)

    if not orders:
        st.info("No orders found.")
    else:
        for o in orders:
            header = f"{format_status(o['status'])} | {o['team_code']} | {o['template_day']} | Issue {o['run_date']}"

            with st.expander(header, expanded=False):
                if st.button(
                    "Open Order",
                    key=f"admin_open_{o['order_id']}",
                    use_container_width=True,
                ):
                    st.session_state["active_order_id"] = o["order_id"]
                    st.session_state["page"] = "order_detail"
                    st.rerun()

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Back", use_container_width=True):
            st.session_state["page"] = "home"
            st.rerun()
    with c2:
        st.button("Logout", use_container_width=True, on_click=logout)


# ---------------------------
# Create / open order helper
# ---------------------------
def _open_or_create_order(template_day: str):
    user = st.session_state["user"]
    team_code = user["team_code"]

    today = date.today()

    if template_day == "OT":
        # OT is special and can be made any day; for now use today's date as issue date
        run_date = today.isoformat()
    else:
        run_date = compute_run_date(today, template_day).isoformat()

    existing = get_order_by_unique(team_code, template_day, run_date)

    if existing:
        st.session_state["active_order_id"] = existing["order_id"]
        st.session_state["page"] = "order_detail"
        st.rerun()
    else:
        new_id = create_order(
            team_code=team_code,
            template_day=template_day,
            run_date=run_date,
            created_by=user["username"],
        )
        st.session_state["active_order_id"] = new_id
        st.session_state["page"] = "order_detail"
        st.rerun()


# ---------------------------
# Order detail page
# ---------------------------
def page_order_detail():
    require_login()
    user = st.session_state["user"]

    if user["role"] == "BOSS":
        st.error("Access denied.")
        return

    order_id = st.session_state.get("active_order_id")
    if not order_id:
        st.session_state["page"] = "orders"
        st.rerun()

    o = get_order(int(order_id))
    if not o:
        st.error("Order not found.")
        st.session_state["page"] = "orders"
        st.rerun()

    lines = get_order_lines(int(order_id), template_day=o["template_day"])

    if user["role"] in ("STORE", "ADMIN"):
        lines = [r for r in lines if int(r["qty_requested"] or 0) > 0]

    st.title("Order Detail")
    st.caption(
        f"{o['team_code']} | {o['template_day']} | Issue {o['run_date']} | "
        f"Status: {format_status(o['status'])} | Created by {o['created_by']}"
    )

    is_team_creator = (user["role"] == "TEAM") and (o["created_by"] == user["username"])
    can_edit_request = (user["role"] == "TEAM") and is_team_creator and (o["status"] == "PENDING")

    # STORE / ADMIN can increase issued quantities
    can_edit_issued = user["role"] in ("STORE", "ADMIN")

    item_lookup = get_item_master_lookup()

    data = []
    for r in lines:
        item_name = r["item_name"]
        unit = ""
        details = ""
        category = "Others"

        if item_name in item_lookup:
            unit = item_lookup[item_name]["unit"]
            details = item_lookup[item_name]["note"]
            category = item_lookup[item_name].get("category", "Others") or "Others"

        data.append({
            "line_id": r["line_id"],
            "Category": category,
            "Item": item_name,
            "UOM": unit,
            "Details": details,
            "Request": int(r["qty_requested"] or 0),
            "Issued": int(r["qty_issued"] or 0),
        })

    data_key = f"order_detail_data_{order_id}"

    data.sort(key=lambda x: (x.get("Category", "Others"), x.get("Item", "")))

    if data_key not in st.session_state:
        st.session_state[data_key] = data

    original_by_line_id = {int(r["line_id"]): r for r in lines}

    st.subheader("Items")

    disabled_cols = ["No", "UOM", "Details"]
    if not can_edit_request:
        disabled_cols.append("Request")
    if not can_edit_issued:
        disabled_cols.append("Issued")

    if user["role"] in ("STORE", "ADMIN"):
        if st.button("Fulfill All", use_container_width=True, key=f"fulfill_all_{order_id}"):
            fulfilled_rows = []
            for row in st.session_state[data_key]:
                row = dict(row)
                row["Issued"] = int(row["Request"] or 0)
                fulfilled_rows.append(row)

            st.session_state[data_key] = fulfilled_rows
            st.rerun()

    # Group rows by category
    grouped = {}
    for row in st.session_state[data_key]:
        category = row.get("Category", "Others") or "Others"
        grouped.setdefault(category, []).append(row)

    edited_all = []

    for category, rows in grouped.items():
        st.markdown(f"### {category}")

        edited_part = st.data_editor(
            rows,
            hide_index=True,
            disabled=disabled_cols + ["line_id", "Category"],
            use_container_width=True,
            column_config={
                "line_id": None,
                "Category": None,
                "Item": st.column_config.TextColumn("Item", disabled=True),
                "UOM": st.column_config.TextColumn("UOM", disabled=True),
                "Details": st.column_config.TextColumn("Details", disabled=True),
                "Request": st.column_config.NumberColumn("Request", min_value=0, step=1),
                "Issued": st.column_config.NumberColumn("Issued", min_value=0, step=1),
            },
            key=f"editor_{order_id}_{category}",
        )

        edited_all.extend(edited_part)

    st.session_state[data_key] = edited_all
    edited = edited_all

    show_save = can_edit_request or can_edit_issued

    if show_save:
        save_label = "Save Issued" if user["role"] in ("STORE", "ADMIN") else "Save Request"

        if st.button(save_label, key=f"save_{order_id}_{user['role']}", use_container_width=True):
            rows_to_save = []

            for row in edited:
                lid = int(row.get("line_id") or 0)
                orig = original_by_line_id.get(lid)

                if not orig:
                    continue

                req = int(row["Request"])
                iss = int(row["Issued"])

                if not can_edit_request:
                    req = int(orig["qty_requested"] or 0)

                if not can_edit_issued:
                    iss = int(orig["qty_issued"] or 0)

                # Team users cannot issue. Store/Admin can only increase issued quantity.
                prev_issued = int(orig["qty_issued"] or 0)
                if can_edit_issued and iss < prev_issued:
                    iss = prev_issued

                if iss > req:
                    iss = req

                rows_to_save.append({
                    "line_id": lid,
                    "qty_requested": req,
                    "qty_issued": iss,
                })

            try:
                update_order_lines(
                    int(order_id),
                    rows_to_save,
                    updated_by=user["username"]
                )
                st.success(
                    "Issued quantities saved."
                    if user["role"] in ("STORE", "ADMIN")
                    else "Request saved."
                )

                if data_key in st.session_state:
                    del st.session_state[data_key]

                st.rerun()
            except ValueError as e:
                st.error(str(e))
    else:
        st.caption("No editable fields for your role on this order.")

    st.divider()

    if o["status"] == "ISSUED":
        if user["role"] in ("STORE", "ADMIN"):
            st.info("This order is fully issued. Further increases may still be saved if needed.")
        else:
            st.info("This order is issued and locked.")
    elif o["status"] == "PARTIALLY_ISSUED":
        if user["role"] in ("STORE", "ADMIN"):
            st.info("This order is partially issued. Continue issuing as needed.")
        else:
            st.info("This order is partially issued and locked for team users.")
    else:
        if can_edit_request:
            st.success("You can edit Request quantities.")
        elif user["role"] == "TEAM":
            st.warning("Read-only: only the creator can edit this pending order.")
        if can_edit_issued:
            st.info("Store/Admin can edit Issued quantities.")

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Back to Orders", use_container_width=True):
            st.session_state["page"] = "orders"
            st.rerun()
    with c2:
        st.button("Logout", use_container_width=True, on_click=logout)


def page_issue_stock():
    require_login()
    user = st.session_state["user"]

    if user["role"] not in ("STORE", "ADMIN"):
        st.error("Access denied.")
        return

    st.title("Issue Stock")

    if "flash_message" in st.session_state:
        st.success(st.session_state["flash_message"])
        del st.session_state["flash_message"]

    destination_rows = load_destinations_rows()
    active_destinations = [
        r["destination_name"]
        for r in destination_rows
        if r["active"] == "Y"
    ]

    issued_to = st.selectbox("Issue To", active_destinations)
    final_destination = issued_to

    rows = get_inventory_rows()

    if not rows:
        st.info("No items found.")
        return

    data = []
    for r in rows:
        data.append({
            "Item": r["item_name"],
            "Current": int(r["current_stock"]),
            "Issue Qty": "",
        })

    edited = st.data_editor(
        data,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Item": st.column_config.TextColumn("Item", disabled=True),
            "Current": st.column_config.NumberColumn("Current", disabled=True),
            "Issue Qty": st.column_config.TextColumn("Issue Qty"),
        },
        key="issue_stock_editor",
    )

    if st.button("Issue", use_container_width=True):
        if not final_destination:
            st.warning("Please select or specify a destination.")
            return

    if st.button("Back", use_container_width=True):
        st.session_state["page"] = "home"
        st.rerun()

    issue_rows = []
    count = 0

    for row in edited:
        raw_qty = str(row["Issue Qty"]).strip()

        if raw_qty == "":
            continue

        try:
            qty = int(raw_qty.replace(",", ""))
        except Exception:
            st.error(f"Invalid Issue Qty for item: {row['Item']}")
            return

        if qty < 0:
            st.error(f"Negative Issue Qty is not allowed for item: {row['Item']}")
            return

        if qty > 0:
            issue_rows.append({
                "item_name": row["Item"],
                "qty": qty,
            })
            count += 1

    if count == 0:
        st.warning("No quantities entered.")
        return

    try:
        create_adhoc_issue_batch(
            issue_rows=issue_rows,
            issued_to=final_destination,
            created_by=user["username"],
        )
        st.session_state["flash_message"] = f"Stock issued for {count} item(s) to {final_destination}."
        st.rerun()
    except ValueError as e:
        st.error(str(e))

    st.divider()


def page_stock_in():
    require_login()
    user = st.session_state["user"]

    if user["role"] not in ("STORE", "ADMIN"):
        st.error("Access denied.")
        return

    st.title("Stock In")

    if "flash_message" in st.session_state:
        st.success(st.session_state["flash_message"])
        del st.session_state["flash_message"]

    rows = get_inventory_rows()

    if not rows:
        st.info("No items found.")
        return

    data = []
    for r in rows:
        data.append({
            "Category": r.get("category") or r.get("Category") or "Others",
            "Item": r["item_name"],
            "Current": int(r["current_stock"]),
            "Stock In": ""
        })

    grouped = {}
    for row in data:
        category = row.get("Category", "Others") or "Others"
        grouped.setdefault(category, []).append(row)

    edited_all = []

    for category, rows in grouped.items():
        st.markdown(f"### {category}")

        edited_part = st.data_editor(
            rows,
            hide_index=True,
            use_container_width=True,
            column_config={
                "Category": None,
                "Item": st.column_config.TextColumn("Item", disabled=True),
                "Current": st.column_config.NumberColumn("Current", disabled=True),
                "Stock In": st.column_config.TextColumn("Stock In"),
            },
            key=f"stock_in_editor_{category}",
        )

        edited_all.extend(edited_part)

    edited = edited_all

    if st.button("Stock In", use_container_width=True):
        count = 0

        for row in edited:
            raw_qty = str(row["Stock In"]).strip()

            if raw_qty == "":
                continue

            try:
                qty = int(raw_qty.replace(",", ""))
            except Exception:
                st.error(f"Invalid Stock In quantity for item: {row['Item']}")
                return

            if qty < 0:
                st.error(f"Negative Stock In quantity is not allowed for item: {row['Item']}")
                return

            if qty == 0:
                continue

            create_stock_in(
                item_name=row["Item"],
                qty=qty,
                created_by=user["username"]
            )
            count += 1

        if count == 0:
            st.warning("No quantities entered.")
        else:
            st.session_state["flash_message"] = f"Stock-in recorded for {count} item(s)."
            st.rerun()

    st.divider()

    if st.button("Back", use_container_width=True):
        st.session_state["page"] = "home"
        st.rerun()

def page_inventory():
    require_login()
    user = st.session_state["user"]

    if user["role"] not in ("STORE", "ADMIN"):
        st.error("Access denied.")
        return

    st.title("Inventory")

    rows = get_inventory_rows()

    if not rows:
        st.info("No inventory data.")
        return

    data = []
    for r in rows:
        data.append({
            "Category": r.get("category", "Others"),
            "Item": r["item_name"],
            "Current Stock": int(r["current_stock"])
        })

    grouped = {}

    for row in data:
        category = row.get("Category", "Others") or "Others"
        grouped.setdefault(category, []).append(row)

    for category, rows in grouped.items():
        st.markdown(f"### {category}")

        st.dataframe(
            rows,
            hide_index=True,
            use_container_width=True,
            column_config={
                "Category": None,
            },
        )

    st.divider()

    if st.button("Back", use_container_width=True):
        st.session_state["page"] = "home"
        st.rerun()


def page_monthly_report():
    require_login()
    user = st.session_state["user"]

    if user["role"] not in ("STORE", "ADMIN", "BOSS"):
        st.error("Access denied.")
        return

    st.title("Monthly Report")
    st.caption("Export fixed half-year stock records template.")

    year = date.today().year

    c1, c2 = st.columns(2)

    with c1:
        if st.button("H1 (Jan - Jun)", use_container_width=True):
            _generate_half_year_report_excel(year, "H1")

    with c2:
        if st.button("H2 (Jul - Dec)", use_container_width=True):
            _generate_half_year_report_excel(year, "H2")

    st.divider()

    if st.button("Back", use_container_width=True):
        st.session_state["page"] = "home"
        st.rerun()

def blank_if_zero(value):
    if value in (0, 0.0, "0", "0.0"):
        return ""
    return value

def _generate_half_year_report_excel(year: int, period_code: str):
    report_data = get_half_year_report_data(year, period_code)

    wb = Workbook()

    # remove default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)

    if period_code == "H1":
        months = [1, 2, 3, 4, 5, 6]
    else:
        months = [7, 8, 9, 10, 11, 12]

    month_headers = [date(year, m, 1).strftime("%b %y") for m in months]

    sheet_order = ["Tower ABC", "Tower A", "Tower B", "Tower C", "ANX Blk", "Others"]

    for sheet_name in sheet_order:
        ws = wb.create_sheet(title=sheet_name)

        # Top title
        ws["A1"] = "Stock Records"
        ws["A2"] = sheet_name
        ws["A3"] = f"{period_code} {year}"

        # Static headers A:F
        ws["A5"] = "SN"
        ws["B5"] = "Category"
        ws["C5"] = "Description"
        ws["D5"] = "UOM"
        ws["E5"] = ""
        ws["F5"] = "Unit Rate"

        # Dynamic month headers G:L
        start_col = 7  # G
        for i, header in enumerate(month_headers):
            ws.cell(row=5, column=start_col + i, value=header)

        excel_row = 6

        for row in report_data[sheet_name]:
            # Static columns A:F
            ws.cell(row=excel_row, column=1, value=row["report_line_id"])     # A
            ws.cell(row=excel_row, column=2, value=row["for_column_b"])       # B
            ws.cell(row=excel_row, column=3, value=row["report_line_name"])   # C
            ws.cell(row=excel_row, column=4, value=row["report_uom"])         # D
            ws.cell(row=excel_row, column=5, value=row["for_column_e"])       # E
            ws.cell(row=excel_row, column=6, value=row["for_column_f"])       # F

            # Month columns G:L
            monthly_qty = row["monthly_qty"]
            for i, month in enumerate(months):
                ws.cell(row=excel_row, column=start_col + i, value=blank_if_zero(monthly_qty.get(month, 0)))

            excel_row += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    file_name = f"stock_records_{period_code}_{year}.xlsx"

    st.download_button(
        label=f"Download {period_code} Report Excel",
        data=output,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

def page_stock_card():
    require_login()
    user = st.session_state["user"]

    if user["role"] not in ("STORE", "ADMIN", "BOSS"):
        st.error("Access denied.")
        return

    st.title("Stock Card Export")

    rows = get_inventory_rows()

    if not rows:
        st.info("No inventory items found.")
        return

    st.caption("Select items to include in the stock card export.")

    date_from = st.date_input("Date From", value=None)
    date_to = st.date_input("Date To", value=None)

    selections = []
    for idx, r in enumerate(rows):
        col1, col2, col3 = st.columns([1, 6, 2])
        with col1:
            checked = st.checkbox("", key=f"stock_card_item_{idx}")
        with col2:
            st.write(r["item_name"])
        with col3:
            st.write(r["current_stock"])
        if checked:
            selections.append(r["item_name"])

    if st.button("Generate Stock Card Excel", use_container_width=True):
        if not selections:
            st.warning("Please select at least one item.")
            return

        wb = Workbook()
        # remove default sheet
        default_sheet = wb.active
        wb.remove(default_sheet)

        for item_name in selections:
            sheet_name = item_name[:31] if item_name else "Item"
            ws = wb.create_sheet(title=sheet_name)

            ws["A1"] = "Item"
            ws["B1"] = item_name

            ws["A3"] = "Date"
            ws["B3"] = "Stock In"
            ws["C3"] = "Stock Out"
            ws["D3"] = "Balance"
            ws["E3"] = "Issued To"

            stock_rows = get_stock_card_rows(
                item_name=item_name,
                date_from=str(date_from) if date_from else None,
                date_to=str(date_to) if date_to else None,
            )

            excel_row = 4
            for row in stock_rows:
                ws.cell(row=excel_row, column=1, value=row["Date"])
                ws.cell(row=excel_row, column=2, value=row["Stock In"])
                ws.cell(row=excel_row, column=3, value=row["Stock Out"])
                ws.cell(row=excel_row, column=4, value=row["Balance"])
                ws.cell(row=excel_row, column=5, value=row["Issued To"])
                excel_row += 1

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.download_button(
            label="Download Stock Card Excel",
            data=output,
            file_name="stock_card_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()

    if st.button("Back", use_container_width=True):
        st.session_state["page"] = "home"
        st.rerun()


def generate_packing_list_excel(mode="TUE_COMBINED"):
    data = get_packing_list_data(mode=mode)

    items = data["items"]
    teams = data["team_order"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Packing List"

    # -----------------------
    # PAGE SETUP (A3 Landscape)
    # -----------------------
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = False

    # -----------------------
    # HEADER
    # -----------------------
    ws["A1"] = "Packing List"
    ws["A2"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    ws["A1"].font = Font(size=14, bold=True)

    # -----------------------
    # TABLE HEADER
    # -----------------------
    headers = ["Item", "Stock"] + teams + ["Remarks"]

    # --------------------------
    # TITLE
    # --------------------------

    title_map = {
        "TUE_COMBINED": "Packing List - Tuesday",
        "FRI_COMBINED": "Packing List - Friday",
        "TUE": "Packing List - Tuesday",
        "FRI": "Packing List - Friday",
        "ANNEX_TUE": "Packing List - Annex Tuesday",
        "ANNEX_FRI": "Packing List - Annex Friday",
    }

    title = title_map.get(mode, "Packing List")

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2 + len(teams))
    cell = ws.cell(row=1, column=1, value=title)
    cell.font = Font(size=16, bold=True)
    cell.alignment = Alignment(horizontal="center")

    header_row = 4

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # -----------------------
    # DATA ROWS
    # -----------------------
    row_idx = header_row + 1

    for item in items:
        ws.cell(row=row_idx, column=1, value=item["item_name"])
        ws.cell(row=row_idx, column=2, value=item["stock"])

        col_offset = 3
        for i, team in enumerate(teams):
            team_data = item["teams"].get(team, {})
            qty = int(team_data.get("requested", 0) or 0)
            ws.cell(row=row_idx, column=col_offset + i, value=qty)

        # remarks column (blank for writing)
        ws.cell(row=row_idx, column=col_offset + len(teams), value="")

        row_idx += 1

    # -----------------------
    # COLUMN WIDTHS
    # -----------------------
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 10

    for i in range(len(teams)):
        col_letter = chr(ord("C") + i)
        ws.column_dimensions[col_letter].width = 10

    last_col_letter = chr(ord("C") + len(teams))
    ws.column_dimensions[last_col_letter].width = 20  # Remarks

    # -----------------------
    # FREEZE HEADER
    # -----------------------
    ws.freeze_panes = "A5"

    # -----------------------
    # EXPORT
    # -----------------------
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output

def page_packing_fulfillment():
    require_login()
    user = st.session_state["user"]

    if user["role"] not in ("STORE", "ADMIN"):
        st.error("Access denied.")
        return

    st.title("Packing / Fulfillment")
    st.caption("Outstanding items across selected order type.")

    mode_label = st.selectbox(
        "Packing Mode",
        ["Tuesday", "Friday", "Annex - Tuesday", "Annex - Friday", "OT"],
        key="packing_mode"
    )

    mode_map = {
        "Tuesday": "Tuesday",
        "Friday": "Friday",
        "Annex - Tuesday": "ANNEX_TUE",
        "Annex - Friday": "ANNEX_FRI",
        "OT": "OT",
    }

    mode = mode_map[mode_label]

    data = get_packing_list_data(mode)
    items = data["items"]
    teams = data["team_order"]

    if not items:
        st.info("No outstanding items.")

        st.divider()

        c1, c2 = st.columns(2)
        with c1:
            if st.button("Back", use_container_width=True):
                st.session_state["page"] = "home"
                st.rerun()
        with c2:
            if st.button("Orders", use_container_width=True):
                st.session_state["page"] = "orders"
                st.rerun()

        return

    # -----------------------
    # Build editable table rows
    # -----------------------
    table_rows = []
    for item in items:
        row = {
            "Item": item["item_name"],
            "Stock": item["stock"],
            "Status": "⚠ LOW" if item.get("low_stock") else "OK",
        }

        for team in teams:
            team_data = item["teams"].get(team, {})
            req = int(team_data.get("requested", 0) or 0)
            iss = int(team_data.get("issued", 0) or 0)

            if req > 0 and iss >= req:
                stat = "🟢"
            elif iss > 0:
                stat = "🟡"
            else:
                stat = "⚪"

            row[f"{team} Req"] = req
            row[f"{team} Iss"] = iss
            row[f"{team} Stat"] = stat

        table_rows.append(row)

    # -----------------------
    # Editable grid
    # -----------------------
    data_key = f"packing_data_{mode}"
    editor_key = f"packing_editor_{mode}"

    # initialize backing data only once per mode
    if data_key not in st.session_state:
        st.session_state[data_key] = table_rows

    # -----------------------
    # Fulfill Team buttons
    # -----------------------
    if teams:
        st.caption("Quick fill helpers")

        cols = st.columns(len(teams))
        for idx, team in enumerate(teams):
            with cols[idx]:
                if st.button(f"Fulfill {team}", use_container_width=True, key=f"fulfill_{team}_{mode}"):
                    current_rows = st.session_state.get(data_key, table_rows)

                    updated_rows = []
                    for row in current_rows:
                        req_col = f"{team} Req"
                        iss_col = f"{team} Iss"

                        req_val = int(row.get(req_col, 0) or 0)
                        iss_val = int(row.get(iss_col, 0) or 0)

                        # only fill untouched cells
                        if req_val > 0 and iss_val == 0:
                            row[iss_col] = req_val

                        updated_rows.append(row)

                    st.session_state[data_key] = updated_rows
                    st.rerun()

    # -----------------------
    # Editable grid
    # -----------------------
    editor_data = st.session_state.get(data_key, table_rows)

    disabled_cols = ["Item", "Stock", "Status"]
    for team in teams:
        disabled_cols.append(f"{team} Req")
        disabled_cols.append(f"{team} Stat")

    edited = st.data_editor(
        editor_data,
        hide_index=True,
        use_container_width=True,
        disabled=disabled_cols,
        key=editor_key,
    )

    # keep latest edited table rows separately from widget state
    st.session_state[data_key] = edited

    st.divider()

    if st.button("Issue / Save Issued", use_container_width=True):
        try:
            save_packing_board_issued(
                mode=mode,
                edited_rows=edited,
                updated_by=user["username"],
            )
            st.success("Packing board issued quantities saved.")
            if data_key in st.session_state:
                del st.session_state[data_key]
            st.rerun()
        except Exception as e:
            st.error(f"Could not save issued quantities: {e}")

    st.info("This screen is currently prefill + edit only. Save/Issue from this board comes next.")

    st.divider()

    if st.button("Generate Packing List (A3)", use_container_width=True):
        st.session_state["packing_output"] = generate_packing_list_excel()
        st.rerun()

    if "packing_output" in st.session_state:
        st.download_button(
            label="Download Packing List",
            data=st.session_state["packing_output"],
            file_name="packing_list.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Back", use_container_width=True, key="packing_back"):
            st.session_state["page"] = "home"
            st.rerun()
    with c2:
        if st.button("Orders", use_container_width=True, key="packing_orders"):
            st.session_state["page"] = "orders"
            st.rerun()


def page_system_tools():
    require_login()
    user = st.session_state["user"]

    if user["role"] != "ADMIN":
        st.error("Access denied.")
        return

    st.title("System Tools")
    st.caption("Danger zone. Use only during UAT / controlled resets.")

    confirm_text = st.text_input("Type RESET to enable reset actions")

    can_reset = confirm_text.strip().upper() == "RESET"

    st.divider()
    st.subheader("Reset Options")

    c1, c2 = st.columns(2)

    with c1:
        if st.button("Reset Orders Only", use_container_width=True, disabled=not can_reset):
            try:
                reset_orders_only()
                st.success("Orders and order lines reset successfully.")
            except Exception as e:
                st.error(f"Reset failed: {e}")

    with c2:
        if st.button("Reset Orders + Movements", use_container_width=True, disabled=not can_reset):
            try:
                reset_orders_and_movements()
                st.success("Orders, order lines, and stock movements reset successfully.")
            except Exception as e:
                st.error(f"Reset failed: {e}")

    st.divider()

    st.warning(
        "Full database reset is still best done manually by deleting "
        "`data/housekeeping_hub.db` while the app is stopped."
    )

    if st.button("Back", use_container_width=True):
        st.session_state["page"] = "home"
        st.rerun()

def page_glo_gel_audit_detail():
    require_login()
    user = st.session_state["user"]

    audit_id = st.session_state.get("selected_audit_id")
    if not audit_id:
        st.warning("No audit selected.")
        return

    audit = get_audit_header(int(audit_id))
    if not audit:
        st.error("Audit not found.")
        return

    st.title("Glo Gel Audit Detail")

    is_completed = audit["status"] == "COMPLETED"

    st.caption(
        f"{audit['location_name']} | {audit['audit_date']} | "
        f"Status: {audit['status']} | Created by {audit['created_by']}"
    )

    audit_date = st.date_input(
        "Date of Inspection",
        value=date.fromisoformat(audit["audit_date"]),
        disabled=is_completed,
        key=f"audit_date_{audit_id}"
    )
    auditor_name = st.text_input(
        "Auditor Name",
        value=audit["auditor_name"],
        disabled=is_completed,
        key=f"audit_auditor_{audit_id}"
    )
    staff_name = st.text_input(
        "Staff Name",
        value=audit["staff_name"],
        disabled=is_completed,
        key=f"audit_staff_{audit_id}"
    )

    remarks = st.text_area(
        "Remarks",
        value=audit["remarks"] or "",
        disabled=is_completed,
        key=f"audit_remarks_{audit_id}"
    )

    surfaces = get_surfaces_for_template(audit["template_group"])
    saved_results = get_audit_results(int(audit_id))

    saved_standard = {
        r["surface_id"]: r["result"]
        for r in saved_results
        if r["is_additional"] == "N"
    }
    saved_additional = [
        r for r in saved_results if r["is_additional"] == "Y"
    ]

    st.divider()
    st.subheader("High Touch Surfaces")

    grouped = {}
    for s in surfaces:
        area = s["area_group"] or "General"
        grouped.setdefault(area, []).append(s)

    if audit["template_group"] == "B":
        area_order = ["Patient", "Toilet", "General"]
    else:
        area_order = ["Clinic", "Toilet", "General"]

    ordered_areas = (
        [a for a in area_order if a in grouped] +
        [a for a in grouped.keys() if a not in area_order]
    )

    standard_results = []

    for area in ordered_areas:
        items = grouped[area]
        st.markdown(f"**{area}**")

        for s in items:
            default_result = saved_standard.get(s["surface_id"], "NA")
            choice = st.radio(
                s["surface_name"],
                options=["C", "NC", "NA"],
                horizontal=True,
                disabled=is_completed,
                index=["C", "NC", "NA"].index(default_result),
                key=f"audit_{audit_id}_{s['surface_id']}"
            )

            standard_results.append({
                "surface_id": s["surface_id"],
                "surface_name": s["surface_name"],
                "result": choice,
                "area_group": s["area_group"],
                "display_order": s["display_order"],
            })

    st.divider()
    st.subheader("Additional HTS")

    additional_results = []

    for i in range(3):
        existing_name = saved_additional[i]["surface_name"] if i < len(saved_additional) else ""
        existing_result = saved_additional[i]["result"] if i < len(saved_additional) else "NA"

        c1, c2 = st.columns([3, 2])

        with c1:
            name = st.text_input(
                f"Surface {i+1}",
                value=existing_name,
                disabled=is_completed,
                key=f"add_name_{audit_id}_{i}"
            )

        with c2:
            result = st.radio(
                "",
                options=["C", "NC", "NA"],
                horizontal=True,
                disabled=is_completed,
                index=["C", "NC", "NA"].index(existing_result),
                key=f"add_result_{audit_id}_{i}",
                label_visibility="collapsed"
            )

        additional_results.append({
            "surface_name": name,
            "result": result,
        })

    if not is_completed:
        c1, c2 = st.columns(2)

        with c1:
            if st.button("Save Draft", use_container_width=True, key=f"save_draft_{audit_id}"):
                save_audit_detail(
                    audit_id=int(audit_id),
                    audit_date=str(audit_date),
                    auditor_name=auditor_name,
                    staff_name=staff_name,
                    remarks=remarks,
                    standard_results=standard_results,
                    additional_results=additional_results,
                    final_status="DRAFT",
                )
                st.session_state["audit_success"] = True
                st.rerun()

        with c2:
            if st.button("Complete Audit", use_container_width=True, key=f"complete_audit_{audit_id}"):
                if not staff_name.strip():
                    st.warning("Please enter staff name.")
                else:
                    save_audit_detail(
                        audit_id=int(audit_id),
                        audit_date=str(audit_date),
                        auditor_name=auditor_name,
                        staff_name=staff_name,
                        remarks=remarks,
                        standard_results=standard_results,
                        additional_results=additional_results,
                        final_status="COMPLETED",
                    )
                    st.session_state["audit_success"] = True
                    st.session_state["page"] = "glo_gel_audits"
                    st.rerun()

    st.divider()
    if st.button("Back", use_container_width=True, key=f"audit_detail_back_{audit_id}"):
        st.session_state["page"] = "glo_gel_audits"
        st.rerun()
        
def page_glo_gel_audits():
    require_login()
    user = st.session_state["user"]

    st.title("Glo Gel Audits")

    # ---------------------------
    # Success message
    # ---------------------------
    if st.session_state.pop("audit_success", False):
        st.success("Audit saved successfully.")

    if user["role"] == "ADMIN":
        st.divider()
        st.subheader("Admin Report Export")

        if st.button("Prepare Compiled Audit Report", use_container_width=True, key="prepare_compiled_audit_report"):
            try:
                st.session_state["compiled_audit_report"] = generate_compiled_glo_gel_report_excel()
                st.rerun()
            except Exception as e:
                st.error(str(e))

        if "compiled_audit_report" in st.session_state:
            st.download_button(
                label="Download Compiled Audit Report",
                data=st.session_state["compiled_audit_report"],
                file_name="compiled_glo_gel_audit_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="download_compiled_audit_report"
            )

    # ---------------------------
    # Create New Audit
    # ---------------------------
    st.subheader("Create New Audit")

    zones = get_visible_zones_for_user(user)

    if not zones:
        st.warning("No zones available for your group.")
    else:
        selected_zone = st.selectbox(
            "Zone",
            zones,
            key="new_audit_zone"
        )

        locations = get_visible_locations_for_user_and_zone(user, selected_zone)

        if not locations:
            st.warning("No locations available for the selected zone.")
        else:
            location_names = [l["location_name"] for l in locations]
            selected_location_name = st.selectbox(
                "Area / Location",
                location_names,
                key="new_audit_location"
            )

            selected_location = next(
                l for l in locations if l["location_name"] == selected_location_name
            )

        auditor_name = st.text_input(
            "Auditor Name",
            value=user["username"],
            key="new_audit_auditor"
        )

        staff_name = st.text_input(
            "Staff Name",
            key="new_audit_staff"
        )

        audit_date = st.date_input(
            "Date of Inspection",
            key="new_audit_date"
        )

        remarks = st.text_area(
            "Remarks",
            key="new_audit_remarks"
        )

        if st.button("Create Audit Entry", use_container_width=True):
            if not staff_name.strip():
                st.warning("Please enter staff name.")
            else:
                template_group = get_template_for_tower(selected_location["tower"])

                audit_id = create_audit_header(
                    audit_date=str(audit_date),
                    auditor_name=auditor_name,
                    staff_name=staff_name,
                    location=selected_location,
                    template_group=template_group,
                    remarks=remarks,
                    created_by=user["username"],
                )

                if not audit_id:
                    st.error("Audit was not created. No audit_id returned.")
                    return

                st.session_state["selected_audit_id"] = audit_id
                st.session_state["page"] = "glo_gel_audit_detail"
                st.rerun()

    # ---------------------------
    # Audit History
    # ---------------------------
    st.divider()
    st.subheader("Audit History")

    audits = list_audits_for_user(user)

    if not audits:
        st.info("No audits yet.")
    else:
        # Filters
        col1, col2, col3 = st.columns(3)

        with col1:
            filter_status = st.selectbox(
                "Status",
                ["All", "DRAFT", "COMPLETED"],
                key="audit_filter_status"
            )

        with col2:
            filter_location = st.selectbox(
                "Location",
                ["All"] + sorted({a["location_name"] for a in audits}),
                key="audit_filter_location"
            )

        with col3:
            filter_staff = st.text_input(
                "Search Staff",
                key="audit_filter_staff"
            )

        # Apply filters
        filtered = []

        for a in audits:
            if filter_status != "All" and a["status"] != filter_status:
                continue

            if filter_location != "All" and a["location_name"] != filter_location:
                continue

            if filter_staff and filter_staff.lower() not in (a["staff_name"] or "").lower():
                continue

            filtered.append(a)

        # Display list
        for a in filtered:
            c1, c2 = st.columns([6, 1])

            status_icon = "🟡" if a["status"] == "DRAFT" else "🟢"

            with c1:
                status_icon = "🟡" if a["status"] == "DRAFT" else "🟢"

                zone = (a.get("zone") or "").strip()
                location_display = f"{zone} - {a['location_name']}" if zone else a["location_name"]

                st.write(
                    f"📅 {a['audit_date']} | "
                    f"📍 {location_display} | "
                    f"👤 {a['staff_name']} | "
                    f"🧑‍⚕️ {a['auditor_name']} | "
                    f"{status_icon} {a['status']}"
            )

            with c2:
                if st.button(
                    "Open",
                    key=f"history_open_audit_{a['audit_id']}",  # <- UNIQUE KEY FIX
                    use_container_width=True
                ):
                    st.session_state["selected_audit_id"] = a["audit_id"]
                    st.session_state["page"] = "glo_gel_audit_detail"
                    st.rerun()

    # ---------------------------
    # Back button
    # ---------------------------
    st.divider()

    if st.button("Back", use_container_width=True, key="audits_back"):
        st.session_state["page"] = "home"
        st.rerun()

def generate_audit_export_excel(audit_id: int):
    audit = get_audit_header(audit_id)
    if not audit:
        raise ValueError("Audit not found.")

    results = get_audit_results(audit_id)

    standard_results = [r for r in results if r["is_additional"] == "N"]
    additional_results = [r for r in results if r["is_additional"] == "Y"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Audit"

    # ---------------------------
    # Styles
    # ---------------------------
    bold = Font(bold=True)
    title_font = Font(size=14, bold=True)
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ---------------------------
    # Header
    # ---------------------------
    ws["A1"] = "Glo Gel Audit"
    ws["A1"].font = title_font

    ws["A3"] = "Auditor Name"
    ws["B3"] = audit["auditor_name"]

    ws["D3"] = "Staff Name"
    ws["E3"] = audit["staff_name"]

    ws["A4"] = "Date of Inspection"
    ws["B4"] = audit["audit_date"]

    ws["D4"] = "Location"
    ws["E4"] = audit["location_name"]

    ws["A5"] = "Template"
    ws["B5"] = audit["template_group"]

    for cell in ["A3", "D3", "A4", "D4", "A5"]:
        ws[cell].font = bold

    # ---------------------------
    # Surface Results Table
    # ---------------------------
    start_row = 7
    ws[f"A{start_row}"] = "Surface"
    ws[f"B{start_row}"] = "C"
    ws[f"C{start_row}"] = "NC"
    ws[f"D{start_row}"] = "NA"

    for col in ["A", "B", "C", "D"]:
        ws[f"{col}{start_row}"].font = bold
        ws[f"{col}{start_row}"].alignment = center
        ws[f"{col}{start_row}"].border = border

    current_row = start_row + 1

    # Group standard surfaces by area_group
    grouped = {}
    for r in standard_results:
        area = r["area_group"] or "General"
        grouped.setdefault(area, []).append(r)

    for area, items in grouped.items():
        ws[f"A{current_row}"] = area
        ws[f"A{current_row}"].font = bold
        current_row += 1

        items.sort(key=lambda x: (x["display_order"] or 0, x["surface_name"]))

        for r in items:
            ws[f"A{current_row}"] = r["surface_name"]

            result = r["result"]
            ws[f"B{current_row}"] = "C" if result == "C" else ""
            ws[f"C{current_row}"] = "NC" if result == "NC" else ""
            ws[f"D{current_row}"] = "NA" if result == "NA" else ""

            for col in ["A", "B", "C", "D"]:
                ws[f"{col}{current_row}"].border = border

            ws[f"B{current_row}"].alignment = center
            ws[f"C{current_row}"].alignment = center
            ws[f"D{current_row}"].alignment = center

            current_row += 1

        current_row += 1

    # ---------------------------
    # Additional HTS
    # ---------------------------
    ws[f"A{current_row}"] = "Additional HTS"
    ws[f"A{current_row}"].font = bold
    current_row += 1

    ws[f"A{current_row}"] = "Surface"
    ws[f"B{current_row}"] = "C"
    ws[f"C{current_row}"] = "NC"
    ws[f"D{current_row}"] = "NA"

    for col in ["A", "B", "C", "D"]:
        ws[f"{col}{current_row}"].font = bold
        ws[f"{col}{current_row}"].alignment = center
        ws[f"{col}{current_row}"].border = border

    current_row += 1

    if additional_results:
        for r in additional_results:
            ws[f"A{current_row}"] = r["surface_name"]

            result = r["result"]
            ws[f"B{current_row}"] = "C" if result == "C" else ""
            ws[f"C{current_row}"] = "NC" if result == "NC" else ""
            ws[f"D{current_row}"] = "NA" if result == "NA" else ""

            for col in ["A", "B", "C", "D"]:
                ws[f"{col}{current_row}"].border = border

            ws[f"B{current_row}"].alignment = center
            ws[f"C{current_row}"].alignment = center
            ws[f"D{current_row}"].alignment = center

            current_row += 1
    else:
        ws[f"A{current_row}"] = "-"
        for col in ["A", "B", "C", "D"]:
            ws[f"{col}{current_row}"].border = border
        current_row += 1

    current_row += 1

    # ---------------------------
    # Remarks
    # ---------------------------
    ws[f"A{current_row}"] = "Remarks"
    ws[f"A{current_row}"].font = bold
    ws[f"B{current_row}"] = audit["remarks"] or ""

    # ---------------------------
    # Column widths
    # ---------------------------
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 30

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def generate_compiled_glo_gel_report_excel():
    grouped = get_completed_audits_grouped_by_tower()

    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    bold = Font(bold=True)
    title_font = Font(size=14, bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    sheet_specs = [
        ("A", "Tower A"),
        ("B", "Tower B"),
        ("C", "Tower C"),
    ]

    def write_audit_block(ws, start_row, start_col, audit_entry, template_surfaces):
        audit = audit_entry["audit"]
        standard_results = audit_entry["standard_results"]
        additional_results = audit_entry["additional_results"]

        result_lookup = {
            r["surface_id"]: r["result"]
            for r in standard_results
            if r.get("surface_id")
        }

        # Column map for one block
        c0 = start_col       # label/value col 1
        c1 = start_col + 1   # label/value col 2
        c2 = start_col + 2   # result C
        c3 = start_col + 3   # result NC
        c4 = start_col + 4   # result NA

        # Header rows
        ws.cell(start_row, c0, "Date").font = bold
        ws.cell(start_row, c1, audit["audit_date"])

        ws.cell(start_row + 1, c0, "Level").font = bold
        ws.cell(start_row + 1, c1, audit.get("zone") or "")

        ws.cell(start_row + 2, c0, "Location").font = bold
        ws.cell(start_row + 2, c1, audit["location_name"])

        ws.cell(start_row + 3, c0, "Staff").font = bold
        ws.cell(start_row + 3, c1, audit["staff_name"])

        ws.cell(start_row + 4, c0, "Auditor").font = bold
        ws.cell(start_row + 4, c1, audit["auditor_name"])

        ws.cell(start_row + 5, c0, "Remarks").font = bold
        ws.cell(start_row + 5, c1, audit.get("remarks") or "")

        # Surface header
        header_row = start_row + 7
        ws.merge_cells(start_row=header_row, start_column=c0, end_row=header_row, end_column=c1)
        ws.merge_cells(start_row=header_row, start_column=c0, end_row=header_row, end_column=c1)

        for col in range(c0, c1 + 1):
            ws.cell(header_row, col).border = border

        ws.cell(header_row, c0).value = "Surface"
        ws.cell(header_row, c0).font = bold
        ws.cell(header_row, c0).alignment = center
        ws.cell(header_row, c2, "C").font = bold
        ws.cell(header_row, c3, "NC").font = bold
        ws.cell(header_row, c4, "NA").font = bold

        for col in [c0, c2, c3, c4]:
            ws.cell(header_row, col).alignment = center
            ws.cell(header_row, col).border = border

        # Surface rows
        row = header_row + 1
        for s in template_surfaces:
            ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)

            for col in range(c0, c1 + 1):
                cell = ws.cell(row, col)
                cell.border = border

            ws.cell(row, c0).value = s["surface_name"]
            ws.cell(row, c0).alignment = Alignment(vertical="center")

            result = result_lookup.get(s["surface_id"], "")
            ws.cell(row, c2, "C" if result == "C" else "")
            ws.cell(row, c3, "NC" if result == "NC" else "")
            ws.cell(row, c4, "NA" if result == "NA" else "")

            for col in [c2, c3, c4]:
                ws.cell(row, col).alignment = center
                ws.cell(row, col).border = border

            row += 1

        # Additional HTS
        # Additional HTS title row
        ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)

        for col in range(c0, c1 + 1):
            ws.cell(row, col).border = border

        ws.cell(row, c0).value = "Additional HTS"
        ws.cell(row, c0).font = bold

        row += 1

        # Clear entire row first (important)
        for col in range(c0, c4 + 1):
            ws.cell(row, col).border = Border()

        # Now apply only the borders you want
        ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)

        for col in range(c0, c1 + 1):
            ws.cell(row, col).border = border

        ws.cell(row, c0).value = "-"

        for col in [c2, c3, c4]:
            ws.cell(row, col).border = border

        return row

    for tower_code, sheet_name in sheet_specs:
        ws = wb.create_sheet(title=sheet_name)
        audits = grouped.get(tower_code, [])
        template_surfaces = get_surface_template_for_tower(tower_code)

        ws["A1"] = f"Glo Gel Audit Report - {sheet_name}"
        ws["A1"].font = title_font

        if not audits:
            ws["A3"] = "No completed audits."
            continue

        # Layout:
        # left block starts at col 1 (A)
        # right block starts at col 8 (H)
        left_start_col = 1
        right_start_col = 8

        # block height depends on surface count
        # 5 header rows + blank + surface header + surfaces + add hts title + at least 1 row + gap
        block_height = 6 + 1 + len(template_surfaces) + 1 + 1 + 2

        for idx, entry in enumerate(audits):
            pair_index = idx // 2
            is_right = (idx % 2 == 1)

            start_row = 3 + pair_index * block_height
            start_col = right_start_col if is_right else left_start_col

            write_audit_block(ws, start_row, start_col, entry, template_surfaces)

        # widths for left block
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 22
        ws.column_dimensions["C"].width = 8
        ws.column_dimensions["D"].width = 8
        ws.column_dimensions["E"].width = 8

        # spacer columns
        ws.column_dimensions["F"].width = 4
        ws.column_dimensions["G"].width = 4

        # widths for right block
        ws.column_dimensions["H"].width = 18
        ws.column_dimensions["I"].width = 22
        ws.column_dimensions["J"].width = 8
        ws.column_dimensions["K"].width = 8
        ws.column_dimensions["L"].width = 8

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ---------------------------
# Router
# ---------------------------
def router():
    page = st.session_state.get("page", "login")

    if page == "login":
        page_login()
    elif page == "home":
        page_home()
    elif page == "orders":
        page_orders()
    elif page == "order_detail":
        page_order_detail()
    elif page == "issue_stock":
        page_issue_stock()
    elif page == "stock_in":
        page_stock_in()
    elif page == "inventory":
        page_inventory()
    elif page == "monthly_report":
        page_monthly_report()
    elif page == "stock_card":
        page_stock_card()
    elif page == "system_tools":
        page_system_tools()
    elif page == "glo_gel_audits":
        page_glo_gel_audits()
    elif page == "glo_gel_audit_detail":
        page_glo_gel_audit_detail()
    else:
        st.session_state["page"] = "login"
        page_login()


# ---------------------------
# Main
# ---------------------------
def main():
    ensure_bootstrap()
    router()


if __name__ == "__main__":
    main()
