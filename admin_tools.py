from db import get_conn


def reset_orders_only():
    """
    Deletes only transactional order data.
    Keeps:
    - users
    - stock movements
    """
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("DELETE FROM order_lines")
    cur.execute("DELETE FROM orders")

    conn.commit()
    conn.close()


def reset_orders_and_movements():
    """
    Deletes:
    - orders
    - order lines
    - stock movements

    Keeps:
    - users
    - workbook-driven master structure
    """
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("DELETE FROM stock_movements")
    cur.execute("DELETE FROM order_lines")
    cur.execute("DELETE FROM orders")

    conn.commit()
    conn.close()