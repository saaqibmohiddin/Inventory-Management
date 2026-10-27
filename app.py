import os
import sqlite3
import platform
import subprocess
from datetime import datetime

# ============================================================
# OPTIONAL PACKAGES
# ============================================================

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.styles import (
        getSampleStyleSheet,
        ParagraphStyle
    )
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle
    )
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


# ============================================================
# SETTINGS
# ============================================================

DB_NAME = "inventory_system.db"

COMPANY_NAME = "MY INVENTORY STORE"
COMPANY_ADDRESS = "India"
COMPANY_PHONE = "Phone: 9876543210"

INVOICE_FOLDER = "invoices"


# ============================================================
# DATABASE CONNECTION
# ============================================================

def db_connect():
    conn = sqlite3.connect(
        DB_NAME,
        timeout=15
    )

    # Important:
    # The database does NOT use restrictive foreign keys
    # between history and products.
    conn.execute("PRAGMA foreign_keys = OFF")

    return conn


# ============================================================
# INITIALIZE DATABASE
# ============================================================

def initialize_database():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        # ----------------------------------------------------
        # USERS
        # ----------------------------------------------------

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL
            )
        """)

        cursor.execute("""
            INSERT OR IGNORE INTO users
            (
                username,
                password
            )
            VALUES
            (
                'admin',
                'admin123'
            )
        """)


        # ----------------------------------------------------
        # PRODUCTS
        # ----------------------------------------------------

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                category TEXT NOT NULL DEFAULT 'General',
                barcode TEXT NOT NULL DEFAULT '',
                selling_price REAL NOT NULL DEFAULT 0,
                cost_price REAL NOT NULL DEFAULT 0,
                quantity INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL
            )
        """)


        # ----------------------------------------------------
        # SALES
        # ----------------------------------------------------

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_number TEXT NOT NULL,
                customer_name TEXT NOT NULL DEFAULT '',
                product_id INTEGER,
                product_name TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                selling_price REAL NOT NULL,
                cost_price REAL NOT NULL,
                subtotal REAL NOT NULL,
                sale_date TEXT NOT NULL
            )
        """)


        # ----------------------------------------------------
        # PURCHASES
        # ----------------------------------------------------

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS purchases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER,
                product_name TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                cost_price REAL NOT NULL,
                total_cost REAL NOT NULL,
                purchase_date TEXT NOT NULL
            )
        """)


        # ----------------------------------------------------
        # STOCK MOVEMENTS
        # ----------------------------------------------------

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS stock_movements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER,
                product_name TEXT NOT NULL,
                movement_type TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                previous_quantity INTEGER NOT NULL,
                new_quantity INTEGER NOT NULL,
                movement_date TEXT NOT NULL
            )
        """)


        # ----------------------------------------------------
        # BARCODE UNIQUE INDEX
        # ----------------------------------------------------

        cursor.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS
            idx_products_barcode
            ON products(barcode)
            WHERE barcode <> ''
        """)


        conn.commit()

    except Exception:

        conn.rollback()
        raise

    finally:

        conn.close()


# ============================================================
# LOGIN
# ============================================================

def check_login(username, password):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                username
            FROM users
            WHERE username = ?
            AND password = ?
        """, (
            username,
            password
        ))

        return cursor.fetchone()

    finally:

        conn.close()


# ============================================================
# PRODUCT FUNCTIONS
# ============================================================

def get_all_products():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                name,
                category,
                barcode,
                selling_price,
                cost_price,
                quantity
            FROM products
            ORDER BY id DESC
        """)

        return cursor.fetchall()

    finally:

        conn.close()


# ============================================================

def get_product(product_id):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                name,
                category,
                barcode,
                selling_price,
                cost_price,
                quantity
            FROM products
            WHERE id = ?
        """, (
            product_id,
        ))

        return cursor.fetchone()

    finally:

        conn.close()


# ============================================================

def add_product_db(
    name,
    category,
    barcode,
    selling_price,
    cost_price,
    quantity
):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        barcode = barcode.strip()

        # Check duplicate barcode
        if barcode:

            cursor.execute("""
                SELECT id
                FROM products
                WHERE barcode = ?
            """, (
                barcode,
            ))

            existing = cursor.fetchone()

            if existing:

                return (
                    False,
                    "Barcode already exists."
                )


        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )


        cursor.execute("""
            INSERT INTO products
            (
                name,
                category,
                barcode,
                selling_price,
                cost_price,
                quantity,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            name.strip(),
            category.strip(),
            barcode,
            selling_price,
            cost_price,
            quantity,
            now
        ))


        product_id = cursor.lastrowid


        # Stock history
        cursor.execute("""
            INSERT INTO stock_movements
            (
                product_id,
                product_name,
                movement_type,
                quantity,
                previous_quantity,
                new_quantity,
                movement_date
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            product_id,
            name.strip(),
            "ADD",
            quantity,
            0,
            quantity,
            now
        ))


        conn.commit()

        return (
            True,
            "Product added successfully."
        )


    except sqlite3.IntegrityError as error:

        conn.rollback()

        if "barcode" in str(error).lower():

            return (
                False,
                "Barcode already exists."
            )

        return (
            False,
            str(error)
        )


    except Exception as error:

        conn.rollback()

        return (
            False,
            str(error)
        )


    finally:

        conn.close()


# ============================================================

def update_product_db(
    product_id,
    name,
    category,
    barcode,
    selling_price,
    cost_price,
    quantity
):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                name,
                quantity
            FROM products
            WHERE id = ?
        """, (
            product_id,
        ))

        old_product = cursor.fetchone()

        if not old_product:

            return (
                False,
                "Product not found."
            )


        old_quantity = int(
            old_product[1]
        )

        barcode = barcode.strip()


        # Check duplicate barcode
        if barcode:

            cursor.execute("""
                SELECT id
                FROM products
                WHERE barcode = ?
                AND id <> ?
            """, (
                barcode,
                product_id
            ))

            existing = cursor.fetchone()

            if existing:

                return (
                    False,
                    "Barcode already exists."
                )


        cursor.execute("""
            UPDATE products
            SET
                name = ?,
                category = ?,
                barcode = ?,
                selling_price = ?,
                cost_price = ?,
                quantity = ?
            WHERE id = ?
        """, (
            name.strip(),
            category.strip(),
            barcode,
            selling_price,
            cost_price,
            quantity,
            product_id
        ))


        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )


        # Save stock movement when quantity changes
        if old_quantity != quantity:

            if quantity > old_quantity:

                movement_type = "UPDATE + STOCK"

            else:

                movement_type = "UPDATE - STOCK"


            cursor.execute("""
                INSERT INTO stock_movements
                (
                    product_id,
                    product_name,
                    movement_type,
                    quantity,
                    previous_quantity,
                    new_quantity,
                    movement_date
                )
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                product_id,
                name.strip(),
                movement_type,
                abs(quantity - old_quantity),
                old_quantity,
                quantity,
                now
            ))


        conn.commit()

        return (
            True,
            "Product updated successfully."
        )


    except Exception as error:

        conn.rollback()

        return (
            False,
            str(error)
        )


    finally:

        conn.close()


# ============================================================
# DELETE PRODUCT
# ============================================================

def delete_product_db(product_id):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        # Find product
        cursor.execute("""
            SELECT
                name
            FROM products
            WHERE id = ?
        """, (
            product_id,
        ))

        product = cursor.fetchone()

        if not product:

            return (
                False,
                "Product not found."
            )


        product_name = product[0]


        # IMPORTANT:
        # We only delete the product.
        # Sales, purchases and stock history remain.
        # Therefore there is no foreign-key error.

        cursor.execute("""
            DELETE FROM products
            WHERE id = ?
        """, (
            product_id,
        ))


        if cursor.rowcount != 1:

            conn.rollback()

            return (
                False,
                "Product could not be deleted."
            )


        conn.commit()


        return (
            True,
            f"{product_name} deleted successfully."
        )


    except Exception as error:

        conn.rollback()

        return (
            False,
            str(error)
        )


    finally:

        conn.close()


# ============================================================
# SEARCH / FILTER
# ============================================================

def filter_products_db(
    search_text="",
    category="All Categories",
    stock_status_filter="All Stock"
):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        query = """
            SELECT
                id,
                name,
                category,
                barcode,
                selling_price,
                cost_price,
                quantity
            FROM products
            WHERE 1=1
        """

        parameters = []


        # Search
        if search_text.strip():

            search_pattern = (
                "%"
                +
                search_text.strip()
                +
                "%"
            )

            query += """
                AND (
                    name LIKE ?
                    OR category LIKE ?
                    OR barcode LIKE ?
                )
            """

            parameters.extend([
                search_pattern,
                search_pattern,
                search_pattern
            ])


        # Category
        if (
            category
            and
            category != "All Categories"
        ):

            query += """
                AND category = ?
            """

            parameters.append(
                category
            )


        # Stock
        if stock_status_filter == "Out of Stock":

            query += """
                AND quantity = 0
            """

        elif stock_status_filter == "Low Stock":

            query += """
                AND quantity > 0
                AND quantity <= 5
            """

        elif stock_status_filter == "Normal":

            query += """
                AND quantity > 5
            """


        query += """
            ORDER BY id DESC
        """


        cursor.execute(
            query,
            parameters
        )

        return cursor.fetchall()

    finally:

        conn.close()


# ============================================================
# DASHBOARD
# ============================================================

def dashboard_data():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        # Total products
        cursor.execute("""
            SELECT COUNT(*)
            FROM products
        """)

        total_products = (
            cursor.fetchone()[0]
        )


        # Total stock
        cursor.execute("""
            SELECT COALESCE(
                SUM(quantity),
                0
            )
            FROM products
        """)

        total_stock = (
            cursor.fetchone()[0]
        )


        # Low stock
        cursor.execute("""
            SELECT COUNT(*)
            FROM products
            WHERE quantity > 0
            AND quantity <= 5
        """)

        low_stock = (
            cursor.fetchone()[0]
        )


        # Out of stock
        cursor.execute("""
            SELECT COUNT(*)
            FROM products
            WHERE quantity = 0
        """)

        out_of_stock = (
            cursor.fetchone()[0]
        )


        # Inventory value
        cursor.execute("""
            SELECT COALESCE(
                SUM(
                    selling_price * quantity
                ),
                0
            )
            FROM products
        """)

        inventory_value = (
            cursor.fetchone()[0]
        )


        return (
            total_products,
            total_stock,
            low_stock,
            out_of_stock,
            inventory_value
        )

    finally:

        conn.close()


# ============================================================
# STOCK STATUS
# ============================================================

def get_stock_status(quantity):

    quantity = int(quantity)

    if quantity == 0:

        return "Out of Stock"

    if quantity <= 5:

        return "Low Stock"

    return "Normal"


# ============================================================
# SALES
# ============================================================

def next_invoice_number():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        today = datetime.now().strftime(
            "%Y%m%d"
        )

        cursor.execute("""
            SELECT COUNT(*)
            FROM sales
            WHERE sale_date LIKE ?
        """, (
            today + "%",
        ))

        count = (
            cursor.fetchone()[0]
            +
            1
        )

        return (
            f"INV-{today}-{count:04d}"
        )

    finally:

        conn.close()


# ============================================================

def make_sale(
    product_id,
    quantity,
    customer_name
):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                name,
                selling_price,
                cost_price,
                quantity
            FROM products
            WHERE id = ?
        """, (
            product_id,
        ))

        product = cursor.fetchone()

        if not product:

            return (
                False,
                "Product not found."
            )


        (
            product_name,
            selling_price,
            cost_price,
            current_stock
        ) = product


        if quantity <= 0:

            return (
                False,
                "Quantity must be greater than 0."
            )


        if quantity > current_stock:

            return (
                False,
                f"Not enough stock.\n"
                f"Available stock: {current_stock}"
            )


        invoice_number = (
            next_invoice_number()
        )

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )


        subtotal = (
            float(selling_price)
            *
            quantity
        )


        new_stock = (
            current_stock
            -
            quantity
        )


        # Update stock
        cursor.execute("""
            UPDATE products
            SET quantity = ?
            WHERE id = ?
        """, (
            new_stock,
            product_id
        ))


        # Save sale
        cursor.execute("""
            INSERT INTO sales
            (
                invoice_number,
                customer_name,
                product_id,
                product_name,
                quantity,
                selling_price,
                cost_price,
                subtotal,
                sale_date
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            invoice_number,
            customer_name,
            product_id,
            product_name,
            quantity,
            selling_price,
            cost_price,
            subtotal,
            now
        ))


        sale_id = cursor.lastrowid


        # Stock history
        cursor.execute("""
            INSERT INTO stock_movements
            (
                product_id,
                product_name,
                movement_type,
                quantity,
                previous_quantity,
                new_quantity,
                movement_date
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            product_id,
            product_name,
            "SALE",
            quantity,
            current_stock,
            new_stock,
            now
        ))


        conn.commit()


        return (
            True,
            {
                "sale_id": sale_id,
                "invoice": invoice_number,
                "customer": customer_name,
                "product": product_name,
                "quantity": quantity,
                "selling_price": float(
                    selling_price
                ),
                "cost_price": float(
                    cost_price
                ),
                "subtotal": subtotal,
                "date": now
            }
        )


    except Exception as error:

        conn.rollback()

        return (
            False,
            str(error)
        )


    finally:

        conn.close()


# ============================================================
# SALES HISTORY
# ============================================================

def get_sales_history():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                invoice_number,
                customer_name,
                product_name,
                quantity,
                selling_price,
                subtotal,
                cost_price,
                sale_date
            FROM sales
            ORDER BY id DESC
        """)

        return cursor.fetchall()

    finally:

        conn.close()


# ============================================================
# SALES SUMMARY
# ============================================================

def sales_summary():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        # Total sales
        cursor.execute("""
            SELECT COALESCE(
                SUM(subtotal),
                0
            )
            FROM sales
        """)

        total_sales = (
            cursor.fetchone()[0]
        )


        # Today's sales
        today = datetime.now().strftime(
            "%Y-%m-%d"
        )

        cursor.execute("""
            SELECT COALESCE(
                SUM(subtotal),
                0
            )
            FROM sales
            WHERE sale_date LIKE ?
        """, (
            today + "%",
        ))

        today_sales = (
            cursor.fetchone()[0]
        )


        # Transactions
        cursor.execute("""
            SELECT COUNT(*)
            FROM sales
        """)

        transactions = (
            cursor.fetchone()[0]
        )


        # Profit
        cursor.execute("""
            SELECT COALESCE(
                SUM(
                    subtotal
                    -
                    (
                        quantity
                        *
                        cost_price
                    )
                ),
                0
            )
            FROM sales
        """)

        profit = (
            cursor.fetchone()[0]
        )


        return (
            today_sales,
            total_sales,
            transactions,
            profit
        )

    finally:

        conn.close()


# ============================================================
# PURCHASES
# ============================================================

def make_purchase(
    product_id,
    quantity,
    purchase_cost
):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                name,
                quantity,
                cost_price
            FROM products
            WHERE id = ?
        """, (
            product_id,
        ))

        product = cursor.fetchone()

        if not product:

            return (
                False,
                "Product not found."
            )


        (
            product_name,
            old_quantity,
            old_cost
        ) = product


        if quantity <= 0:

            return (
                False,
                "Quantity must be greater than 0."
            )


        if purchase_cost <= 0:

            return (
                False,
                "Purchase cost must be greater than 0."
            )


        new_quantity = (
            old_quantity
            +
            quantity
        )


        # Weighted average cost
        new_cost = (
            (
                old_quantity
                *
                float(old_cost)
            )
            +
            (
                quantity
                *
                float(purchase_cost)
            )
        ) / new_quantity


        total_cost = (
            quantity
            *
            float(purchase_cost)
        )


        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )


        # Update product
        cursor.execute("""
            UPDATE products
            SET
                quantity = ?,
                cost_price = ?
            WHERE id = ?
        """, (
            new_quantity,
            new_cost,
            product_id
        ))


        # Save purchase
        cursor.execute("""
            INSERT INTO purchases
            (
                product_id,
                product_name,
                quantity,
                cost_price,
                total_cost,
                purchase_date
            )
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            product_id,
            product_name,
            quantity,
            purchase_cost,
            total_cost,
            now
        ))


        # Stock history
        cursor.execute("""
            INSERT INTO stock_movements
            (
                product_id,
                product_name,
                movement_type,
                quantity,
                previous_quantity,
                new_quantity,
                movement_date
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            product_id,
            product_name,
            "PURCHASE",
            quantity,
            old_quantity,
            new_quantity,
            now
        ))


        conn.commit()


        return (
            True,
            {
                "product": product_name,
                "quantity": quantity,
                "cost": float(
                    purchase_cost
                ),
                "total": total_cost,
                "date": now
            }
        )


    except Exception as error:

        conn.rollback()

        return (
            False,
            str(error)
        )


    finally:

        conn.close()


# ============================================================
# PURCHASE HISTORY
# ============================================================

def get_purchase_history():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                product_name,
                quantity,
                cost_price,
                total_cost,
                purchase_date
            FROM purchases
            ORDER BY id DESC
        """)

        return cursor.fetchall()

    finally:

        conn.close()


# ============================================================
# STOCK HISTORY
# ============================================================

def get_stock_history():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                product_name,
                movement_type,
                quantity,
                previous_quantity,
                new_quantity,
                movement_date
            FROM stock_movements
            ORDER BY id DESC
        """)

        return cursor.fetchall()

    finally:

        conn.close()


# ============================================================
# OPEN FILE
# ============================================================

def open_file(path):
    # Desktop file opening is not available in Streamlit Cloud.
    return None


# ============================================================
# CREATE PDF INVOICE
# ============================================================

def generate_invoice_pdf(data):

    if not REPORTLAB_AVAILABLE:

        raise RuntimeError(
            "ReportLab is not installed.\n\n"
            "Run:\n"
            "pip install reportlab"
        )


    os.makedirs(
        INVOICE_FOLDER,
        exist_ok=True
    )


    safe_invoice = ""

    for char in data["invoice"]:

        if (
            char.isalnum()
            or char in "-_"
        ):

            safe_invoice += char

        else:

            safe_invoice += "_"


    path = os.path.join(
        INVOICE_FOLDER,
        safe_invoice + ".pdf"
    )


    styles = getSampleStyleSheet()


    title_style = ParagraphStyle(
        "InvoiceTitle",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=20
    )


    right_style = ParagraphStyle(
        "Right",
        parent=styles["Normal"],
        alignment=TA_RIGHT
    )


    document = SimpleDocTemplate(
        path,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )


    story = []


    story.append(
        Paragraph(
            COMPANY_NAME,
            title_style
        )
    )


    story.append(
        Paragraph(
            f"{COMPANY_ADDRESS} | "
            f"{COMPANY_PHONE}",
            styles["Normal"]
        )
    )


    story.append(
        Spacer(1, 15)
    )


    story.append(
        Paragraph(
            f"<b>Invoice:</b> "
            f"{data['invoice']}",
            styles["Normal"]
        )
    )


    story.append(
        Paragraph(
            f"<b>Date:</b> "
            f"{data['date']}",
            styles["Normal"]
        )
    )


    story.append(
        Paragraph(
            f"<b>Customer:</b> "
            f"{data['customer']}",
            styles["Normal"]
        )
    )


    story.append(
        Spacer(1, 15)
    )


    table_data = [
        [
            "Product",
            "Qty",
            "Selling Price",
            "Subtotal"
        ],
        [
            data["product"],
            str(data["quantity"]),
            f"Rs. "
            f"{data['selling_price']:,.2f}",
            f"Rs. "
            f"{data['subtotal']:,.2f}"
        ]
    ]


    table = Table(
        table_data,
        colWidths=[
            230,
            55,
            100,
            100
        ]
    )


    table.setStyle(
        TableStyle([
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.lightgrey
            ),
            (
                "FONTNAME",
                (0, 0),
                (-1, 0),
                "Helvetica-Bold"
            ),
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.grey
            ),
            (
                "ALIGN",
                (1, 1),
                (-1, -1),
                "RIGHT"
            ),
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),
            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                8
            ),
            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                8
            )
        ])
    )


    story.append(table)


    story.append(
        Spacer(1, 20)
    )


    story.append(
        Paragraph(
            f"<b>Grand Total: "
            f"Rs. {data['subtotal']:,.2f}</b>",
            right_style
        )
    )


    story.append(
        Spacer(1, 30)
    )


    story.append(
        Paragraph(
            "Thank you for your purchase!",
            ParagraphStyle(
                "Thanks",
                parent=styles["Normal"],
                alignment=TA_CENTER
            )
        )
    )


    document.build(story)


    return path



# ============================================================
# STREAMLIT WEB INTERFACE
# ============================================================

import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="StockWise | Inventory Management",
    page_icon="📦",
    layout="wide"
)

initialize_database()

st.title("📦 StockWise")
st.caption("Inventory Management System")

# Existing default account from the original application:
# username: admin | password: admin123
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "username" not in st.session_state:
    st.session_state.username = ""

if not st.session_state.authenticated:
    st.subheader("Login")
    with st.form("login_form"):
        username = st.text_input("Username", value="admin")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login", type="primary")
    if submitted:
        user = check_login(username.strip(), password)
        if user:
            st.session_state.authenticated = True
            st.session_state.username = user[1]
            st.rerun()
        else:
            st.error("Invalid username or password.")
    st.info("Initial login from the original project: admin / admin123. Change this before public use.")
    st.stop()

with st.sidebar:
    st.write(f"Logged in as **{st.session_state.username}**")
    if st.button("Logout"):
        st.session_state.authenticated = False
        st.session_state.username = ""
        st.rerun()

tabs = st.tabs([
    "Dashboard", "Products", "Add Product", "Sales",
    "Purchases", "Sales History", "Purchase History",
    "Stock History", "Reports"
])

# Dashboard
with tabs[0]:
    st.header("Dashboard")
    total_products, total_stock, low_stock, out_of_stock, inventory_value = dashboard_data()
    today_sales, all_sales, transactions, profit = sales_summary()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Products", total_products)
    c2.metric("Total Stock Units", total_stock)
    c3.metric("Low Stock Items", low_stock)
    c4.metric("Out of Stock", out_of_stock)
    c5, c6, c7 = st.columns(3)
    c5.metric("Inventory Value", f"₹{inventory_value:,.2f}")
    c6.metric("Today's Sales", f"₹{today_sales:,.2f}")
    c7.metric("Total Profit", f"₹{profit:,.2f}")
    st.subheader("Current Inventory")
    products = get_all_products()
    if products:
        df = pd.DataFrame(products, columns=[
            "ID", "Product", "Category", "Barcode",
            "Selling Price", "Cost Price", "Quantity"
        ])
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("No products added yet.")

# Products/search/filter/update/delete
with tabs[1]:
    st.header("Product Inventory")
    search = st.text_input("Search by product, category or barcode")
    all_rows = get_all_products()
    categories = ["All Categories"] + sorted({r[2] for r in all_rows if r[2]})
    category = st.selectbox("Category", categories)
    stock_filter = st.selectbox("Stock status", [
        "All Stock", "Normal", "Low Stock", "Out of Stock"
    ])
    rows = filter_products_db(search, category, stock_filter)
    if rows:
        df = pd.DataFrame(rows, columns=[
            "ID", "Product", "Category", "Barcode",
            "Selling Price", "Cost Price", "Quantity"
        ])
        st.dataframe(df, use_container_width=True, hide_index=True)
        ids = [r[0] for r in rows]
        selected_id = st.selectbox("Select product to edit/delete", ids,
                                   format_func=lambda x: next((f"{r[1]} (ID {r[0]})" for r in rows if r[0] == x), str(x)))
        selected = get_product(selected_id)
        with st.form("edit_product"):
            name = st.text_input("Product name", selected[1])
            cat = st.text_input("Category", selected[2])
            barcode = st.text_input("Barcode", selected[3])
            col1, col2, col3 = st.columns(3)
            selling = col1.number_input("Selling price (₹)", min_value=0.0, value=float(selected[4]), step=1.0)
            cost = col2.number_input("Cost price (₹)", min_value=0.0, value=float(selected[5]), step=1.0)
            qty = col3.number_input("Quantity", min_value=0, value=int(selected[6]), step=1)
            save = st.form_submit_button("Update Product")
        if save:
            ok, msg = update_product_db(selected_id, name, cat, barcode, selling, cost, int(qty))
            (st.success if ok else st.error)(msg)
            if ok: st.rerun()
        if st.button("Delete Selected Product", type="secondary"):
            ok, msg = delete_product_db(selected_id)
            (st.success if ok else st.error)(msg)
            if ok: st.rerun()
    else:
        st.info("No matching products found.")

# Add product
with tabs[2]:
    st.header("Add New Product")
    with st.form("add_product_form", clear_on_submit=True):
        name = st.text_input("Product name")
        category = st.selectbox("Category", [
            "Electronics", "Accessories", "Furniture", "Stationery",
            "Clothing", "Food", "Other", "General"
        ])
        barcode = st.text_input("Barcode (optional)")
        c1, c2, c3 = st.columns(3)
        selling = c1.number_input("Selling price (₹)", min_value=0.0, step=1.0)
        cost = c2.number_input("Cost price (₹)", min_value=0.0, step=1.0)
        quantity = c3.number_input("Opening quantity", min_value=0, step=1)
        submit = st.form_submit_button("Add Product", type="primary")
    if submit:
        if not name.strip():
            st.error("Product name is required.")
        else:
            ok, msg = add_product_db(name, category, barcode, selling, cost, int(quantity))
            (st.success if ok else st.error)(msg)

# Sales
with tabs[3]:
    st.header("Record a Sale")
    products = get_all_products()
    available = [p for p in products if p[6] > 0]
    if available:
        options = {f"{p[1]} | Stock: {p[6]} | ₹{p[4]:.2f}": p for p in available}
        with st.form("sale_form"):
            choice = st.selectbox("Product", list(options))
            p = options[choice]
            customer = st.text_input("Customer name")
            quantity = st.number_input("Quantity", min_value=1, max_value=int(p[6]), step=1)
            submit = st.form_submit_button("Complete Sale", type="primary")
        if submit:
            ok, result = make_sale(p[0], int(quantity), customer)
            if ok:
                st.success(f"Sale recorded. Invoice: {result['invoice']} | Total: ₹{result['subtotal']:.2f}")
                try:
                    pdf_path = generate_invoice_pdf(result)
                    with open(pdf_path, "rb") as f:
                        st.download_button("Download Invoice PDF", f.read(),
                                           file_name=os.path.basename(pdf_path),
                                           mime="application/pdf")
                except Exception as e:
                    st.warning(f"Sale saved, but invoice PDF could not be generated: {e}")
            else:
                st.error(result)
    else:
        st.info("No products with available stock.")

# Purchases
with tabs[4]:
    st.header("Record a Purchase / Restock")
    products = get_all_products()
    if products:
        options = {f"{p[1]} | Current stock: {p[6]}": p for p in products}
        with st.form("purchase_form"):
            choice = st.selectbox("Product", list(options))
            p = options[choice]
            quantity = st.number_input("Quantity purchased", min_value=1, step=1)
            purchase_cost = st.number_input("Cost per unit (₹)", min_value=0.01, step=1.0)
            submit = st.form_submit_button("Save Purchase", type="primary")
        if submit:
            ok, result = make_purchase(p[0], int(quantity), float(purchase_cost))
            if ok: st.success(f"Purchase saved: {result['product']} | Total ₹{result['total']:.2f}")
            else: st.error(result)
    else:
        st.info("Add a product before recording a purchase.")

# Histories
with tabs[5]:
    st.header("Sales History")
    rows = get_sales_history()
    if rows:
        st.dataframe(pd.DataFrame(rows, columns=[
            "ID", "Invoice", "Customer", "Product", "Quantity",
            "Selling Price", "Subtotal", "Cost Price", "Date"
        ]), use_container_width=True, hide_index=True)
    else: st.info("No sales recorded.")

with tabs[6]:
    st.header("Purchase History")
    rows = get_purchase_history()
    if rows:
        st.dataframe(pd.DataFrame(rows, columns=[
            "ID", "Product", "Quantity", "Cost Price", "Total Cost", "Date"
        ]), use_container_width=True, hide_index=True)
    else: st.info("No purchases recorded.")

with tabs[7]:
    st.header("Stock Movement History")
    rows = get_stock_history()
    if rows:
        st.dataframe(pd.DataFrame(rows, columns=[
            "ID", "Product", "Movement", "Quantity",
            "Previous Stock", "New Stock", "Date"
        ]), use_container_width=True, hide_index=True)
    else: st.info("No stock movements recorded.")

with tabs[8]:
    st.header("Reports")
    products = get_all_products()
    if products:
        report_df = pd.DataFrame(products, columns=[
            "ID", "Product", "Category", "Barcode",
            "Selling Price", "Cost Price", "Quantity"
        ])
        st.download_button(
            "Download Inventory Report (CSV)",
            report_df.to_csv(index=False).encode("utf-8"),
            file_name="stockwise_inventory_report.csv",
            mime="text/csv"
        )
    st.write("Sales summary")
    today_sales, all_sales, transactions, profit = sales_summary()
    st.metric("Total Sales", f"₹{all_sales:,.2f}")
    st.metric("Transactions", transactions)
    st.metric("Profit", f"₹{profit:,.2f}")

st.caption("StockWise Inventory Management System")
