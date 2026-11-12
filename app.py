import os
from flask import Flask, render_template, request, redirect, url_for, session, send_file
from database import get_connection
from openpyxl import Workbook, load_workbook

app = Flask(__name__)
app.secret_key = "food_ordering_secret_key"

EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "orders.xlsx")


# Export Orders to Excel
def export_orders_to_excel():

    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute("""
            SELECT
                o.order_id,
                c.id,
                c.name,
                c.phone,
                c.address,
                o.total_amount,
                o.order_date,
                o.status
            FROM customer c
            JOIN orders o ON c.id = o.customer_id
            ORDER BY o.order_id DESC
        """)

        records = cursor.fetchall()

        if os.path.exists(EXCEL_FILE):
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active

        ws.title = "Orders"
        ws.delete_rows(1, ws.max_row)

        ws.append([
            "Order ID",
            "Customer ID",
            "Name",
            "Mobile",
            "Address",
            "Total Amount",
            "Order Date",
            "Status"
        ])

        for record in records:
            ws.append(list(record))

        wb.save(EXCEL_FILE)

    finally:
        cursor.close()
        conn.close()


# Home Page
@app.route('/')
def home():
    return render_template('index.html')


# Save Customer and Place Order
@app.route('/save_customer', methods=['POST'])
def save_customer():

    name = request.form.get('name')
    email = request.form.get('email', '')
    mobile = request.form.get('mobile')
    address = request.form.get('address')
    total_amount = request.form.get('total_amount')

    if not all([name, mobile, address, total_amount]):
        return "Please fill all details"

    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute("""
            INSERT INTO customer (name, email, phone, address)
            VALUES (%s, %s, %s, %s)
        """, (name, email, mobile, address))

        customer_id = cursor.lastrowid

        cursor.execute("""
            INSERT INTO orders
            (customer_id, total_amount, status)
            VALUES (%s, %s, %s)
        """, (customer_id, total_amount, "Placed"))

        conn.commit()

    except Exception as e:
        conn.rollback()
        return "Error: " + str(e)

    finally:
        cursor.close()
        conn.close()

    try:
        export_orders_to_excel()
    except Exception as e:
        return "Order saved, but Excel export failed: " + str(e)

    return "Order Placed Successfully!"


# Admin Login
@app.route('/admin_login', methods=['GET', 'POST'])
def admin_login():

    if request.method == 'POST':

        username = request.form.get('username')
        password = request.form.get('password')

        if username == 'admin' and password == 'admin123':
            session['admin_logged_in'] = True
            return redirect(url_for('admin'))

        return "Invalid Username or Password"

    return render_template('admin_login.html')


# Admin Dashboard
@app.route('/admin')
def admin():

    if not session.get('admin_logged_in'):
        return redirect(url_for('admin_login'))

    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute("""
            SELECT
                o.order_id,
                c.id,
                c.name,
                c.phone,
                c.address,
                o.total_amount,
                o.order_date,
                o.status
            FROM customer c
            JOIN orders o ON c.id = o.customer_id
            ORDER BY o.order_id DESC
        """)

        orders = cursor.fetchall()

        cursor.execute("""
            SELECT COUNT(*), COALESCE(SUM(total_amount), 0)
            FROM orders
        """)

        result = cursor.fetchone()

        total_orders = result[0]
        total_sales = result[1]

        return render_template(
            'admin.html',
            orders=orders,
            total_orders=total_orders,
            total_sales=total_sales
        )

    finally:
        cursor.close()
        conn.close()


# Update Order Status
@app.route('/update_status', methods=['POST'])
def update_status():

    if not session.get('admin_logged_in'):
        return redirect(url_for('admin_login'))

    order_id = request.form.get('order_id')
    status = request.form.get('status')

    allowed_statuses = ["Placed", "Preparing", "Delivered"]

    if not order_id or status not in allowed_statuses:
        return "Invalid order or status"

    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute("""
            UPDATE orders
            SET status = %s
            WHERE order_id = %s
        """, (status, order_id))

        conn.commit()

    except Exception as e:
        conn.rollback()
        return "Error: " + str(e)

    finally:
        cursor.close()
        conn.close()

    export_orders_to_excel()

    return redirect(url_for('admin'))


# Download Excel
@app.route('/download_excel')
def download_excel():

    if not session.get('admin_logged_in'):
        return redirect(url_for('admin_login'))

    if os.path.exists(EXCEL_FILE):
        return send_file(EXCEL_FILE, as_attachment=True)

    return "Excel file not found"


# Admin Logout
@app.route('/admin_logout')
def admin_logout():

    session.pop('admin_logged_in', None)
    return redirect(url_for('admin_login'))


if __name__ == '__main__':
    app.run(debug=True)