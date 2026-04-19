from flask import Flask, request, jsonify, send_file, session, redirect
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import sqlite3
import os

app = Flask(__name__)
app.secret_key = "parcurs_secret"

FILE = "curse.xlsx"
DB = "curse.db"

USERS = {
    "admin": "1234",
    "sofer": "1234"
}

# =========================
# INIT EXCEL
# =========================
def init_excel():
    if not os.path.exists(FILE):
        wb = Workbook()
        ws = wb.active

        ws.merge_cells("A1:G1")
        ws["A1"] = "Foaie de parcurs Flota Comteleprest"
        ws["A1"].font = Font(size=14, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center")

        headers = [
            "nr_auto",
            "data",
            "km_plecare",
            "locatie_plecare",
            "km_sosire",
            "locatie_sosire",
            "km_parcurs"
        ]

        ws.append(headers)

        fill = PatternFill(start_color="4F81BD", fill_type="solid")
        for col in range(1, 8):
            c = ws.cell(row=2, column=col)
            c.fill = fill
            c.font = Font(bold=True, color="FFFFFF")

        wb.save(FILE)

init_excel()

# =========================
# INIT DATABASE
# =========================
def init_db():
    conn = sqlite3.connect(DB)
    c = conn.cursor()

    c.execute("""
    CREATE TABLE IF NOT EXISTS curse (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nr_auto TEXT,
        data TEXT,
        km_plecare REAL,
        locatie_plecare TEXT,
        km_sosire REAL,
        locatie_sosire TEXT,
        km_parcurs REAL
    )
    """)

    conn.commit()
    conn.close()

init_db()

# =========================
# TOTAL EXCEL
# =========================
def calculeaza_totaluri(ws):

    for i in range(ws.max_row, 2, -1):
        if ws.cell(i, 1).value == "TOTAL":
            ws.delete_rows(i)

    totaluri = {}

    for i in range(3, ws.max_row + 1):
        data = ws.cell(i, 2).value
        km = ws.cell(i, 7).value or 0

        try:
            km = float(km)
        except:
            km = 0

        totaluri[data] = totaluri.get(data, 0) + km

    rand = ws.max_row + 1
    border = Border(top=Side(style='thin'), bottom=Side(style='thin'))

    for data, total in totaluri.items():
        ws.cell(rand, 1).value = "TOTAL"
        ws.cell(rand, 2).value = data
        ws.cell(rand, 7).value = total

        for col in [1,2,7]:
            c = ws.cell(rand, col)
            c.font = Font(bold=True)
            c.border = border

        rand += 1

# =========================
# CITIRE (din DB, NU Excel)
# =========================
def citeste_curse():
    conn = sqlite3.connect(DB)
    c = conn.cursor()

    c.execute("SELECT * FROM curse ORDER BY id DESC")
    rows = c.fetchall()
    conn.close()

    curse = []

    for r in rows:
        curse.append({
            "id": r[0],
            "nr_auto": r[1],
            "data": r[2],
            "km_plecare": r[3],
            "locatie_plecare": r[4],
            "km_sosire": r[5],
            "locatie_sosire": r[6],
            "km_parcurs": r[7]
        })

    return curse

# =========================
# LOGIN
# =========================
@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        if USERS.get(request.form["user"]) == request.form["parola"]:
            session["user"] = request.form["user"]
            return redirect("/")
        return "Login gresit"

    return """
    <form method="post">
    <input name="user">
    <input name="parola" type="password">
    <button>Login</button>
    </form>
    """

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

# =========================
# UI
# =========================
@app.route("/")
def index():

    if "user" not in session:
        return redirect("/login")

    curse = citeste_curse()

    html = """
    <html>
    <head>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    </head>

    <body class="bg-light">

    <div class="container mt-4">
    <div class="card shadow p-4">

    <div class="d-flex justify-content-between">
        <h3>🚗 Raport curse</h3>
        <a href="/logout" class="btn btn-secondary">Logout</a>
    </div>

    <a href="/download" class="btn btn-success mt-2 mb-3">⬇️ Descarcă Excel</a>

    <table class="table table-striped table-hover">
    <thead class="table-dark">
    <tr>
    <th>ID</th>
    <th>Nr Auto</th>
    <th>Data</th>
    <th>Km Plecare</th>
    <th>Plecare</th>
    <th>Km Sosire</th>
    <th>Sosire</th>
    <th>Km Parcurs</th>
    </tr>
    </thead>
    <tbody>
    """

    total = 0

    for c in curse:
        total += float(c["km_parcurs"] or 0)

        html += f"""
        <tr>
        <td>{c['id']}</td>
        <td>{c['nr_auto']}</td>
        <td>{c['data']}</td>
        <td>{c['km_plecare']}</td>
        <td>{c['locatie_plecare']}</td>
        <td>{c['km_sosire']}</td>
        <td>{c['locatie_sosire']}</td>
        <td>{c['km_parcurs']}</td>
        </tr>
        """

    html += f"""
    </tbody>
    </table>

    <h5>Total km: {total}</h5>

    </div>
    </div>
    </body>
    </html>
    """

    return html

# =========================
# API
# =========================
@app.route("/api/curse")
def api():
    return jsonify(citeste_curse())

# =========================
# ADAUGARE (EXCEL + DB)
# =========================
@app.route("/adauga_cursa", methods=["POST"])
def add():
    try:
        data = request.get_json()

        km_start = float(data.get("kmPlecare",0))
        km_stop = float(data.get("kmSosire",0))
        km_calc = km_stop - km_start

        # 🔹 DB
        conn = sqlite3.connect(DB)
        c = conn.cursor()

        c.execute("""
        INSERT INTO curse (nr_auto, data, km_plecare, locatie_plecare, km_sosire, locatie_sosire, km_parcurs)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            data.get("nrAuto"),
            data.get("data"),
            km_start,
            data.get("locatiePlecare"),
            km_stop,
            data.get("locatieSosire"),
            km_calc
        ))

        conn.commit()
        conn.close()

        # 🔹 EXCEL
        wb = load_workbook(FILE)
        ws = wb.active

        ws.append([
            data.get("nrAuto"),
            data.get("data"),
            km_start,
            data.get("locatiePlecare"),
            km_stop,
            data.get("locatieSosire"),
            km_calc
        ])

        calculeaza_totaluri(ws)
        wb.save(FILE)

        return {"status":"ok"}

    except Exception as e:
        print("EROARE:", e)
        return {"status":"error"}

# =========================
# DOWNLOAD
# =========================
@app.route("/download")
def download():
    return send_file(FILE, as_attachment=True)

# =========================
# START
# =========================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
