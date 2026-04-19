from flask import Flask, request, jsonify, send_file, session, redirect
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os

app = Flask(__name__)
app.secret_key = "parcurs_secret"

FILE = "curse.xlsx"

# 🔹 USERS
USERS = {
    "admin": "1234",
    "sofer": "1234"
}

# 🔹 INIT EXCEL
def init_excel():
    if not os.path.exists(FILE):
        wb = Workbook()
        ws = wb.active

        # TITLU
        ws.merge_cells("A1:G1")
        title = ws["A1"]
        title.value = "Foaie de parcurs Flota Comteleprest"
        title.font = Font(size=14, bold=True)
        title.alignment = Alignment(horizontal="center")

        # HEADER
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

        fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        font = Font(bold=True, color="FFFFFF")

        for col in range(1, len(headers) + 1):
            c = ws.cell(row=2, column=col)
            c.fill = fill
            c.font = font
            c.alignment = Alignment(horizontal="center")

        # latimi
        widths = [15, 15, 15, 40, 15, 40, 15]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + i)].width = w

        wb.save(FILE)

init_excel()

# 🔹 CALCUL TOTALURI PE ZI
def calculeaza_totaluri(ws):

    # sterge totaluri vechi
    rows_to_delete = []
    for i in range(3, ws.max_row + 1):
        if ws.cell(i, 1).value == "TOTAL":
            rows_to_delete.append(i)

    for i in reversed(rows_to_delete):
        ws.delete_rows(i)

    totaluri = {}

    for i in range(3, ws.max_row + 1):
        data = ws.cell(i, 2).value
        km = ws.cell(i, 7).value or 0

        try:
            km = float(km)
        except:
            km = 0

        if data not in totaluri:
            totaluri[data] = 0

        totaluri[data] += km

    rand = ws.max_row + 1

    border = Border(
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    for data, total in totaluri.items():
        ws.cell(rand, 1).value = "TOTAL"
        ws.cell(rand, 2).value = data
        ws.cell(rand, 7).value = total

        for col in [1, 2, 7]:
            c = ws.cell(rand, col)
            c.font = Font(bold=True)
            c.border = border

        rand += 1

# 🔹 CITESTE CURSE
def citeste_curse():
    wb = load_workbook(FILE)
    ws = wb.active

    curse = []

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 2:
            continue

        if row[0] == "TOTAL":
            continue

        curse.append({
            "id": i - 2,
            "nr_auto": row[0] or "",
            "data": row[1] or "",
            "km_plecare": row[2] or "",
            "locatie_plecare": row[3] or "",
            "km_sosire": row[4] or "",
            "locatie_sosire": row[5] or "",
            "km_parcurs": row[6] or 0
        })

    return curse

# 🔐 LOGIN
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = request.form.get("user")
        parola = request.form.get("parola")

        if USERS.get(user) == parola:
            session["user"] = user
            return redirect("/")
        else:
            return "Login greșit"

    return """
    <html><body>
    <form method="post">
    <input name="user">
    <input name="parola" type="password">
    <button>Login</button>
    </form>
    </body></html>
    """

# 🔓 LOGOUT
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

# 🔥 PAGINA PRINCIPALA
@app.route("/")
def index():
    if "user" not in session:
        return redirect("/login")

    curse = citeste_curse()

    html = """
    <html><body>
    <h2>Raport curse</h2>
    <a href="/download">Descarca Excel</a>
    <table border="1">
    <tr>
    <th>ID</th><th>Nr</th><th>Data</th>
    <th>Km Plecare</th><th>Plecare</th>
    <th>Km Sosire</th><th>Sosire</th>
    <th>Km Parcurs</th>
    </tr>
    """

    total = 0

    for c in curse:
        total += float(c["km_parcurs"])

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

    html += f"</table><h3>Total km: {total}</h3></body></html>"

    return html

# 🔹 DOWNLOAD
@app.route("/download")
def download():
    return send_file(FILE, as_attachment=True)

# 🔹 API LISTA (IMPORTANT FIX)
@app.route("/api/curse")
def api_curse():
    return jsonify(citeste_curse())

# 🔹 ADAUGARE CURSA
@app.route("/adauga_cursa", methods=["POST"])
def adauga():
    try:
        data = request.get_json(silent=True)
        if not data:
            data = request.form

        nr = data.get("nrAuto", "")
        data_c = data.get("data", "")
        plecare = data.get("locatiePlecare", "")
        sosire = data.get("locatieSosire", "")

        km_start = float(data.get("kmPlecare", 0))
        km_stop = float(data.get("kmSosire", 0))
        km_calc = km_stop - km_start

        wb = load_workbook(FILE)
        ws = wb.active

        ws.append([
            nr,
            data_c,
            km_start,
            plecare,
            km_stop,
            sosire,
            km_calc
        ])

        # 🔥 RECALCUL TOTALURI
        calculeaza_totaluri(ws)

        wb.save(FILE)

        return {"status": "ok"}

    except Exception as e:
        print("EROARE:", e)
        return {"status": "error"}

# 🔥 STERGERE
@app.route("/sterge_cursa/<int:id>", methods=["DELETE"])
def sterge(id):
    try:
        wb = load_workbook(FILE)
        ws = wb.active

        ws.delete_rows(id + 3)

        calculeaza_totaluri(ws)

        wb.save(FILE)

        return {"status": "deleted"}

    except Exception as e:
        print("EROARE:", e)
        return {"status": "error"}

# 🔹 START
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
