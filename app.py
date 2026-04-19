from flask import Flask, request, jsonify, send_file, session, redirect
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

app = Flask(__name__)
app.secret_key = "parcurs_secret"

FILE = "curse.xlsx"

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
        ws["A1"] = "Foaie de parcurs Flota Comteleprest"

        # HEADER
        ws.append([
            "nr_auto",
            "data",
            "locatie_plecare",
            "locatie_sosire",
            "km_plecare",
            "km_sosire",
            "km_parcurs"
        ])

        stilizeaza_excel(ws)
        wb.save(FILE)

init_excel()

# 🔹 STIL PROFESIONAL
def stilizeaza_excel(ws):
    thin = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # TITLU
    ws["A1"].font = Font(size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")

    # HEADER
    for col in range(1, 8):
        cell = ws.cell(row=2, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="4F81BD", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin

    # DATE + TOTAL
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=7):
        for cell in row:
            cell.border = thin

            if cell.column in [5, 6, 7]:
                cell.alignment = Alignment(horizontal="right")

        if row[0].value == "TOTAL":
            for cell in row:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="FFF2CC", fill_type="solid")

    # AUTO WIDTH
    for col in range(1, 8):
        max_length = 0
        col_letter = get_column_letter(col)

        for cell in ws[col_letter]:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))

        ws.column_dimensions[col_letter].width = max_length + 2

    ws.freeze_panes = "A3"

# 🔹 CALCUL TOTAL PE ZI
def calculeaza_totaluri(ws):
    rows = list(ws.iter_rows(values_only=True))

    # Șterge TOTAL vechi
    for i in range(len(rows), 2, -1):
        if rows[i-1][0] == "TOTAL":
            ws.delete_rows(i)

    data_curenta = None
    total = 0

    for i in range(3, ws.max_row + 1):
        data = ws.cell(i, 2).value
        km = ws.cell(i, 7).value or 0

        if data_curenta is None:
            data_curenta = data

        if data != data_curenta:
            ws.insert_rows(i)
            ws.cell(i, 1, "TOTAL")
            ws.cell(i, 2, data_curenta)
            ws.cell(i, 7, total)

            data_curenta = data
            total = 0

        total += float(km)

    # ULTIMUL TOTAL
    ws.append(["TOTAL", data_curenta, "", "", "", "", total])

# 🔹 CITIRE
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
            "locatie_plecare": row[2] or "",
            "locatie_sosire": row[3] or "",
            "km_plecare": row[4] or "",
            "km_sosire": row[5] or "",
            "km_parcurs": row[6] or ""
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
        return "Login greșit"

    return """
    <html><body style="font-family:Arial;padding:50px">
    <h2>Login</h2>
    <form method="post">
    User: <input name="user"><br><br>
    Parola: <input name="parola" type="password"><br><br>
    <button>Login</button>
    </form></body></html>
    """

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

# 🔹 PAGINA PRINCIPALĂ
@app.route("/")
def index():
    if "user" not in session:
        return redirect("/login")

    curse = citeste_curse()

    html = "<h2>Raport curse</h2><table border=1><tr><th>ID</th><th>Auto</th><th>Data</th><th>Plecare</th><th>Sosire</th><th>Km</th></tr>"

    for c in curse:
        html += f"<tr><td>{c['id']}</td><td>{c['nr_auto']}</td><td>{c['data']}</td><td>{c['locatie_plecare']}</td><td>{c['locatie_sosire']}</td><td>{c['km_parcurs']}</td></tr>"

    html += "</table><br><a href='/download'>Descarca Excel</a>"

    return html

# 🔹 DOWNLOAD
@app.route("/download")
def download_excel():
    return send_file(FILE, as_attachment=True)

# 🔹 ADAUGĂ CURSĂ
@app.route("/adauga_cursa", methods=["POST"])
def adauga_cursa():
    try:
        data = request.get_json(force=True)

        wb = load_workbook(FILE)
        ws = wb.active

        km_plecare = float(data.get("kmPlecare", 0))
        km_sosire = float(data.get("kmSosire", 0))
        km_parcurs = km_sosire - km_plecare

        ws.append([
            data.get("nrAuto", ""),
            data.get("data", ""),
            data.get("locatiePlecare", ""),
            data.get("locatieSosire", ""),
            km_plecare,
            km_sosire,
            km_parcurs
        ])

        calculeaza_totaluri(ws)
        stilizeaza_excel(ws)

        wb.save(FILE)

        return {"status": "ok"}

    except Exception as e:
        print("EROARE:", e)
        return {"status": "error"}

# 🔹 START
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
