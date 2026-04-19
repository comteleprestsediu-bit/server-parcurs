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

# 🔹 INIT
def init_excel():
    if not os.path.exists(FILE):
        wb = Workbook()
        ws = wb.active
        wb.save(FILE)

init_excel()

# 🔹 STIL
def stilizeaza(ws):
    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'), bottom=Side(style='thin'))

    # titlu
    ws.merge_cells("A1:G1")
    ws["A1"] = "Foaie de parcurs Flota Comteleprest"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")

    # header
    headers = ["nr_auto","data","locatie_plecare","locatie_sosire","km_plecare","km_sosire","km_parcurs"]
    ws.append(headers)

    for col in range(1, 8):
        c = ws.cell(2, col)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="4F81BD", fill_type="solid")
        c.alignment = Alignment(horizontal="center")
        c.border = thin

    # borduri + total
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=7):
        for cell in row:
            cell.border = thin
            if cell.column in [5,6,7]:
                cell.alignment = Alignment(horizontal="right")

        if row[0].value == "TOTAL":
            for cell in row:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="FFF2CC", fill_type="solid")

    # auto width
    for col in range(1,8):
        max_len = 0
        col_letter = get_column_letter(col)
        for cell in ws[col_letter]:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_len + 2

    ws.freeze_panes = "A3"

# 🔹 CITEȘTE
def citeste():
    if not os.path.exists(FILE):
        return []

    wb = load_workbook(FILE)
    ws = wb.active

    curse = []

    for row in ws.iter_rows(min_row=3, values_only=True):
        if row[0] == "TOTAL":
            continue

        curse.append({
            "nr_auto": row[0],
            "data": row[1],
            "locatie_plecare": row[2],
            "locatie_sosire": row[3],
            "km_plecare": row[4],
            "km_sosire": row[5],
            "km_parcurs": row[6]
        })

    return curse

# 🔹 RESCRIE COMPLET (SOLUȚIA CHEIE)
def rescrie_excel(curse):
    wb = Workbook()
    ws = wb.active

    # titlu + header
    ws.merge_cells("A1:G1")
    ws["A1"] = "Foaie de parcurs Flota Comteleprest"

    ws.append([
        "nr_auto","data","locatie_plecare","locatie_sosire",
        "km_plecare","km_sosire","km_parcurs"
    ])

    # grupare pe zile
    cur_day = None
    total = 0

    for c in curse:
        km_parcurs = float(c["km_sosire"]) - float(c["km_plecare"])

        if cur_day is None:
            cur_day = c["data"]

        if c["data"] != cur_day:
            ws.append(["TOTAL", cur_day, "", "", "", "", total])
            total = 0
            cur_day = c["data"]

        ws.append([
            c["nr_auto"],
            c["data"],
            c["locatie_plecare"],
            c["locatie_sosire"],
            c["km_plecare"],
            c["km_sosire"],
            km_parcurs
        ])

        total += km_parcurs

    if cur_day:
        ws.append(["TOTAL", cur_day, "", "", "", "", total])

    stilizeaza(ws)
    wb.save(FILE)

# 🔐 LOGIN
@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        if USERS.get(request.form["user"]) == request.form["parola"]:
            session["user"] = request.form["user"]
            return redirect("/")
        return "Login gresit"

    return "<form method=post>User:<input name=user><br>Parola:<input name=parola><br><button>Login</button></form>"

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

# 🔹 PAGINA
@app.route("/")
def index():
    if "user" not in session:
        return redirect("/login")

    curse = citeste()

    html = "<h2>Raport curse</h2><table border=1>"
    html += "<tr><th>Auto</th><th>Data</th><th>Plecare</th><th>Sosire</th><th>Km</th></tr>"

    for c in curse:
        html += f"<tr><td>{c['nr_auto']}</td><td>{c['data']}</td><td>{c['locatie_plecare']}</td><td>{c['locatie_sosire']}</td><td>{c['km_parcurs']}</td></tr>"

    html += "</table><br><a href='/download'>Download Excel</a>"
    return html

# 🔹 DOWNLOAD
@app.route("/download")
def download():
    return send_file(FILE, as_attachment=True)

# 🔹 ADAUGĂ
@app.route("/adauga_cursa", methods=["POST"])
def adauga():
    try:
        data = request.get_json(force=True)

        curse = citeste()

        curse.append({
            "nr_auto": data.get("nrAuto"),
            "data": data.get("data"),
            "locatie_plecare": data.get("locatiePlecare"),
            "locatie_sosire": data.get("locatieSosire"),
            "km_plecare": float(data.get("kmPlecare",0)),
            "km_sosire": float(data.get("kmSosire",0))
        })

        rescrie_excel(curse)

        return {"status":"ok"}

    except Exception as e:
        print("EROARE:", e)
        return {"status":"error"}

# 🔹 START
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
