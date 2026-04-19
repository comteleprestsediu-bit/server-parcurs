from flask import Flask, request, jsonify, send_file, session, redirect
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
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

        ws.merge_cells("A1:G1")
        title = ws["A1"]
        title.value = "Foaie de parcurs Flota Comteleprest"
        title.font = Font(size=14, bold=True)
        title.alignment = Alignment(horizontal="center")

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
            cell = ws.cell(row=2, column=col)
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(horizontal="center")

        wb.save(FILE)

init_excel()


# 🔹 TOTAL PE ZI
def calculeaza_total_pe_zi(ws):
    rows = list(ws.iter_rows(values_only=True))

    # șterge TOTAL vechi
    for i in range(len(rows), 2, -1):
        if rows[i-1][0] == "TOTAL":
            ws.delete_rows(i)

    current_date = None
    total = 0

    for i in range(3, ws.max_row + 1):
        data = ws.cell(i, 2).value
        km = ws.cell(i, 7).value

        km = float(km) if km not in (None, "", " ") else 0

        if current_date is None:
            current_date = data

        if data != current_date:
            ws.insert_rows(i)
            ws.cell(i, 1, "TOTAL")
            ws.cell(i, 2, current_date)
            ws.cell(i, 7, total)

            current_date = data
            total = 0

        total += km

    if current_date:
        ws.append(["TOTAL", current_date, "", "", "", "", total])


# 🔹 CITESTE CURSE (CORECT)
def citeste_curse():
    wb = load_workbook(FILE)
    ws = wb.active

    curse = []

    for row in ws.iter_rows(min_row=3, values_only=True):
        if row[0] == "TOTAL":
            continue

        curse.append({
            "id": len(curse),
            "nr_auto": row[0] or "",
            "data": row[1] or "",
            "locatie_plecare": row[3] or "",
            "locatie_sosire": row[5] or "",
            "km_parcurs": str(row[6] or "")
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
    <html>
    <head>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    </head>
    <body class="bg-light">
    <div class="container mt-5">
        <div class="card p-4 shadow" style="max-width:400px;margin:auto;">
            <h3>Login</h3>
            <form method="post">
                <input name="user" class="form-control mb-2">
                <input name="parola" type="password" class="form-control mb-2">
                <button class="btn btn-primary w-100">Login</button>
            </form>
        </div>
    </div>
    </body>
    </html>
    """


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

    total = 0
    for c in curse:
        try:
            total += float(c["km_parcurs"])
        except:
            pass

    html = """
    <html>
    <head>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    </head>
    <body>
    <div class="container mt-4">
        <div class="card p-4 shadow">

        <div class="d-flex justify-content-between">
            <h3>🚗 Raport curse</h3>
            <a href="/logout" class="btn btn-secondary">Logout</a>
        </div>

        <a href="/download" class="btn btn-success mt-2 mb-3">⬇️ Excel</a>

        <table class="table table-striped">
        <thead class="table-dark">
        <tr>
        <th>ID</th><th>Nr Auto</th><th>Data</th><th>Plecare</th><th>Sosire</th><th>Km</th>
        </tr>
        </thead><tbody>
    """

    for c in curse:
        html += f"""
        <tr>
        <td>{c['id']}</td>
        <td>{c['nr_auto']}</td>
        <td>{c['data']}</td>
        <td>{c['locatie_plecare']}</td>
        <td>{c['locatie_sosire']}</td>
        <td>{c['km_parcurs']}</td>
        </tr>
        """

    html += f"""
    </tbody></table>
    <h5>Total KM: {total}</h5>
    </div></div></body></html>
    """

    return html


# 🔹 DOWNLOAD
@app.route("/download")
def download_excel():
    return send_file(FILE, as_attachment=True)


# 🔹 ADAUGARE CURSA (100% COMPATIBIL TELEFON)
@app.route("/adauga_cursa", methods=["POST"])
def adauga_cursa():
    try:
        data = request.get_json(silent=True)

        if not data:
            data = request.form

        nr_auto = data.get("nrAuto") or data.get("nr_auto", "")
        data_cursa = data.get("data", "")
        plecare = data.get("locatiePlecare") or data.get("locatie_plecare", "")
        sosire = data.get("locatieSosire") or data.get("locatie_sosire", "")

        km_parcurs = data.get("kmParcurs") or data.get("km_parcurs")

        # NU forțăm valori
        km_plecare = data.get("kmPlecare") or ""
        km_sosire = data.get("kmSosire") or ""

        wb = load_workbook(FILE)
        ws = wb.active

        ws.append([
            nr_auto,
            data_cursa,
            km_plecare,
            plecare,
            km_sosire,
            sosire,
            km_parcurs
        ])

        calculeaza_total_pe_zi(ws)

        wb.save(FILE)

        return {"status": "ok"}

    except Exception as e:
        print("EROARE:", e)
        return {"status": "error"}


# 🔹 START
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
