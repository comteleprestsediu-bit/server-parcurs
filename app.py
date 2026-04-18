from flask import Flask, request, jsonify, send_file, redirect, session
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import os

app = Flask(__name__)
app.secret_key = "secret123"

FILE = "curse.xlsx"

USER = "admin"
PASS = "1234"

# INIT EXCEL
def init_excel():
    if not os.path.exists(FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Raport Curse"

        ws.append([
            "Nr Auto",
            "Data",
            "Locatie Plecare",
            "Locatie Sosire",
            "Km Parcurs"
        ])

        wb.save(FILE)

init_excel()

# LOGIN
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        if request.form.get("user") == USER and request.form.get("pass") == PASS:
            session["login"] = True
            return redirect("/")
        return "Date incorecte"

    return """
    <html>
    <style>
    body {font-family: Arial; background:#f2f4f7;}
    .box {width:300px;margin:auto;margin-top:100px;padding:20px;background:white;border-radius:8px;box-shadow:0 0 10px #ccc;}
    input {width:100%;padding:10px;margin:5px 0;}
    button {width:100%;padding:10px;background:#6200EE;color:white;border:none;}
    </style>
    <div class="box">
        <h2>Login</h2>
        <form method="post">
            <input name="user" placeholder="User">
            <input name="pass" type="password" placeholder="Parola">
            <button>Login</button>
        </form>
    </div>
    </html>
    """

# INDEX
@app.route("/")
def index():
    if not session.get("login"):
        return redirect("/login")

    wb = load_workbook(FILE)
    ws = wb.active

    rows = ""
    for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        rows += f"""
        <tr>
            <td>{i}</td>
            <td>{r[0]}</td>
            <td>{r[1]}</td>
            <td>{r[2]}</td>
            <td>{r[3]}</td>
            <td>{r[4]}</td>
            <td><a href="/delete/{i}">🗑️</a></td>
        </tr>
        """

    return f"""
    <html>
    <style>
    body {{font-family: Arial; background:#f2f4f7;}}
    table {{width:95%;margin:auto;border-collapse:collapse;}}
    th {{background:#6200EE;color:white;padding:10px;}}
    td {{padding:8px;border-bottom:1px solid #ccc;text-align:center;}}
    tr:nth-child(even) {{background:#f9f9f9;}}
    .btn {{padding:10px 20px;background:#03DAC5;border:none;color:black;cursor:pointer;}}
    </style>

    <h2 style="text-align:center;">Raport curse Flota Comteleprest</h2>

    <div style="text-align:center;margin:20px;">
        <a href="/download"><button class="btn">⬇ Export Excel</button></a>
    </div>

    <table>
        <tr>
            <th>ID</th>
            <th>Nr Auto</th>
            <th>Data</th>
            <th>Locatie Plecare</th>
            <th>Locatie Sosire</th>
            <th>Km Parcurs</th>
            <th>Actiune</th>
        </tr>
        {rows}
    </table>
    </html>
    """

# ADAUGARE CURSA (FIX IMPORTANT AICI)
@app.route("/api/adauga", methods=["POST"])
def add():
    try:
        data = request.json

        nr_auto = data.get("nr_auto") or data.get("nrAuto")
        data_cursa = data.get("data")
        plecare = data.get("locatie_plecare") or data.get("locPlecare")
        sosire = data.get("locatie_sosire") or data.get("locSosire")
        km = data.get("km_parcurs") or data.get("km")

        wb = load_workbook(FILE)
        ws = wb.active

        ws.append([
            nr_auto,
            data_cursa,
            plecare,
            sosire,
            km
        ])

        wb.save(FILE)

        return jsonify({"status": "ok"})

    except Exception as e:
        return jsonify({"error": str(e)})

# STERGERE
@app.route("/delete/<int:id>")
def delete(id):
    wb = load_workbook(FILE)
    ws = wb.active

    ws.delete_rows(id + 2)
    wb.save(FILE)

    return redirect("/")

# EXPORT EXCEL PRO
@app.route("/download")
def download():
    wb = load_workbook(FILE)
    ws = wb.active

    header_fill = PatternFill(start_color="6200EE", end_color="6200EE", fill_type="solid")
    bold = Font(bold=True, color="FFFFFF")
    center = Alignment(horizontal="center")

    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = bold
        cell.alignment = center

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = center
            cell.border = border

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 40
    ws.column_dimensions["E"].width = 15

    wb.save(FILE)

    return send_file(FILE, as_attachment=True)

if __name__ == "__main__":
    app.run()
