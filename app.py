from flask import Flask, request, jsonify, send_file, redirect, session
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
import os

app = Flask(__name__)
app.secret_key = "parcurs_secret"

FILE = "curse.xlsx"

# LOGIN DATE
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

# LOGIN PAGE
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = request.form.get("user")
        password = request.form.get("pass")

        if user == USER and password == PASS:
            session["login"] = True
            return redirect("/")
        else:
            return "Login greșit"

    return '''
    <h2>Login</h2>
    <form method="post">
        User: <input name="user"><br><br>
        Parola: <input name="pass" type="password"><br><br>
        <button type="submit">Login</button>
    </form>
    '''

# PAGINA PRINCIPALA
@app.route("/")
def index():
    if not session.get("login"):
        return redirect("/login")

    wb = load_workbook(FILE)
    ws = wb.active

    rows = ""
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        rows += f"""
        <tr>
            <td>{i}</td>
            <td>{row[0]}</td>
            <td>{row[1]}</td>
            <td>{row[2]}</td>
            <td>{row[3]}</td>
            <td>{row[4]}</td>
            <td>
                <a href="/delete/{i}">Șterge</a>
            </td>
        </tr>
        """

    return f"""
    <h2 style="text-align:center;">Raport curse Flota Comteleprest</h2>

    <div style="text-align:center; margin:20px;">
        <a href="/download"><button>⬇ Export Excel</button></a>
    </div>

    <table border="1" style="width:100%; border-collapse: collapse;">
        <tr style="background:#ddd;">
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
    """

# ADAUGARE CURSA (ANDROID)
@app.route("/api/adauga", methods=["POST"])
def adauga():
    data = request.json

    wb = load_workbook(FILE)
    ws = wb.active

    ws.append([
        data.get("nr_auto"),
        data.get("data"),
        data.get("locatie_plecare"),
        data.get("locatie_sosire"),
        data.get("km_parcurs")
    ])

    wb.save(FILE)

    return jsonify({"status": "ok"})

# LISTA API
@app.route("/api/curse")
def lista():
    wb = load_workbook(FILE)
    ws = wb.active

    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        data.append({
            "nr_auto": row[0],
            "data": row[1],
            "locatie_plecare": row[2],
            "locatie_sosire": row[3],
            "km_parcurs": row[4]
        })

    return jsonify(data)

# STERGERE
@app.route("/delete/<int:id>")
def delete(id):
    wb = load_workbook(FILE)
    ws = wb.active

    ws.delete_rows(id + 2)

    wb.save(FILE)
    return redirect("/")

# EXPORT EXCEL PROFESIONAL
@app.route("/download")
def download():
    wb = load_workbook(FILE)
    ws = wb.active

    bold = Font(bold=True)
    center = Alignment(horizontal="center")
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = center

    for cell in ws[1]:
        cell.font = bold

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 40
    ws.column_dimensions["E"].width = 15

    wb.save(FILE)

    return send_file(FILE, as_attachment=True)

if __name__ == "__main__":
    app.run()
