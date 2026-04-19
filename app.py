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

        fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        font = Font(bold=True, color="FFFFFF")

        for col in range(1, len(headers)+1):
            c = ws.cell(row=2, column=col)
            c.fill = fill
            c.font = font
            c.alignment = Alignment(horizontal="center")

        widths = [15, 15, 12, 40, 12, 40, 15]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64+i)].width = w

        wb.save(FILE)

init_excel()


# 🔹 CITIRE CURSE (fără TOTAL)
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
            "km_plecare": str(row[2] or ""),
            "locatie_plecare": row[3] or "",
            "km_sosire": str(row[4] or ""),
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
                <input name="user" class="form-control mb-2" placeholder="User">
                <input name="parola" type="password" class="form-control mb-2" placeholder="Parola">
                <button class="btn btn-primary w-100">Login</button>
            </form>
        </div>
    </div>
    </body>
    </html>
    """


# 🔓 LOGOUT
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


# 🔥 PAGINA PRINCIPALĂ
@app.route("/")
def index():
    if "user" not in session:
        return redirect("/login")

    curse = citeste_curse()

    html = """
    <html>
    <head>
        <title>Raport curse</title>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    </head>

    <body class="bg-light">
    <div class="container mt-4">
    <div class="card p-4 shadow">

    <div class="d-flex justify-content-between">
        <h3>🚗 Raport curse</h3>
        <a href="/logout" class="btn btn-secondary">Logout</a>
    </div>

    <a href="/download" class="btn btn-success mt-2 mb-3">⬇️ Descarcă Excel</a>

    <input id="search" class="form-control mb-3" placeholder="Caută...">

    <table class="table table-striped table-hover" id="tabel">
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
            <th>Șterge</th>
        </tr>
        </thead>
        <tbody>
    """

    for c in curse:
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
            <td><button onclick="sterge({c['id']})" class="btn btn-danger btn-sm">🗑</button></td>
        </tr>
        """

    html += """
        </tbody>
    </table>

    <h5 id="total"></h5>

    </div>
    </div>

    <script>
    document.getElementById("search").addEventListener("keyup", function(){
        let f = this.value.toLowerCase();
        document.querySelectorAll("#tabel tbody tr").forEach(r=>{
            r.style.display = r.innerText.toLowerCase().includes(f) ? "" : "none";
        });
    });

    let total = 0;
    document.querySelectorAll("#tabel tbody tr").forEach(r=>{
        let km = parseFloat(r.cells[7].innerText) || 0;
        total += km;
    });
    document.getElementById("total").innerText = "Total km: " + total;

    function sterge(id){
        fetch("/sterge_cursa/"+id,{method:"DELETE"})
        .then(()=>location.reload());
    }
    </script>

    </body>
    </html>
    """

    return html


# 🔹 DOWNLOAD
@app.route("/download")
def download_excel():
    return send_file(FILE, as_attachment=True)


# 🔹 API LISTĂ (FIX IMPORTANT)
@app.route("/api/curse", methods=["GET"])
def api_curse():
    return jsonify(citeste_curse())


# 🔹 ADAUGARE CURSĂ (FIX TELEFON)
@app.route("/adauga_cursa", methods=["POST"])
def adauga_cursa():
    try:
        data = request.get_json(silent=True)
        if not data:
            data = request.form

        nr_auto = data.get("nrAuto", "")
        data_cursa = data.get("data", "")
        plecare = data.get("locatiePlecare", "")
        sosire = data.get("locatieSosire", "")
        km_plecare = data.get("kmPlecare", "")
        km_sosire = data.get("kmSosire", "")
        km_parcurs = data.get("kmParcurs", "")

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

        wb.save(FILE)

        return {"status": "ok"}

    except Exception as e:
        print("EROARE:", e)
        return {"status": "error"}


# 🔥 ȘTERGERE
@app.route("/sterge_cursa/<int:id>", methods=["DELETE"])
def sterge_cursa(id):
    wb = load_workbook(FILE)
    ws = wb.active
    ws.delete_rows(id + 3)
    wb.save(FILE)
    return {"status": "deleted"}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
