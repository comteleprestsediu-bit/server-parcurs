from flask import Flask, request, jsonify, send_file, redirect
from openpyxl import Workbook, load_workbook
import os

app = Flask(__name__)

FILE = "curse.xlsx"

# 🔹 INIT EXCEL
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

# 🔹 CITEȘTE DATE
def citeste_curse():
    wb = load_workbook(FILE)
    ws = wb.active

    curse = []

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue

        curse.append({
            "id": i - 1,
            "nr_auto": row[0] or "",
            "data": row[1] or "",
            "locatie_plecare": row[2] or "",
            "locatie_sosire": row[3] or "",
            "km_parcurs": str(row[4] or "")
        })

    return curse

# 🔹 PAGINA PRINCIPALĂ
@app.route("/")
def index():
    curse = citeste_curse()

    html = """
    <h2 style='text-align:center'>Raport curse Flota Comteleprest</h2>

    <div style='text-align:center; margin-bottom:20px;'>
        <a href='/download'>
            <button style='padding:10px 20px;'>⬇ Export Excel</button>
        </a>
    </div>

    <table border='1' cellpadding='8' style='width:100%; border-collapse:collapse;'>
    <tr style='background:#ddd'>
        <th>ID</th>
        <th>Nr Auto</th>
        <th>Data</th>
        <th>Locatie Plecare</th>
        <th>Locatie Sosire</th>
        <th>Km Parcurs</th>
    </tr>
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

    html += "</table>"

    return html

# 🔹 API LISTĂ
@app.route("/api/curse")
def api_curse():
    return jsonify(citeste_curse())

# 🔹 ADAUGĂ CURSĂ
@app.route("/adauga_cursa", methods=["POST"])
def adauga_cursa():
    try:
        data = request.get_json(force=True)

        wb = load_workbook(FILE)
        ws = wb.active

        ws.append([
            data.get("nrAuto"),
            data.get("data"),
            data.get("locatiePlecare"),
            data.get("locatieSosire"),
            data.get("kmParcurs")
        ])

        wb.save(FILE)

        return {"status": "ok"}

    except Exception as e:
        print("EROARE:", e)
        return {"status": "error"}

# 🔹 DOWNLOAD EXCEL (FIXAT 100%)
@app.route("/download")
def download():
    try:
        if not os.path.exists(FILE):
            return "Fisierul nu exista", 404

        return send_file(
            FILE,
            as_attachment=True,
            download_name="Raport_Curse_Comteleprest.xlsx"
        )

    except Exception as e:
        return f"Eroare download: {e}"

# 🔹 START SERVER
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
