from flask import Flask, request, jsonify, send_file
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
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
            "km_parcurs": str(row[4] or "0")
        })

    return curse

# 🔹 PAGINA PRINCIPALĂ (WEB)
@app.route("/")
def index():
    curse = citeste_curse()

    html = """
    <h2 style='text-align:center;'>Raport curse Flota Comteleprest</h2>
    <div style='text-align:center;margin-bottom:10px;'>
        <a href="/download"><button>⬇️ Export Excel</button></a>
    </div>
    <table border='1' cellpadding='8' style='border-collapse:collapse;width:100%;'>
    """

    html += """
    <tr style='background:#ddd;font-weight:bold;'>
        <th>ID</th>
        <th>Nr Auto</th>
        <th>Data</th>
        <th>Locatie Plecare</th>
        <th>Locatie Sosire</th>
        <th>Km Parcurs</th>
        <th>Actiune</th>
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
            <td>
                <button onclick="sterge({c['id']})">🗑️ Șterge</button>
            </td>
        </tr>
        """

    html += """
    </table>

    <script>
    function sterge(id){
        fetch("/sterge_cursa/" + id, {method: "DELETE"})
        .then(() => location.reload());
    }
    </script>
    """

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

# 🔹 ȘTERGE
@app.route("/sterge_cursa/<int:id>", methods=["DELETE"])
def sterge_cursa(id):
    try:
        wb = load_workbook(FILE)
        ws = wb.active

        ws.delete_rows(id + 2)

        wb.save(FILE)

        return {"status": "deleted"}

    except Exception as e:
        print("EROARE:", e)
        return {"status": "error"}

# 🔹 EXPORT EXCEL PROFESIONAL
@app.route("/download")
def download_excel():
    try:
        wb = load_workbook(FILE)
        ws = wb.active

        # Titlu
        ws.insert_rows(1)
        ws["A1"] = "Raport curse Flota Comteleprest"
        ws.merge_cells("A1:E1")

        ws["A1"].font = Font(size=16, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center")

        # Header bold
        for cell in ws[2]:
            cell.font = Font(bold=True)

        # Lățime coloane
        ws.column_dimensions["A"].width = 15
        ws.column_dimensions["B"].width = 15
        ws.column_dimensions["C"].width = 30
        ws.column_dimensions["D"].width = 30
        ws.column_dimensions["E"].width = 15

        export_file = "raport_curse.xlsx"
        wb.save(export_file)

        return send_file(export_file, as_attachment=True)

    except Exception as e:
        print("EROARE DOWNLOAD:", e)
        return "Eroare export Excel", 500


# 🔹 START SERVER
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
