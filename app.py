from flask import Flask, request, jsonify
from openpyxl import Workbook, load_workbook
import os

app = Flask(__name__)

FILE = "curse.xlsx"

# 🔹 INIT EXCEL
def init_excel():
    if not os.path.exists(FILE):
        wb = Workbook()
        ws = wb.active
        ws.append([
            "nr_auto",
            "data",
            "locatie_plecare",
            "locatie_sosire",
            "km_parcurs"
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
            continue  # skip header

        curse.append({
            "id": i - 1,  # 🔥 IMPORTANT pentru swipe delete
            "nr_auto": row[0] or "",
            "data": row[1] or "",
            "locatie_plecare": row[2] or "",
            "locatie_sosire": row[3] or "",
            "km_parcurs": str(row[4] or "0")
        })

    return curse


# 🔹 PAGINA PRINCIPALĂ (browser)
@app.route("/")
def index():
    curse = citeste_curse()

    html = "<h2>Foaie Parcurs</h2><table border='1' cellpadding='5'>"

    html += """
    <tr>
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


# 🔹 API LISTĂ CURSE
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
        print("❌ Eroare:", e)
        return {"status": "error"}


# 🔥 ȘTERGE CURSĂ (SWIPE)
@app.route("/sterge_cursa/<int:id>", methods=["DELETE"])
def sterge_cursa(id):
    try:
        wb = load_workbook(FILE)
        ws = wb.active

        # +2 pentru că Excel începe de la 1 + header
        ws.delete_rows(id + 2)

        wb.save(FILE)

        return {"status": "deleted"}

    except Exception as e:
        print("❌ Eroare ștergere:", e)
        return {"status": "error"}


# 🔹 TOTAL KM PE MAȘINĂ
@app.route("/api/total")
def total_km():
    curse = citeste_curse()
    rezultat = {}

    for c in curse:
        nr = c["nr_auto"]

        try:
            km = float(str(c["km_parcurs"]).replace(",", "."))
        except:
            km = 0

        if nr not in rezultat:
            rezultat[nr] = 0

        rezultat[nr] += km

    return jsonify(rezultat)


# 🔹 START SERVER
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)