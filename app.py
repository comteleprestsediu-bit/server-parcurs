from flask import Flask, request, jsonify
from openpyxl import Workbook, load_workbook
import os

app = Flask(__name__)

FILE = "curse.xlsx"

# ============================
# 🔹 INIT EXCEL
# ============================
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


# ============================
# 🔹 CITEȘTE DATE
# ============================
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


# ============================
# 🔹 PAGINA PRINCIPALĂ
# ============================
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


# ============================
# 🔹 API LISTĂ
# ============================
@app.route("/api/curse")
def api_curse():
    return jsonify(citeste_curse())


# ============================
# 🔹 ADAUGĂ CURSĂ
# ============================
@app.route("/adauga_cursa", methods=["POST"])
def adauga_cursa():
    try:
        data = request.get_json(force=True)

        print("📥 DATA PRIMITA:", data)  # 🔥 DEBUG IMPORTANT

        # 🔴 VALIDARE
        if not data:
            return {"status": "error", "msg": "No data"}, 400

        nr = data.get("nrAuto", "")
        data_c = data.get("data", "")
        lp = data.get("locatiePlecare", "")
        ls = data.get("locatieSosire", "")
        km = data.get("kmParcurs", "0")

        if nr == "" or lp == "" or ls == "":
            return {"status": "error", "msg": "Campuri lipsa"}, 400

        wb = load_workbook(FILE)
        ws = wb.active

        ws.append([nr, data_c, lp, ls, km])

        wb.save(FILE)

        print("✅ SALVAT CU SUCCES")

        return {"status": "ok"}

    except Exception as e:
        print("❌ EROARE:", e)
        return {"status": "error"}, 500


# ============================
# 🔥 ȘTERGE CURSĂ
# ============================
@app.route("/sterge_cursa/<int:id>", methods=["DELETE"])
def sterge_cursa(id):
    try:
        wb = load_workbook(FILE)
        ws = wb.active

        ws.delete_rows(id + 2)

        wb.save(FILE)

        return {"status": "deleted"}

    except Exception as e:
        print("❌ EROARE STERGERE:", e)
        return {"status": "error"}


# ============================
# 🔹 TOTAL KM
# ============================
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


# ============================
# 🔹 START SERVER
# ============================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
