from flask import Flask, request, jsonify, send_file
from openpyxl import Workbook, load_workbook
import os

app = Flask(__name__)

FILE = "curse.xlsx"

def init_excel():
    if not os.path.exists(FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Raport Curse"
        ws.append(["Nr Auto", "Data", "Locatie Plecare", "Locatie Sosire", "Km Parcurs"])
        wb.save(FILE)

init_excel()

@app.route("/")
def index():
    wb = load_workbook(FILE)
    ws = wb.active

    rows = list(ws.values)

    html = """
    <h2>Raport Curse</h2>
    <table border=1>
    """

    for i, row in enumerate(rows):
        html += "<tr>"
        for cell in row:
            if i == 0:
                html += f"<th>{cell}</th>"
            else:
                html += f"<td>{cell}</td>"
        html += "</tr>"

    html += "</table>"
    return html

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

@app.route("/download")
def download():
    return send_file(FILE, as_attachment=True)

if __name__ == "__main__":
    app.run()
