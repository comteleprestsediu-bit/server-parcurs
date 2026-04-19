from flask import Flask, request, jsonify, send_file
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
from io import BytesIO

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

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


# 🔹 PAGINA WEB
@app.route("/")
def index():
    curse = citeste_curse()

    html = """
    <h2 style='text-align:center;'>Raport curse Flota Comteleprest</h2>
    <div style='text-align:center;margin:10px;'>
        <a href='/download'>⬇️ Excel</a> |
        <a href='/download-pdf'>⬇️ PDF</a>
    </div>
    <table border='1' cellpadding='6' style='border-collapse:collapse;margin:auto;'>
    <tr style='background:#4F81BD;color:white;'>
        <th>ID</th>
        <th>Nr Auto</th>
        <th>Data</th>
        <th>Locatie Plecare</th>
        <th>Locatie Sosire</th>
        <th>Km Parcurs</th>
        <th>Actiuni</th>
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
            <td><a href="/sterge_cursa/{c['id']}">❌ Șterge</a></td>
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
        print("❌ Eroare:", e)
        return {"status": "error"}


# 🔥 ȘTERGE CURSĂ
@app.route("/sterge_cursa/<int:id>")
def sterge_cursa(id):
    try:
        wb = load_workbook(FILE)
        ws = wb.active

        ws.delete_rows(id + 2)

        wb.save(FILE)

        return "<script>window.location.href='/'</script>"

    except Exception as e:
        print("❌ Eroare ștergere:", e)
        return "error"


# 🔥 EXPORT EXCEL (FĂRĂ 500 ERROR)
@app.route("/download")
def download_excel():

    wb = Workbook()
    ws = wb.active

    ws.merge_cells("A1:F1")
    ws["A1"] = "Raport curse Flota Comteleprest"
    ws["A1"].font = Font(size=16, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")

    headers = ["ID", "Nr Auto", "Data", "Locatie Plecare", "Locatie Sosire", "Km Parcurs"]

    fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")

    for row_index, c in enumerate(citeste_curse(), start=3):
        ws.cell(row=row_index, column=1, value=c["id"])
        ws.cell(row=row_index, column=2, value=c["nr_auto"])
        ws.cell(row=row_index, column=3, value=c["data"])
        ws.cell(row=row_index, column=4, value=c["locatie_plecare"])
        ws.cell(row=row_index, column=5, value=c["locatie_sosire"])
        ws.cell(row=row_index, column=6, value=c["km_parcurs"])

    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_len + 2

    thin = Side(style="thin")
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=6):
        for cell in row:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)

    return send_file(
        stream,
        as_attachment=True,
        download_name="Raport_Flota_Comteleprest.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# 🔥 EXPORT PDF
@app.route("/download-pdf")
def download_pdf():

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)

    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("Raport curse Flota Comteleprest", styles["Title"]))
    elements.append(Spacer(1, 20))

    data = [["ID", "Nr Auto", "Data", "Plecare", "Sosire", "Km"]]

    for c in citeste_curse():
        data.append([
            c["id"],
            c["nr_auto"],
            c["data"],
            c["locatie_plecare"],
            c["locatie_sosire"],
            c["km_parcurs"]
        ])

    table = Table(data)

    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.grey),
        ("TEXTCOLOR",(0,0),(-1,0),colors.whitesmoke),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("GRID", (0,0), (-1,-1), 1, colors.black)
    ]))

    elements.append(table)

    doc.build(elements)

    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="Raport_Flota_Comteleprest.pdf",
        mimetype="application/pdf"
    )


# 🔹 START SERVER
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
