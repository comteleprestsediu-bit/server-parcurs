from io import BytesIO

@app.route("/download")
def download_excel():

    wb = Workbook()
    ws = wb.active

    # TITLU
    ws.merge_cells("A1:F1")
    ws["A1"] = "Raport curse Flota Comteleprest"
    ws["A1"].font = Font(size=16, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")

    # HEADER
    headers = ["ID", "Nr Auto", "Data", "Locatie Plecare", "Locatie Sosire", "Km Parcurs"]

    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # DATE
    curse = citeste_curse()

    for row_index, c in enumerate(curse, start=3):
        ws.cell(row=row_index, column=1, value=c["id"])
        ws.cell(row=row_index, column=2, value=c["nr_auto"])
        ws.cell(row=row_index, column=3, value=c["data"])
        ws.cell(row=row_index, column=4, value=c["locatie_plecare"])
        ws.cell(row=row_index, column=5, value=c["locatie_sosire"])
        ws.cell(row=row_index, column=6, value=c["km_parcurs"])

    # AUTO WIDTH
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

    # BORDER
    thin = Side(style="thin")
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=6):
        for cell in row:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

    # 🔥 SALVARE ÎN MEMORIE (NU PE DISC)
    file_stream = BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name="Raport_Flota_Comteleprest.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
