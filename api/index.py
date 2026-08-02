from flask import Flask, request, send_file
import xlwt
import io

app = Flask(__name__)

@app.route('/api/convert', methods=['POST'])
def convert_to_xls():
    try:
        data = request.get_json()
        if not data:
            return "No se recibieron datos", 400

        # Crear el libro de trabajo de Excel (Formato antiguo BIFF8)
        wb = xlwt.Workbook(encoding='utf-8')

        # Iteramos sobre el diccionario que nos enviará Google Sheets
        # Ej: {"Productos": [[...]], "Ingredientes": [[...]]}
        if isinstance(data, dict):
            for sheet_name, sheet_data in data.items():
                ws = wb.add_sheet(sheet_name)
                for row_idx, row_data in enumerate(sheet_data):
                    for col_idx, cell_value in enumerate(row_data):
                        ws.write(row_idx, col_idx, cell_value)

        # Guardar el archivo virtualmente en la memoria RAM
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        # Enviar el archivo binario de vuelta
        return send_file(
            output,
            as_attachment=True,
            download_name='stock.xls',
            mimetype='application/vnd.ms-excel'
        )
    except Exception as e:
        return str(e), 500
