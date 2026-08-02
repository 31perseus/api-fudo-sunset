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

        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('Ingredientes')

        for row_idx, row_data in enumerate(data):
            for col_idx, cell_value in enumerate(row_data):
                ws.write(row_idx, col_idx, cell_value)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name='stock.xls',
            mimetype='application/vnd.ms-excel'
        )
    except Exception as e:
        return str(e), 500
