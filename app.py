from flask import Flask, request, send_file
from flask_cors import CORS
import pandas as pd
import camelot
import tempfile
import os

app = Flask(__name__)
CORS(app)  # allow requests from frontend

@app.route('/convert', methods=['POST'])
def convert_pdf_to_excel():
    file = request.files.get('file')
    pages = request.form.get('pages', '').strip()  # get pages input; default '' -> means all pages

    if not file:
        return "No file uploaded", 400

    # default to 'all' if pages input is empty
    pages_to_parse = pages if pages else 'all'

    print(f"Received file: {file.filename}")
    print(f"Parsing pages: {pages_to_parse}")

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = os.path.join(tmpdir, file.filename)
            file.save(pdf_path)

            # parse tables
            tables = camelot.read_pdf(pdf_path, pages=pages_to_parse)

            if not tables:
                return "No tables found", 400

            excel_path = os.path.join(tmpdir, 'output.xlsx')
            with pd.ExcelWriter(excel_path) as writer:
                for i, table in enumerate(tables):
                    table.df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)

            return send_file(excel_path, as_attachment=True, download_name='converted.xlsx')

    except Exception as e:
        print(f"Error: {e}")
        return f"Error processing PDF: {str(e)}", 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
