from flask import Flask, request, send_file
import pandas as pd
import camelot
import tempfile
import os

app = Flask(__name__)

@app.route('/convert', methods=['POST'])
def convert_pdf_to_excel():
    file = request.files['file']
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = os.path.join(tmpdir, file.filename)
        file.save(pdf_path)

        # Extract tables using camelot
        tables = camelot.read_pdf(pdf_path, pages='all')

        # Save tables into one Excel file
        excel_path = os.path.join(tmpdir, 'output.xlsx')
        with pd.ExcelWriter(excel_path) as writer:
            for i, table in enumerate(tables):
                table.df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)

        return send_file(excel_path, as_attachment=True, download_name='converted.xlsx')

if __name__ == '__main__':
    app.run(port=5000, debug=True)
