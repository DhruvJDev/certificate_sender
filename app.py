from flask import Flask, render_template, request, jsonify
import os
from io import BytesIO
from email_sender import send_bulk_emails
import pandas as pd

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Create uploads folder if not exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/preview-excel', methods=['POST'])
def preview_excel():
    excel = request.files.get('excel')

    if not excel or not excel.filename:
        return jsonify({'error': 'No Excel file provided'}), 400

    try:
        file_name = excel.filename.lower()
        engine = 'xlrd' if file_name.endswith('.xls') else 'openpyxl'
        excel_bytes = BytesIO(excel.read())
        dataframe = pd.read_excel(excel_bytes, engine=engine).fillna('')

        return jsonify({
            'headers': dataframe.columns.tolist(),
            'rows': dataframe.to_dict(orient='records')
        })
    except Exception as error:
        return jsonify({'error': str(error)}), 400

@app.route('/send', methods=['POST'])
def send():
    excel = request.files['excel']
    pdfs = request.files.getlist('pdfs')
    message = request.form['message']
    subject = request.form.get('subject', 'Your Certificate')
    cc = request.form.get('cc', '').strip()
    bcc = request.form.get('bcc', '').strip()

    # Save Excel
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel.filename)
    excel.save(excel_path)

    # Save PDFs
    for pdf in pdfs:
        pdf.save(os.path.join(app.config['UPLOAD_FOLDER'], pdf.filename))

    # Send Emails
    success, failed = send_bulk_emails(excel_path, app.config['UPLOAD_FOLDER'], message, subject, cc, bcc)

    return render_template('index.html', success=success, failed=failed)

if __name__ == "__main__":
    app.run(debug=True)