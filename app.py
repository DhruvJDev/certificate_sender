from flask import Flask, render_template, request
import os
from email_sender import send_bulk_emails

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Create uploads folder if not exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/send', methods=['POST'])
def send():
    excel = request.files['excel']
    pdfs = request.files.getlist('pdfs')
    message = request.form['message']

    # Save Excel
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel.filename)
    excel.save(excel_path)

    # Save PDFs
    for pdf in pdfs:
        pdf.save(os.path.join(app.config['UPLOAD_FOLDER'], pdf.filename))

    # Send Emails
    success, failed = send_bulk_emails(excel_path, app.config['UPLOAD_FOLDER'], message)

    return render_template('index.html', success=success, failed=failed)

if __name__ == "__main__":
    app.run(debug=True)