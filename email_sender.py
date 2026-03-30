import pandas as pd
import smtplib
from email.message import EmailMessage
import os
from config import EMAIL, PASSWORD

def send_bulk_emails(excel_path, pdf_folder, message_template):
    df = pd.read_excel(excel_path)

    success = []
    failed = []

    for index, row in df.iterrows():
        try:
            name = row['Name']
            receiver = row['Email']
            file_name = row['FileName']

            msg = EmailMessage()
            msg['Subject'] = "Your Certificate"
            msg['From'] = EMAIL
            msg['To'] = receiver

            # Personalize message
            message = message_template.replace("{name}", name)
            msg.set_content(message)

            # Attach PDF
            file_path = os.path.join(pdf_folder, file_name)

            if not os.path.exists(file_path):
                failed.append(f"{receiver} (PDF not found)")
                continue

            with open(file_path, 'rb') as f:
                file_data = f.read()
                msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

            # Send email
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL, PASSWORD)
                smtp.send_message(msg)

            success.append(receiver)

        except Exception as e:
            failed.append(f"{receiver} ({str(e)})")

    return success, failed