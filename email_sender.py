import pandas as pd
import smtplib
from email.message import EmailMessage
import os
from config import EMAIL, PASSWORD

def send_bulk_emails(excel_path, pdf_folder, message_template, subject_template="Your Certificate", cc="", bcc=""):
    df = pd.read_excel(excel_path)

    success = []
    failed = []
    
    # Parse CC and BCC addresses
    cc_list = [email.strip() for email in cc.split(',') if email.strip()] if cc else []
    bcc_list = [email.strip() for email in bcc.split(',') if email.strip()] if bcc else []

    for index, row in df.iterrows():
        try:
            name = row['Name']
            receiver = row['Email']
            file_name = row['FileName']

            msg = EmailMessage()
            # Personalize subject
            subject = subject_template.replace("{name}", name)
            msg['Subject'] = subject
            msg['From'] = EMAIL
            msg['To'] = receiver
            
            # Add CC recipients
            if cc_list:
                msg['Cc'] = ', '.join(cc_list)
            
            # Add BCC recipients
            if bcc_list:
                msg['Bcc'] = ', '.join(bcc_list)

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