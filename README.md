# Bulk Email Certificate Sender

A Flask web app to send personalized emails with PDF attachments in bulk.

The app reads recipient data from an Excel file and matches each row to a PDF file in the uploads folder.

## Features

- Upload one Excel file with recipient details
- Upload multiple PDF files in one go
- Personalize email message using `{name}` placeholder
- Send emails through Gmail SMTP
- Show success and failure results after processing

## Project Structure

```text
certificate_sender/
|-- app.py
|-- email_sender.py
|-- config.py
|-- requirements.txt
|-- .env.example
|-- .env                # local only, not committed
|-- uploads/
|-- templates/
|   `-- index.html
`-- static/
    `-- style.css
```

## Requirements

- Python 3.9+
- Gmail account with App Password enabled

## Setup

1. Clone the repository.
2. Create and activate a virtual environment.
3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Create environment file from sample:

```bash
cp .env.example .env
```

5. Update `.env` values:

```env
EMAIL=your_email@gmail.com
PASSWORD=your_app_password
```

## Excel File Format

Your Excel file must include these exact column names:

- `Name`
- `Email`
- `FileName`

Example:

| Name | Email            | FileName |
| ---- | ---------------- | -------- |
| Alex | alex@example.com | alex.pdf |
| Sam  | sam@example.com  | sam.pdf  |

`FileName` must match an uploaded PDF filename exactly.

## Run the App

```bash
python app.py
```

Open in browser:

- http://127.0.0.1:5000

## How to Use

1. Upload the Excel file.
2. Upload all required PDF files.
3. Enter email message text (use `{name}` for personalization).
4. Click Send Emails.
5. Review success and failed email lists on the page.

## Notes

- Use Gmail App Password, not your main Gmail password.
- For production use, disable Flask debug mode.
- Clean old files from `uploads/` regularly.
