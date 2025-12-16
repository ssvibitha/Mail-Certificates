import pandas as pd
import smtplib
from email.message import EmailMessage
import ssl
import os
from pathlib import Path

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Config
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

SENDER_EMAIL = "EMAIL ADDRESS" #Change it with your email
APP_PASSWORD = "YOUR EMAIL PASSWD" #Change it with your email passwd

CSV_PATH = "sensorverse_mail_data.csv"
TEMPLATE_PATH = "template.txt"
FAILED_CSV = "failed_mails.csv"

# Load students data
df = pd.read_csv(CSV_PATH)

# Load email template
with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
    email_template = f.read()

# Prepare failed mails list (fresh every run)
failed_records = []

# Secure SSL context
context = ssl.create_default_context()

with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
    server.starttls(context=context)
    server.login(SENDER_EMAIL, APP_PASSWORD)

    for _, row in df.iterrows():
        name = row["Name"]
        recipient = row["Email"]
        cert_file = row["CertificateFile"]
        CERT_FOLDER = Path(BASE_DIR) / "sensorverse_certificates"
        cert_path = CERT_FOLDER / cert_file

        try:
            if not os.path.exists(cert_path):
                print(f"File not found: {cert_path}")
                continue

            msg = EmailMessage()
            msg["From"] = SENDER_EMAIL
            msg["To"] = recipient
            msg["Subject"] = "Sample Mail" # Subject

            # Personalize message
            body = email_template.replace("{name}", name)
            msg.set_content(body)

            # Attach certificate
            with open(cert_path, "rb") as f:
                file_data = f.read()
                file_name = os.path.basename(cert_path)

            msg.add_attachment(
                file_data,
                maintype="application",
                subtype="pdf",
                filename=file_name
            )

            server.send_message(msg)
            print(f"Sent certificate to {recipient}")
            break #Comment to send to all participants
        
        except Exception as e:
            failed_records.append({
                "Name": name,
                "Email": recipient,
                "CertificateFile": cert_file,
                "Reason": str(e),
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
            print(f"Failed for {recipient}: {e}")

# Write failed mails to fresh CSV
if failed_records:
    pd.DataFrame(failed_records).to_csv(FAILED_CSV, index=False)
    print(f"Failed mails saved to {FAILED_CSV}")
else:
    print("No failed mails")

print("All emails processed.")
