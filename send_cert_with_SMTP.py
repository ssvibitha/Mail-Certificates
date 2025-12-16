import pandas as pd
import smtplib
from email.message import EmailMessage
import ssl
import os
from pathlib import Path
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

import sys

# Config
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

SENDER_EMAIL = "robotics@snuchennai.edu.in" #Change it with your email
APP_PASSWORD = "INSERT PASSWD" #Change it with your email passwd

CSV_PATH = "mail_req_csv/linebot_mail_data.csv" #Update with req csv file
TEMPLATE_PATH = "mail_template.txt"
FAILED_CSV = "failed_mails.csv"

CERT_FOLDER = Path(BASE_DIR) / "linebot_certificates"
SUBJECT = "Participation Certificate - Line Following Bot Workshop" # Subject
EVENT_NAME = "Line Following Bot Workshop"
# Pre-flight Checks
if not os.path.exists(CSV_PATH):
    print(f"Error: CSV file not found at {CSV_PATH}")
    sys.exit(1)

if not os.path.exists(TEMPLATE_PATH):
    print(f"Error: Template file not found at {TEMPLATE_PATH}")
    sys.exit(1)

# Load students data
try:
    df = pd.read_csv(CSV_PATH)
except Exception as e:
    print(f"Error reading CSV file: {e}")
    sys.exit(1)


# Validating all certificates exist before sending
print("Checking for certificate files...")
missing_certs = []
for _, row in df.iterrows():
    cert_file = row["CertificateFile"]
    # Handle potential NaN or non-string values
    if pd.isna(cert_file): 
        missing_certs.append(f"Row {_}: Missing CertificateFile name")
        continue
        
    cert_path = CERT_FOLDER / str(cert_file)
    if not os.path.exists(cert_path):
        missing_certs.append(cert_file)

if missing_certs:
    print(f"Error: Found {len(missing_certs)} missing certificate files:")
    for cert in missing_certs:
        print(f"- {cert}")
    sys.exit(1)
print("All certificates present.")

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

    email_count = 0
    for _, row in df.iterrows():
        if email_count >= 1:
            print("Breaking at 1")
            break
        # if email_count < 1:
        #    email_count = email_count+1
        #    continue

        name = row["Name"]
        recipient = row["Email"]
        cert_file = row["CertificateFile"]
        cert_path = CERT_FOLDER / cert_file

        try:
            # Check for certificate existence
            if not os.path.exists(cert_path):
                # Raise error to trigger the except block and log to CSV
                raise FileNotFoundError(f"Certificate file not found: {cert_file}")

            msg = EmailMessage()
            msg["From"] = SENDER_EMAIL
            msg["To"] = recipient
            msg["Subject"] = SUBJECT 

            # Personalize message
            body = email_template.replace("{name}", name).replace("{EVENT_NAME}",EVENT_NAME)
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
            email_count += 1
            # break #Uncomment to send to only one participant
        
        except smtplib.SMTPException as e:
            failed_records.append({
                "Name": name,
                "Email": recipient,
                "CertificateFile": cert_file,
                "Reason": f"SMTP Error: {str(e)}",
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
            print(f"SMTP Failed for {recipient}: {e}")
            
        except Exception as e:
            failed_records.append({
                "Name": name,
                "Email": recipient,
                "CertificateFile": cert_file,
                "Reason": f"Error: {str(e)}",
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
