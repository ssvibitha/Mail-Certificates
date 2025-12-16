# Automatically send emails from outlook desktop
# Without SMTP Protocol

import csv
import os
import time
import win32com.client as win32
from pathlib import Path

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Config
CSV_FILE = "mail_req_csv/sensorverse_mail_data.csv"
CERT_FOLDER = "sensorverse_certificates"
TEMPLATE_FILE = "mail_template.txt"
LOG_FILE = "send_log.txt"
FAILED_CSV = "failed_emails.csv"
EMAIL_SUBJECT = "Participation Certificate - SensorVerse"
DELAY_BETWEEN_EMAILS = 2  # in seconds, optional


# Load email template
with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
    TEMPLATE = f.read()

# Connect to Outlook
outlook = win32.Dispatch('Outlook.Application')

# Open Failed mail csv
if not os.path.exists(FAILED_CSV):
    with open(FAILED_CSV, "w", newline='', encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Name", "Email", "CertificateFile", "FailureReason"])

# Open log file
with open(LOG_FILE, "a", encoding="utf-8") as log_file:

    # Read CSV and send emails
    with open(CSV_FILE, newline='', encoding="utf-8") as csvfile:
        reader = csv.DictReader(csvfile)
        CERT_FOLDER = Path(BASE_DIR) / "sensorverse_certificates"

        for row in reader:
            name = row["Name"]
            email = row["Email"]
            cert_file = row["CertificateFile"]
            cert_path = CERT_FOLDER / cert_file

            if not cert_path.is_file():
                reason = "Certificate file not found"
                msg = f"Missing file: {cert_path}"
                print(msg)
                log_file.write(msg + "\n")

                with open(FAILED_CSV, "a", newline='', encoding="utf-8") as fail_file:
                    writer = csv.writer(fail_file)
                    writer.writerow([name, email, cert_file, reason])
                continue

            try:
                mail = outlook.CreateItem(0)
                mail.To = email
                mail.Subject = EMAIL_SUBJECT
                mail.Body = TEMPLATE.format(name=name)

                mail.Attachments.Add(str(cert_path.resolve()))
                mail.Send()

                msg = f"Sent certificate to {name} ({email})"
                print(msg)
                log_file.write(msg + "\n")

                time.sleep(DELAY_BETWEEN_EMAILS)

            except Exception as e:
                reason = str(e)
                msg = f"Failed to send to {name} ({email}): {e}"
                print(msg)
                log_file.write(msg + "\n")

                with open(FAILED_CSV, "a", newline='', encoding="utf-8") as fail_file:
                    writer = csv.writer(fail_file)
                    writer.writerow([name, email, cert_file, reason])

print("All emails processed")