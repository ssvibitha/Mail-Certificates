import pandas as pd
import smtplib
from email.message import EmailMessage
import ssl
from pathlib import Path
from datetime import datetime

# Config
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
SENDER_EMAIL = "EMAIL ADDRESS"
APP_PASSWORD = "YOUR EMAIL PASSWD"

FAILED_CSV = "failed_mails.csv"
CERT_FOLDER = Path("sensorverse_certificates")
TEMPLATE_PATH = "template.txt"

RESEND_FAILED_CSV = "failed_mails_retry.csv"
RESENT_SUCCESS_CSV = "resent_success.csv"

# Load data
df = pd.read_csv(FAILED_CSV)

with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
    email_template = f.read()

context = ssl.create_default_context()

still_failed = []
resent_success = []

with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
    server.starttls(context=context)
    server.login(SENDER_EMAIL, APP_PASSWORD)

    for _, row in df.iterrows():
        name = row["Name"]
        recipient = row["Email"]
        cert_file = row["CertificateFile"]

        try:
            cert_path = CERT_FOLDER / cert_file

            if not cert_path.exists():
                raise FileNotFoundError("Certificate file still missing")

            msg = EmailMessage()
            msg["From"] = SENDER_EMAIL
            msg["To"] = recipient
            msg["Subject"] = "Sample Mail"

            body = email_template.replace("{name}", name)
            msg.set_content(body)

            with open(cert_path, "rb") as f:
                msg.add_attachment(
                    f.read(),
                    maintype="application",
                    subtype="pdf",
                    filename=cert_file
                )

            server.send_message(msg)

            resent_success.append({
                "Name": name,
                "Email": recipient,
                "CertificateFile": cert_file,
                "ResentAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })

            print(f"Resent successfully → {recipient}")

        except Exception as e:
            still_failed.append({
                "Name": name,
                "Email": recipient,
                "CertificateFile": cert_file,
                "Reason": str(e),
                "RetryAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
            print(f"Still failed → {recipient}")

# Save results
if resent_success:
    pd.DataFrame(resent_success).to_csv(RESENT_SUCCESS_CSV, index=False)

if still_failed:
    pd.DataFrame(still_failed).to_csv(RESEND_FAILED_CSV, index=False)

print("Resend process completed.")
