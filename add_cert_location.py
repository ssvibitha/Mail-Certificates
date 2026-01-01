import pandas as pd
import re

# Read original CSV (DO NOT modify it)
df = pd.read_csv("event_data/linebot_data.csv") 

# Create CertificateFile column safely
df["CertificateFile"] = (
    re.sub(r'[^a-zA-Z0-9]+', '_', df["Name2"]).strip()     
    + ".pdf"
)

# Select only required columns
output_df = df[["Name", "Email", "CertificateFile"]]

# Save to a NEW CSV
output_df.to_csv("mail_req_csv/linebot_mail_data.csv", index=False)

print("linebot_mail_data.csv created successfully")
print(output_df.head())
