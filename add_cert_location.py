import pandas as pd

# Read original CSV (DO NOT modify it)
df = pd.read_csv("event_data/linebot_data.csv") 

# Create CertificateFile column safely
df["CertificateFile"] = (
    df["Name2"]
    .str.strip()
    .str.replace(r"[^\w\s]", "", regex=True)  # remove special characters
    .str.replace(" ", "_", regex=False)       # replace spaces
    + ".pdf"
)

# Select only required columns
output_df = df[["Name", "Email", "CertificateFile"]]

# Save to a NEW CSV
output_df.to_csv("mail_req_csv/linebot_mail_data.csv", index=False)

print("linebot_mail_data.csv created successfully")
print(output_df.head())
