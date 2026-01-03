import gspread
import os
import pandas as pd
from google.oauth2.service_account import Credentials
from google.oauth2 import service_account
from dotenv import load_dotenv
from datetime import datetime
import pytz

# Load environment variables
load_dotenv()
ist = pytz.timezone('Asia/Kolkata')
now_ist = datetime.now(ist)

with open("credentials.json", "w") as f:
    f.write(os.getenv("GOOGLE_SHEET_CONNECTOR"))

scope = ["https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"]
SERVICE_ACCOUNT_FILE = "service_account.json"

creds = service_account.Credentials.from_service_account_file(
    "credentials.json",
    scopes=scope
)

gs_client = gspread.authorize(creds)

worksheet = gs_client.open("vista logs").worksheet("to be logged") #Change INput sheet name here

# Convert to DataFrame
data = worksheet.get_all_values()
df = pd.DataFrame(data)

print(df.head())
print(df.shape)
