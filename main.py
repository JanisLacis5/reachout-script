import os
from itertools import zip_longest

import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = "1Rzhie35_kv7LaBpfiky1iqxzRqxz5E0MljoLQlXj_Ts"
RANGE_NAME = "Sheet1"

creds = None
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)

# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)

# Save the credentials for the next run
with open("token.json", "w") as token:
    token.write(creds.to_json())

service = build("sheets", "v4", credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
values = result.get("values", [])

# Make a dataframe of values for easier work
df_cols = values[0]
df_vals = [dict(zip_longest(df_cols, val)) for val in values[1:]]
df = pd.DataFrame(df_vals, columns=df_cols)
