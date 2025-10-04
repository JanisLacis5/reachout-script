import os
from typing import Tuple, List
from itertools import zip_longest

from dotenv import load_dotenv; load_dotenv()
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# If SCOPES is modified, delete token.json
SCOPES = os.getenv("SCOPES")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
RANGE_NAME = "Sheet1"

def auth():
    if not os.path.exists("credentials.json"):
        raise Exception("No credentials.josn found")

    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    creds = None
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
    return creds

def get_sheet_lines(creds) -> Tuple[List[str], List[List[str]]]:
    """
    Function that reads the contents of the Google Sheet and returns them
    in a format tuple(sheet columns, list of sheet rows)
    """
    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get("values", [])

    return values[0], values[1:]

creds = auth()
cols, rows = get_sheet_lines(creds)

from pprint import pprint
for val in rows:
    pprint(dict(zip_longest(cols, val)))
    
