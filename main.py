import os
import base64
from datetime import datetime
from typing import Literal, Tuple, List
from itertools import zip_longest

from dotenv import load_dotenv
from email.message import EmailMessage
import google.auth
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

load_dotenv()

# If SCOPES is modified, delete token.json
SCOPES = ["https://www.googleapis.com/auth/gmail.send", "https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
GMAIL_MAIN = os.getenv("GMAIL_MAIN")
GMAIL_PASS = os.getenv("GMAIL_PASS")
RANGE_NAME = "Sheet1"


class EmailClient:
    def __init__(self, email_limit=10):
        self.creds = self._auth()
        cols, rows = self._get_sheet_lines()
        self.cols = cols
        self.rows = rows
        self.cols_mapping = {v: self._col_letter(i) for i, v in enumerate(cols)}
        self.email_limit = email_limit
        self.smtp_server = "smtp.gmail.com"
        self.smtp_port = 587

    def _col_letter(self, col: int, start_index=0):
        """
        Maps column number to excel column letters.
        Example (assuming start_index == 0, otherwise, all column numbers are shifted by start_index):
        0 -> A, 1 -> B, ..., 26 -> Z, 27 -> AA, ...
        """
        col += 1 - start_index
        letter = ""
        ascii_a = ord("A")

        while col > 0:
            letter += chr(ascii_a + ((col - 1) % 26))
            col = (col - 1) // 26

        return letter[::-1]

    def _auth(self):
        if not os.path.exists("credentials.json"):
            raise Exception("No credentials.josn found")

        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first time
        creds = None
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", [SCOPES])

        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES
                )
                creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())
        return creds

    def _get_sheet_lines(self) -> Tuple[List[str], List[List[str]]]:
        """
        Function that reads the contents of the Google Sheet and returns them
        in a format tuple(sheet columns, list of sheet rows)
        """
        service = build("sheets", "v4", credentials=self.creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = (
            sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
        )
        values = result.get("values", [])

        return values[0], values[1:]

    def _update_values(
        self,
        range: str,
        value_input_option: Literal["RAW", "USER_ENTERED"],
        values: List[List[str]],
    ):
        creds = self._auth()
        service = build("sheets", "v4", credentials=creds)

        body = {"values": values}
        result = (
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=SPREADSHEET_ID,
                range=range,
                body=body,
                valueInputOption=value_input_option,
            )
            .execute()
        )
        print(f"{(result.get('updatedCells'))} cells updated.")
        return result

    def _send_email(self, subject: str, row_no: int):
        row_list = self.rows[row_no]
        row = dict(zip_longest(self.cols, row_list))
        body = self._format_email_text(row_no)

        service = build("gmail", "v1", credentials=self.creds)
        message = EmailMessage()
        message.set_content(body)
        message["To"] = row['Contact email']
        message["From"] = GMAIL_MAIN
        message["Subject"] = subject

        # encoded message
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

        create_message = {"raw": encoded_message}
        send_message = (
            service.users()
            .messages()
            .send(userId="me", body=create_message)
            .execute()
        )
        print(f'Message Id: {send_message["id"]}')
        return send_message

    def _get_email_text(self, language: Literal["LV", "EN"], email_no: int):
        """
        Gets the necessary email text.
        params:
            - language: str - language in which the email will be sent
            - email_no: int - which email in the row this is
        """
        filename = f"{email_no}_{language}.txt"
        with open(filename, "r") as f:
            text = f.read()
        return text

    def _format_email_text(self, row_no: int):
        row_list = self.rows[row_no]
        row = dict(zip_longest(self.cols, row_list))

        language = row["Language"]
        email_no = int(row["Emails Sent"])
        text = self._get_email_text(language, email_no)
        text = text.format(name=row["Contact Name"])

        return text

    def _after_email(self, row_no: int):
        """
        Function that updates the Google Sheet after an email is sent. Specifically,
        it: adds 'Approached (Date)' column as the current date, adds +1 to 'Emails Sent'
        column.
        Parameter row_no has to be passed excluding the header - 2nd row in the sheet is
        the first row that is not header and row_no for this row in 0
        """
        row_list = self.rows[row_no]
        row = dict(zip_longest(self.cols, row_list))
        row_no += 2  # to align with the Sheet

        # Add date if it does not exist
        approached_date = row["Approached (Date)"]
        if not approached_date:
            date_letter = self.cols_mapping["Approached (Date)"]
            curr_date = datetime.now().strftime("%d.%m.%y")
            self._update_values(
                f"{date_letter}{row_no}:{date_letter}{row_no}",
                "USER_ENTERED",
                [[curr_date]],
            )

        # Add +1 to 'Emails Sent' column
        emails_sent = row["Emails Sent"]
        emails_sent = 0 if not emails_sent else int(emails_sent) + 1
        emails_letter = self.cols_mapping["Emails Sent"]
        self._update_values(
            f"{emails_letter}{row_no}:{emails_letter}{row_no}",
            "USER_ENTERED",
            [[emails_sent]],
        )

    def mainloop(self):
        pass


cl = EmailClient()
print(cl._send_email("hehe", 20))
