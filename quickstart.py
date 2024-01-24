import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from . import create

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = "15XYAD0HpbF2I0sXH29Os2mmshWnG_HAzbX4-s1iUuHg"
REPORT_1_RANGE = "'Report 1'!A1:E6"
REPORT_2_RANGE = "'Report 2'!A1:E6"
RECONCILIATION_RANGE = "Reconciliation!A1:E6"


try:
    non_matching = create("Non Matching Results")
except HttpError as error:
    print(f"Unable to created new non-matching table: {error}")
    non_matching = error


def main():
  """
  Shows basic usage of the Sheets API.
  Prints values from a sample spreadsheet.
  """
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
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Read data from Report 1 and Report 2
        sheet = service.spreadsheets()
        report1 = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='Report 1').execute().get('values', [])
        report2 = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='Report 2').execute().get('values', [])

        # Process data and prepare updates
        reconciliation_data = []
        non_matching_data = [['Day', 'Shift', 'Report 1 Value', 'Report 2 Value']]
        for i, row in enumerate(report1):
            new_row = []
            for j, cell in enumerate(row):
                cell_value = cell if i < len(report2) and j < len(report2[i]) and cell == report2[i][j] else ''
                new_row.append(cell_value)
                if cell_value == '' and i != 0 and j != 0:  # Skip header
                    non_matching_data.append([report1[i][0], report1[0][j], cell, report2[i][j] if i < len(report2) and j < len(report2[i]) else ''])
            reconciliation_data.append(new_row)

        # Update Reconciliation sheet
        body = {'values': reconciliation_data}
        result = sheet.values().update(spreadsheetId=SPREADSHEET_ID, range='Reconciliation', valueInputOption='USER_ENTERED', body=body).execute()

        # Create and populate Non Matching Results sheet
        sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID, body={
            "requests": [{"addSheet": {"properties": {"title": "Non Matching Results"}}}]
        }).execute()
        body = {'values': non_matching_data}
        result = sheet.values().update(spreadsheetId=SPREADSHEET_ID, range='Non Matching Results', valueInputOption='USER_ENTERED', body=body).execute()

        print(f"{result.get('updatedCells')} cells updated.")

    except HttpError as err:
        print(err)



if __name__ == "__main__":
  main()
