import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from google.oauth2 import service_account

# Google Sheets API setup
SERVICE_ACCOUNT_FILE = './credentials.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = 'your-spreadsheet-id'

creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

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
