import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import *

# Set up the connection
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file"]
creds = Credentials.from_service_account_file('./credentials.json', scopes=scope)
client = gspread.authorize(creds)


spreadsheet = client.open('15XYAD0HpbF2I0sXH29Os2mmshWnG_HAzbX4-s1iUuHg')

# Open the sheets
report1 = spreadsheet.worksheet('Report 1')
report2 = spreadsheet.worksheet('Report 2')
reconciliation = spreadsheet.worksheet('Reconciliation')
try:
    non_matching = spreadsheet.add_worksheet(
        title="Non Matching Results",
        rows="100",
        cols="20")
except:
    non_matching = spreadsheet.worksheet('Non Matching Results')

report1_values = report1.get_all_values()
report2_values = report2.get_all_values()


non_matching_data = [["Day", "Shift", "Report 1 Value", "Report 2 Value"]]
for i, row in enumerate(report1_values):
    reconciliation_row = []
    for j, cell in enumerate(row):
        if i < len(report2_values) and j < len(report2_values[i]) and cell == report2_values[i][j]:
            reconciliation_row.append(cell)
        else:
            reconciliation_row.append("")
            fmt = cellFormat(backgroundColor=color(1, 0, 0))
            format_cell_range(reconciliation, f"{chr(65+j)}{i+1}", fmt)
            if i != 0 and j != 0:  # Assuming first row and column are headers
                non_matching_data.append([report1_values[i][0], report1_values[0][j], cell, report2_values[i][j] if i < len(report2_values) and j < len(report2_values[i]) else ""])
    reconciliation.append_row(reconciliation_row, value_input_option='USER_ENTERED')

# Update Non Matching Results sheet
non_matching.update('A1', non_matching_data, value_input_option='USER_ENTERED')

print("Reconciliation completed.")
