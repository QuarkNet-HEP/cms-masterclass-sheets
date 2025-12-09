from google.oauth2 import service_account
from googleapiclient.discovery import build

SERVICE_ACCOUNT_FILE = "cms-masterclass-903a6954e33a.json"

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)

drive = build("drive", "v3", credentials=creds)
sheets = build("sheets", "v4", credentials=creds)

FOLDER_ID = "1RkvEGKgLl55BYvnuImmu63plUM3H3bud"

N = 5
body = {
    "properties": {"title": "CMS Masterclass - 25 Dec, North Pole"},
    "sheets": [{"properties": {"title": f"Group {i+1}"}} for i in range(N)]
}

spreadsheet = sheets.spreadsheets().create(body=body).execute()
spreadsheet_id = spreadsheet["spreadsheetId"]

drive.files().update(
    fileId=spreadsheet_id,
    addParents=FOLDER_ID,
    removeParents=None,
    fields="id, parents"
).execute()

print("Created spreadsheet:", spreadsheet_id)
print(f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit")
