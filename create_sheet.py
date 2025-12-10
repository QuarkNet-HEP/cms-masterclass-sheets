from __future__ import print_function
import os.path

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

OAUTH_CLIENT_FILE = "oauth_client_secret.json"
TOKEN_FILE = "token.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

FOLDER_ID = "1RkvEGKgLl55BYvnuImmu63plUM3H3bud"
TEMPLATE_ID = "1oMps3LKppfXMSI9I2l0JyqkhReWN9kl3X_bwNUauKxQ"

def get_creds():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    '''
    FIXME Request fails. If this happens, remove the token and re-run
    '''
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                OAUTH_CLIENT_FILE, SCOPES
            )
            creds = flow.run_local_server(port=0)
    
        with open(TOKEN_FILE, "w") as token:
            token.write(creds.to_json())

    return creds

'''
TODO

Use click to add command line options such as
name of sheet and number of sheets (groups of students)
'''

def main():
    creds = get_creds()

    sheets_service = build("sheets", "v4", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    copy_body = {
        "name": "CMS Masterclass - 25 Dec, North Pole",
        "parents": [FOLDER_ID]
    }

    copy = drive_service.files().copy(
        fileId=TEMPLATE_ID,
        body=copy_body
    ).execute()

    NEW_SPREADSHEET_ID = copy["id"]
    print(f'New spreadsheet ID: {NEW_SPREADSHEET_ID}')

    spreadsheet = sheets_service.spreadsheets().get(
        spreadsheetId=NEW_SPREADSHEET_ID
    ).execute()

    sheets_meta = spreadsheet["sheets"]

    TEMPLATE_SHEET_TITLE = "Group 1"
    template_sheet = next(
        s for s in sheets_meta if s["properties"]["title"] == TEMPLATE_SHEET_TITLE
    )
    template_sheet_id = template_sheet["properties"]["sheetId"]

    requests = []

    N = 5 

    '''
    Group 1 is already there so start from 2
    '''
    for i in range(2, N+1):
        requests.append({
            "duplicateSheet": {
                "sourceSheetId": template_sheet_id,
                "insertSheetIndex": i,
                "newSheetName": f"Group {i}"
            }
        })

    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=NEW_SPREADSHEET_ID,
        body={"requests": requests}
    ).execute()

    print(f'{N-1} sheets addded to the new file')

    '''
    Add a Results sheet
    '''
    results_request = {
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": "Results",
                        "index": N + 1
                    }
                }
            }
        ]
    }

    result = sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=NEW_SPREADSHEET_ID,
        body=results_request
    ).execute()

    summary_sheet_id = result["replies"][0]["addSheet"]["properties"]["sheetId"]

    
if __name__ == "__main__":
    main()
