from __future__ import print_function
import os.path
import re

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
TEMPLATE_ID = "11sJ0EfePCpdH3nVO7__pynvSlcXnv7WzvNf5TeACrxc"

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

def get_sheet_id_by_title(service, spreadsheet_id, title):

    # Get sheet metadata
    metadata = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id
    ).execute()

    # Get and return ID
    for s in metadata['sheets']:
        if s['properties']['title'] == title:

            print(s['properties'])
            
            return s['properties']['sheetId']

    raise ValueError(f"Tab with title '{title}' not found'")
    
def resize_sheet(service, spreadsheet_id, tab_id, nrows):

    resize_request = {
        "requests": [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": tab_id,
                        "gridProperties": {
                            "rowCount": nrows,
                            "columnCount": 27
                        }
                    },
                    "fields": "gridProperties(rowCount,columnCount)"
                }
            }
        ]
    }
    
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=resize_request
    ).execute()


def total_formula(col_letter):
    return f'=COUNTIF({col_letter}$3:INDEX({col_letter}:{col_letter},ROW()-1),TRUE)'
    
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
        "name": "CMS Masterclass - 31 Dec, London",
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

    '''
    Take the first tab contents as a template
    '''
    TEMPLATE_SHEET_TITLE = "Group 1"
    template_sheet = next(
        s for s in sheets_meta if s["properties"]["title"] == TEMPLATE_SHEET_TITLE
    )
    template_sheet_id = template_sheet["properties"]["sheetId"]
    
    requests = []

    new_groups = [
        {
            'name': 'Sheen',
            'ndatasets': 7
        },
        {
            'name': 'Mortlake',
            'ndatasets': 5
        },
        {
            'name': 'Barnes',
            'ndatasets': 3
        }
    ]

    '''
    First rename the first tab
    '''
    requests.append({
        "updateSheetProperties": {
            "properties": {
                "sheetId": template_sheet_id,
                "title": new_groups[0]['name']
            },
            "fields": "title"
        }
    })

    '''
    Then add the rest of the groups
    '''
    for i in range(1, len(new_groups)):
        requests.append({
            "duplicateSheet": {
                "sourceSheetId": template_sheet_id,
                "insertSheetIndex": i,
                "newSheetName": f"{new_groups[i]['name']}"
            }
        })
        
    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=NEW_SPREADSHEET_ID,
        body={"requests": requests}
    ).execute()

    print('Sheets addded to the new file')

    '''
    Iterate through the new sheets add the content 
    '''
    for gi, group in enumerate(new_groups):
    
        sheet_id = get_sheet_id_by_title(
            sheets_service,
            NEW_SPREADSHEET_ID,
            group['name']
        )

        ndatasets = group['ndatasets']
            
        source_start_row_index = 2
        nrows = 100
        ncols = 14
        
        source = {
            "sheetId": sheet_id,
            "startRowIndex": source_start_row_index,
            "endRowIndex": source_start_row_index + nrows, # exclusive
            "startColumnIndex": 0,
            "endColumnIndex": ncols,
        }

        for n in range(1, ndatasets):

            source_end_row_index = source_start_row_index + n*nrows
            
            destination = {
                "sheetId": sheet_id,
                "startRowIndex": source_end_row_index,
                "endRowIndex": source_end_row_index + nrows,
                "startColumnIndex": 0,
                "endColumnIndex": ncols,
            }

            copy_paste_request = {
                "copyPaste": {
                    "source": source,
                    "destination": destination,
                    "pasteType": "PASTE_NORMAL",
                    "pasteOrientation": "NORMAL",
                }
            }
        
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=NEW_SPREADSHEET_ID,
                body={"requests": [copy_paste_request]}
            ).execute()


        '''
        Once the new rows have been added add the summary row
        '''
        checkbox_cols = [
            "C", "D", "E", "F",
            "G", "H", "I", "J",
            "K", "L", "M"
        ]

        row = ["Totals"] + [""] + [total_formula(c) for c in checkbox_cols]
        
        resp = sheets_service.spreadsheets().values().append(
            spreadsheetId=NEW_SPREADSHEET_ID,
            range=f"{group['name']}!A:A",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={"values": [row]},
        ).execute()
        
        updated_range = resp["updates"]["updatedRange"]

        '''
        Get the value of the new row
        '''
        m = re.search(r'!([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$', updated_range)
        new_row = int(m.group(2)) 

        '''
        Add the summary formulas
        '''
        # Total number of electrons
        sheets_service.spreadsheets().values().update(
            spreadsheetId=NEW_SPREADSHEET_ID,
            range=f"{group['name']}!P5",
            valueInputOption="USER_ENTERED",
            body={"values": [[f'=C{new_row} + 2*(E{new_row}+I{new_row}) + 4*(G{new_row})']]}
        ).execute()

        # Total number of muons
        sheets_service.spreadsheets().values().update(
            spreadsheetId=NEW_SPREADSHEET_ID,
            range=f"{group['name']}!Q5",
            valueInputOption="USER_ENTERED",
            body={"values": [[f'=D{new_row} + 2*(F{new_row}+I{new_row}) + 4*(H{new_row})']]}
        ).execute()

        # Total number of W+
        sheets_service.spreadsheets().values().update(
            spreadsheetId=NEW_SPREADSHEET_ID,
            range=f"{group['name']}!P9",
            valueInputOption="USER_ENTERED",
            body={"values": [[f'=J{new_row}']]}
        ).execute()

        # Total number of W-
        sheets_service.spreadsheets().values().update(
            spreadsheetId=NEW_SPREADSHEET_ID,
            range=f"{group['name']}!Q9",
            valueInputOption="USER_ENTERED",
            body={"values": [[f'=K{new_row}']]}
        ).execute()
        
    '''
    Add a Results sheet
    '''
    results_request = {
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": "Results",
                        "index": len(new_groups) + 1
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
