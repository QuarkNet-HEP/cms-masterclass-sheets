import os.path
import re
import click

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from typing import List, Tuple, Dict, Optional

OAUTH_CLIENT_FILE = "oauth_client_secret.json"
TOKEN_FILE = "token.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

FOLDER_ID = "1RkvEGKgLl55BYvnuImmu63plUM3H3bud"
TEMPLATE_ID = "14FZdsESQMlfHhIt1ikHCZqWJC0ijcP63pRyZ96z4Wlk"

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

def vstack_formula(col: str, sheets: List[str]) -> str:

    refs = [f'IMPORTRANGE("{url}","Results!{col}1:{col}")' for url in sheets]
    vstack = f"VSTACK({', '.join(refs)})"
    
    return (
        f'=IFERROR(LET(data,{vstack},FILTER(data,ISNUMBER(data))),"")'
    )

@click.command()
@click.option(
    "--name",
    "name",
    type=str,
    required=True,
    help="Add a new spreadsheet with the name: --sheetname <str>"
)
@click.option(
    "--sheet",
    "sheets",
    type=str,
    multiple=True,
    required=False,
    help=": Specify sheets to summarize by url --sheet <str>"
) 
def main(name, sheets):
    creds = get_creds()

    sheets_service = build("sheets", "v4", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    copy_body = {"name": name, "parents": [FOLDER_ID]}
    copy = drive_service.files().copy(fileId=TEMPLATE_ID, body=copy_body).execute()

    NEW_SPREADSHEET_ID = copy["id"]

    # Copy the template to a new sheet
    spreadsheet = sheets_service.spreadsheets().get(
        spreadsheetId=NEW_SPREADSHEET_ID,
        fields="sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))",
    ).execute()

    # Populate the cells for electron/muon and W+/W- information
    
    source_cells = [
        "Results!A3",
        "Results!B3",
        "Results!A7",
        "Results!B7"
    ]

    data = []

    for source_cell in source_cells:

        dest_range = source_cell.replace("Results", "Combination")

        parts = [
            f'IFERROR(VALUE(IMPORTRANGE("{url}","{source_cell}")),0)' for url in sheets
        ]

        formula = "=SUM(" + ",".join(parts) + ")"
            
        data.append({"range": dest_range, "values": [[formula]]})

    sheets_service.spreadsheets().values().batchUpdate(
        spreadsheetId=NEW_SPREADSHEET_ID,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": data,
        },
    ).execute()

    # Copy the data from columns Y,Z from Results to Combination in the new sheet
    # The plots will then be populated automatically
    
    formula_y = vstack_formula("Y", sheets)
    formula_z = vstack_formula("Z", sheets)

    sheets_service.spreadsheets().values().batchUpdate(
        spreadsheetId=NEW_SPREADSHEET_ID,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": "Combination!Y1", "values": [[formula_y]]},
                {"range": "Combination!Z1", "values": [[formula_z]]},
            ],
        },
    ).execute()

    
if __name__ == "__main__":
    main()
