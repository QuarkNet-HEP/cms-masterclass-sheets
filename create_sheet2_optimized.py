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

#FOLDER_ID = "1RkvEGKgLl55BYvnuImmu63plUM3H3bud"
#TEMPLATE_ID = "11sJ0EfePCpdH3nVO7__pynvSlcXnv7WzvNf5TeACrxc"

FOLDER_ID = "1Y8aPdGVwqcGHw1mLQdy4WQlwvQ3RadEo"
TEMPLATE_ID = "1zXuH-E73bkvQOUnX8-SCp_3wf-0mXOCRlO1buZHj77Q"

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

def get_sheet_id_by_title(
        service,
        spreadsheet_id: str,
        title: str
) -> str:

    # Get sheet metadata
    metadata = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id
    ).execute()

    # Get and return ID
    for s in metadata['sheets']:
        if s['properties']['title'] == title:            
            return s['properties']['sheetId']

    raise ValueError(f"Tab with title '{title}' not found'")
    
def resize_sheet(
        service,
        spreadsheet_id: str,
        tab_id: str,
        nrows: int
):

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

def total_formula(col_letter: str) -> str:
    return f'=COUNTIF({col_letter}$3:INDEX({col_letter}:{col_letter},ROW()-1),TRUE)'

def append_rows(
        service,
        spreadsheet_id: str,
        tab_id: str,
        n_rows: int
):
    '''
    Add n_rows to a tab given by tab_id
    '''
    body = {
        "requests": [{
            "appendDimension": {
                "sheetId": tab_id,
                "dimension": "ROWS",
                "length": n_rows
            }
        }]
    }
    
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body
    ).execute()

def add_dataset_numbers(
        service,
        sheet_name: str,
        spreadsheet_id: str,
        row_number: int,
        dataset_number: int
):
    updates = []

    updates.append({
        "range": f"{sheet_name}!A{row_number}",
        "values": [[int(dataset_number)]],
    })

    service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "valueInputOption": "RAW",
            "data": updates,
        },
    ).execute()

def get_tabs(service, spreadsheet_id: str) -> List[Dict[str,str]]:

    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties(sheetId,title,hidden))"
    ).execute()

    tabs = []
    for sh in meta.get("sheets", []):
        props = sh["properties"]
        title = props["title"]
        
        if title == 'Results':
            continue
        if props.get("hidden", False):
            continue  # optional
        
        tabs.append({"title": title, "sheetId": props["sheetId"]})

    return tabs

def add_sum_totals_to_results(
        service,
        spreadsheet_id: str,
        tab_names: List[str],
        source_cell: str,
        destination_cell: str
):
    
    sum_term = ",".join([f"'{tb}'!{source_cell}" for tb in tab_names])

    values = [{
        "range": f"Results!{destination_cell}",
        "values": [[f"=SUM({sum_term})"]],
    }]

    body = {
        "valueInputOption": "USER_ENTERED",
        "data": values
    }

    service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body
    ).execute()

def quote_sheet(name: str) -> str:
    # Google Sheets uses single quotes for sheet names; escape any single quote by doubling it
    return "'" + name.replace("'", "''") + "'"

def make_vstack_formula(
        tab_names: List[str],
        col: str,
        start_row: str,
        end_row: Optional[str] = None
) -> str:    
    row_part = f"{start_row}:{col}" if end_row is None else f"{start_row}:{end_row}"

    if end_row is None:
        refs = [f"{quote_sheet(t)}!{col}{start_row}:{col}" for t in tab_names]
    else:
        refs = [f"{quote_sheet(t)}!{col}{start_row}:{col}{end_row}" for t in tab_names]

    vstack = f"VSTACK({', '.join(refs)})"
    return (
        f'=IFERROR(LET(data,{vstack},FILTER(data,ISNUMBER(data))),"")'
    )

def _chunked(iterable, chunk_size: int):
    for i in range(0, len(iterable), chunk_size):
        yield iterable[i:i + chunk_size]

def add_mass_histograms(service, results_sheet_id: str, spreadsheet_id: str):

    start_row_index = 0
    
    def histogram_add_request(*, col_index: int, title: str, anchor_row: int):
        return {
            "addChart": {
                "chart": {
                    "spec": {
                        "title": title,
                        "histogramChart": {
                            "legendPosition": "NO_LEGEND",
                            "series": [
                                {
                                    "data": {
                                        "sourceRange": {
                                            "sources": [
                                                {
                                                    "sheetId": results_sheet_id,
                                                    "startRowIndex": start_row_index,
                                                    # omit endRowIndex to include to bottom
                                                    "startColumnIndex": col_index,
                                                    "endColumnIndex": col_index + 1,  # exclusive
                                                }
                                            ]
                                        }
                                    }
                                }
                            ],
                            # Optional tuning knobs:
                            # "bucketSize": 5,           # fixed bin width, if you want
                            # "outlierPercentile": 0.0,  # include outliers normally
                            # "showItemDividers": False
                        },
                    },
                    "position": {
                        "overlayPosition": {
                            "anchorCell": {
                                "sheetId": results_sheet_id,
                                "rowIndex": anchor_row,
                                "columnIndex": 0,  # column A
                            },
                            "offsetXPixels": 0,
                            "offsetYPixels": 0,
                            "widthPixels": 700,
                            "heightPixels": 380,
                        }
                    },
                }
            }
        }

    requests = [
        histogram_add_request(
            col_index=24,  # Y
            title="2-lepton invariant mass",
            anchor_row=10
        ),
        histogram_add_request(
            col_index=25,  # Z
            title="4-lepton invariant mass",
            anchor_row=30
        ),
    ]

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()


@click.command()
@click.option(
    "--name",
    "name",
    type=str,
    required=True,
    help="Add a new spreadsheet with the name: --sheetname <str>"
)
@click.option(
    "--tab",
    "tabs",
    type=(str, int),
    multiple=True,
    required=True,
    help="Add a new tab (name and number of datasets): --tab <str> <int> (repeatable)"
) 
def main(name, tabs):
    creds = get_creds()

    sheets_service = build("sheets", "v4", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    about = drive_service.about().get(fields="user").execute()
    print("Authenticated as:", about["user"]["emailAddress"])

    copy_body = {"name": name, "parents": [FOLDER_ID]}
    
    copy = drive_service.files().copy(
        fileId=TEMPLATE_ID,
        body=copy_body,
        supportsAllDrives=True
    ).execute()

    NEW_SPREADSHEET_ID = copy["id"]
    print(f"New spreadsheet ID: {NEW_SPREADSHEET_ID}")

    spreadsheet = sheets_service.spreadsheets().get(
        spreadsheetId=NEW_SPREADSHEET_ID,
        fields="sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))",
    ).execute()

    sheets_meta = spreadsheet["sheets"]

    # Take the first tab contents as a template
    TEMPLATE_SHEET_TITLE = "Group 1"
    template_sheet = next(s for s in sheets_meta if s["properties"]["title"] == TEMPLATE_SHEET_TITLE)
    template_sheet_id = template_sheet["properties"]["sheetId"]

    tabs = list(tabs)
    new_groups = [{"name": t[0], "ndatasets": t[1]} for t in tabs]

    print(f"Making a new sheet: {name}")
    print(f"With tabs: {new_groups}")

    # Rename + duplicate sheets in one batchUpdate
    init_requests = [{
        "updateSheetProperties": {
            "properties": {"sheetId": template_sheet_id, "title": new_groups[0]["name"]},
            "fields": "title",
        }
    }]

    for i in range(1, len(new_groups)):
        init_requests.append({
            "duplicateSheet": {
                "sourceSheetId": template_sheet_id,
                "insertSheetIndex": i,
                "newSheetName": f"{new_groups[i]['name']}",
            }
        })

    init_resp = sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=NEW_SPREADSHEET_ID,
        body={"requests": init_requests},
    ).execute()

    # Build a title -> sheetId map
    sheet_id_map = {new_groups[0]["name"]: template_sheet_id}
    replies = init_resp.get("replies", [])

    # replies[0] corresponds to updateSheetProperties (no reply payload),
    # replies[1..] correspond to duplicateSheet calls.
    for idx, grp in enumerate(new_groups[1:], start=1):
        props = None
        if idx < len(replies):
            r = replies[idx] or {}
            if isinstance(r, dict) and "duplicateSheet" in r:
                ds = r["duplicateSheet"] or {}
                if isinstance(ds, dict):
                    props = ds.get("properties")
        if isinstance(props, dict) and "sheetId" in props:
            sheet_id_map[grp["name"]] = props["sheetId"]

    # If some IDs weren't found in replies, fetch once.
    if len(sheet_id_map) != len(new_groups):
        meta2 = sheets_service.spreadsheets().get(
            spreadsheetId=NEW_SPREADSHEET_ID,
            fields="sheets(properties(sheetId,title))",
        ).execute()
        for sh in meta2.get("sheets", []):
            props = sh.get("properties", {})
            if "title" in props and "sheetId" in props:
                sheet_id_map[props["title"]] = props["sheetId"]

    used_ids = set(sheet_id_map.values())
    results_sheet_id = max(used_ids) + 1

    # Everything structural (append rows + copy/paste blocks + add Results sheet)
    # goes into one (or a few chunked) batchUpdate(s).
    structural_requests = []
    value_updates = []

    nrows = 100
    ncols = 14
    source_start_row_index = 2  # 0-based => row 3

    checkbox_cols = ["C","D","E","F","G","H","I","J","K","L","M"]

    for group in new_groups:

        sheet_name = group["name"]
        sheet_id = sheet_id_map[sheet_name]
        ndatasets = int(group["ndatasets"])

        print(f"Name: {group['name']}, Number of datasets: {ndatasets}")
        
        # Ensure enough rows if someone asks for huge ndatasets.
        if ndatasets * nrows > 1000:

            print(f"Add rows for {group['name']}")
            
            structural_requests.append({
                "appendDimension": {"sheetId": sheet_id, "dimension": "ROWS", "length": ndatasets * 110}
            })

        # Copy the dataset block down ndatasets-1 times (build requests; do NOT execute in the loop).
        source = {
            "sheetId": sheet_id,
            "startRowIndex": source_start_row_index,
            "endRowIndex": source_start_row_index + nrows,  # exclusive
            "startColumnIndex": 0,
            "endColumnIndex": ncols,  # exclusive
        }

        for n in range(1, ndatasets):
            source_end_row_index = source_start_row_index + n * nrows

            destination = {
                "sheetId": sheet_id,
                "startRowIndex": source_end_row_index,
                "endRowIndex": source_end_row_index + nrows,
                "startColumnIndex": 0,
                "endColumnIndex": ncols,
            }

            structural_requests.append({
                "copyPaste": {
                    "source": source,
                    "destination": destination,
                    "pasteType": "PASTE_NORMAL",
                    "pasteOrientation": "NORMAL",
                }
            })

            # Dataset number label in column A (1-based row)
            value_updates.append({
                "range": f"{sheet_name}!A{source_end_row_index + 1}",
                "values": [[n + 1]],
            })

        # Totals row is deterministic (no need for values.append + regex parsing).
        totals_row = 3 + ndatasets * nrows
        row_values = ["Totals", ""] + [total_formula(c) for c in checkbox_cols]
        value_updates.append({
            "range": f"{sheet_name}!A{totals_row}:M{totals_row}",
            "values": [row_values],
        })

        # Summary formulas
        value_updates.extend([
            {
                "range": f"{sheet_name}!P5",
                "values": [[f"=C{totals_row} + 2*(E{totals_row}+I{totals_row}) + 4*(G{totals_row})"]],
            },
            {
                "range": f"{sheet_name}!Q5",
                "values": [[f"=D{totals_row} + 2*(F{totals_row}+I{totals_row}) + 4*(H{totals_row})"]],
            },
            {"range": f"{sheet_name}!P9", "values": [[f"=J{totals_row}"]]},
            {"range": f"{sheet_name}!Q9", "values": [[f"=K{totals_row}"]]},
        ])

        # Fill the mass columns by copying the template formula down.
        structural_requests.extend([
            {
                "copyPaste": {
                    "source": {
                        "sheetId": template_sheet_id,
                        "startRowIndex": 2,
                        "endRowIndex": 3,
                        "startColumnIndex": 24,
                        "endColumnIndex": 25,
                    },
                    "destination": {
                        "sheetId": sheet_id,
                        "startRowIndex": 2,
                        "endRowIndex": totals_row - 1,  # exclusive; fill up to row totals_row-1
                        "startColumnIndex": 24,
                        "endColumnIndex": 25,
                    },
                    "pasteType": "PASTE_NORMAL",
                    "pasteOrientation": "NORMAL",
                }
            },
            {
                "copyPaste": {
                    "source": {
                        "sheetId": template_sheet_id,
                        "startRowIndex": 2,
                        "endRowIndex": 3,
                        "startColumnIndex": 25,
                        "endColumnIndex": 26,
                    },
                    "destination": {
                        "sheetId": sheet_id,
                        "startRowIndex": 2,
                        "endRowIndex": totals_row - 1,
                        "startColumnIndex": 25,
                        "endColumnIndex": 26,
                    },
                    "pasteType": "PASTE_NORMAL",
                    "pasteOrientation": "NORMAL",
                }
            },
        ])

    # Add Results sheet with an explicit sheetId so we can reference it in the same batch.
    structural_requests.append({
        "addSheet": {
            "properties": {"sheetId": results_sheet_id, "title": "Results", "index": len(new_groups) + 1}
        }
    })

    # Copy the results block from the template into Results.
    structural_requests.append({
        "copyPaste": {
            "source": {
                "sheetId": template_sheet_id,
                "startRowIndex": 2,
                "endRowIndex": 9,
                "startColumnIndex": 15,
                "endColumnIndex": 18,
            },
            "destination": {
                "sheetId": results_sheet_id,
                "startRowIndex": 0,
                "endRowIndex": 7,
                "startColumnIndex": 0,
                "endColumnIndex": 2,
            },
            "pasteType": "PASTE_NORMAL",
            "pasteOrientation": "NORMAL",
        }
    })

    # Run the structural updates (chunked only if you end up with thousands of requests).
    for chunk in _chunked(structural_requests, chunk_size=900):
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=NEW_SPREADSHEET_ID,
            body={"requests": chunk},
        ).execute()

    # Build Results totals in one values.batchUpdate
    tab_names = [g["name"] for g in new_groups]

    def sum_formula(source_cell: str) -> str:
        sum_term = ",".join([f"{quote_sheet(tb)}!{source_cell}" for tb in tab_names])
        return f"=SUM({sum_term})"

    value_updates.extend([
        {"range": "Results!A3", "values": [[sum_formula("P5")]]},  # electron totals
        {"range": "Results!B3", "values": [[sum_formula("Q5")]]},  # muon totals
        {"range": "Results!A7", "values": [[sum_formula("P9")]]},  # W+ totals
        {"range": "Results!B7", "values": [[sum_formula("Q9")]]},  # W- totals
    ])

    # VSTACK formulas for histogram input.
    twol_formula = make_vstack_formula(tab_names, col="Y", start_row=3)
    fourl_formula = make_vstack_formula(tab_names, col="Z", start_row=3)

    value_updates.extend([
        {"range": "Results!Y1", "values": [[twol_formula]]},
        {"range": "Results!Z1", "values": [[fourl_formula]]},
    ])

    # One HTTP call for all values/formulas.
    sheets_service.spreadsheets().values().batchUpdate(
        spreadsheetId=NEW_SPREADSHEET_ID,
        body={"valueInputOption": "USER_ENTERED", "data": value_updates},
    ).execute()

    '''
    Add mass histograms to Results
    '''
    add_mass_histograms(sheets_service, results_sheet_id, NEW_SPREADSHEET_ID)
    
if __name__ == "__main__":
    main()
