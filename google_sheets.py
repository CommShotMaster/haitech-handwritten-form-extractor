import os
import re
from pathlib import Path
import gspread
from gspread.utils import rowcol_to_a1
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Scopes required to view and manage Google Sheets and Google Drive
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]

def get_gspread_client():
    """Authenticate and return the gspread client."""
    creds = None
    token_path = Path('token.json')
    creds_path = Path('credentials.json')

    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
        
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not creds_path.exists():
                raise FileNotFoundError(
                    "credentials.json not found! Please generate it from Google Cloud Console "
                    "and place it in the same folder as this application."
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for the next run
        with open(token_path, 'w') as token:
            token.write(creds.to_json())

    return gspread.authorize(creds)

def _get_group(class_str):
    """Determine if a class belongs to Green (1-3) or Yellow (4-6)."""
    if not class_str:
        return "unknown"
        
    m = re.search(r'\d', str(class_str))
    if m:
        digit = int(m.group(0))
        if 1 <= digit <= 3:
            return "green"
        elif 4 <= digit <= 6:
            return "yellow"
    return "unknown"

def export_to_google_sheets(all_data, event_title, is_new=True, sheet_url=""):
    """
    Exports data to Google Sheets, applying Green/Yellow grouping and coloring.
    """
    gc = get_gspread_client()
    
    if is_new:
        sh = gc.create(event_title)
        ws = sh.sheet1
        ws.update_title("Extracted Forms")
        current_row = 3
    else:
        # Open existing sheet
        if "spreadsheets/d/" in sheet_url:
            sheet_id = sheet_url.split("spreadsheets/d/")[1].split("/")[0]
            sh = gc.open_by_key(sheet_id)
        else:
            sh = gc.open(sheet_url) # Assume they pasted the exact name
        ws = sh.sheet1
        # Find next empty row (approximate by looking at Col D)
        col_d = ws.col_values(4)
        current_row = len(col_d) + 1
        if current_row < 3:
            current_row = 3
            
    # Prepare the data
    valid_data = [d for d in all_data if d.get("_status") == "success"]
    
    # Sort data: Green first (1-3), then Yellow (4-6), then unknown
    def sort_key(data):
        group = _get_group(data.get("Class", ""))
        class_str = data.get("Class", "")
        # Sort by group priority, then class string
        if group == "green":
            return (0, class_str)
        elif group == "yellow":
            return (1, class_str)
        return (2, class_str)
        
    valid_data.sort(key=sort_key)
    
    # ─── FORMATTING SETUP ───
    # Define colors
    GREEN_RGB = {"red": 217/255.0, "green": 234/255.0, "blue": 211/255.0}
    YELLOW_RGB = {"red": 255/255.0, "green": 242/255.0, "blue": 204/255.0}
    WHITE_RGB = {"red": 1.0, "green": 1.0, "blue": 1.0}
    
    # Batch update format payload
    requests = []
    
    def get_border_dict():
        border_style = {"style": "SOLID", "color": {"red": 0, "green": 0, "blue": 0}}
        return {
            "top": border_style, "bottom": border_style,
            "left": border_style, "right": border_style
        }

    # If it's a new sheet, write Headers
    if is_new:
        # 1. Title
        ws.merge_cells('A1:I1')
        ws.update_acell('A1', event_title)
        requests.append({
            "repeatCell": {
                "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 9},
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "textFormat": {"bold": True, "fontSize": 14}
                    }
                },
                "fields": "userEnteredFormat(horizontalAlignment,textFormat)"
            }
        })
        
        # 2. Headers
        headers = ["", "Label", "Attd.", "NAME", "NAME", "Class", "Meal", "Mobile", "REMARK"]
        ws.update(range_name='A2:I2', values=[headers])
        requests.append({
            "repeatCell": {
                "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 9},
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "textFormat": {"bold": True},
                        "borders": get_border_dict()
                    }
                },
                "fields": "userEnteredFormat(horizontalAlignment,textFormat,borders)"
            }
        })
        
        # 3. Column Widths
        widths = [40, 80, 80, 150, 300, 100, 100, 150, 200]
        for col_idx, width in enumerate(widths):
            requests.append({
                "updateDimensionProperties": {
                    "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": col_idx, "endIndex": col_idx + 1},
                    "properties": {"pixelSize": width},
                    "fields": "pixelSize"
                }
            })
            
    # Track grouping for merge (first column)
    green_start = None
    green_end = None
    yellow_start = None
    yellow_end = None
    
    update_values = []
    
    # ─── DATA INSERTION ───
    for data in valid_data:
        group = _get_group(data.get("Class", ""))
        
        # Prepare row data: A, B, C, D(Chinese), E(English), F(Class), G(Meal), H(Mobile), I(REMARK)
        row_data = [
            "", "", "", 
            data.get("Name (Chinese)", ""),
            data.get("Name (English)", ""),
            data.get("Class", ""),
            data.get("Food Type", ""),
            data.get("Contact Number", ""),
            ""
        ]
        update_values.append(row_data)
        
        # Track ranges
        row_index = current_row - 1 # 0-indexed for API
        if group == "green":
            if green_start is None:
                green_start = row_index
            green_end = row_index
            bg_color = GREEN_RGB
        elif group == "yellow":
            if yellow_start is None:
                yellow_start = row_index
            yellow_end = row_index
            bg_color = YELLOW_RGB
        else:
            bg_color = WHITE_RGB
            
        # Add formatting request for this row
        requests.append({
            "repeatCell": {
                "range": {"sheetId": ws.id, "startRowIndex": row_index, "endRowIndex": row_index + 1, "startColumnIndex": 0, "endColumnIndex": 9},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": bg_color,
                        "borders": get_border_dict()
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,borders)"
            }
        })
        
        # Add center alignment for specific columns (Class=5, Meal=6, Mobile=7)
        requests.append({
            "repeatCell": {
                "range": {"sheetId": ws.id, "startRowIndex": row_index, "endRowIndex": row_index + 1, "startColumnIndex": 5, "endColumnIndex": 8},
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER"
                    }
                },
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })
        
        current_row += 1

    if update_values:
        start_cell = f'A{current_row - len(update_values)}'
        end_cell = f'I{current_row - 1}'
        ws.update(range_name=f'{start_cell}:{end_cell}', values=update_values)
        
    # Merge column A and insert label text
    # Green group
    if green_start is not None and green_end is not None:
        requests.append({
            "mergeCells": {
                "range": {"sheetId": ws.id, "startRowIndex": green_start, "endRowIndex": green_end + 1, "startColumnIndex": 0, "endColumnIndex": 1},
                "mergeType": "MERGE_COLUMNS"
            }
        })
        ws.update_acell(f'A{green_start + 1}', "青组")
        requests.append({
            "repeatCell": {
                "range": {"sheetId": ws.id, "startRowIndex": green_start, "endRowIndex": green_start + 1, "startColumnIndex": 0, "endColumnIndex": 1},
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "textFormat": {"bold": True}
                    }
                },
                "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat)"
            }
        })
        
    # Yellow group
    if yellow_start is not None and yellow_end is not None:
        requests.append({
            "mergeCells": {
                "range": {"sheetId": ws.id, "startRowIndex": yellow_start, "endRowIndex": yellow_end + 1, "startColumnIndex": 0, "endColumnIndex": 1},
                "mergeType": "MERGE_COLUMNS"
            }
        })
        ws.update_acell(f'A{yellow_start + 1}', "黄组")
        requests.append({
            "repeatCell": {
                "range": {"sheetId": ws.id, "startRowIndex": yellow_start, "endRowIndex": yellow_start + 1, "startColumnIndex": 0, "endColumnIndex": 1},
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "textFormat": {"bold": True}
                    }
                },
                "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat)"
            }
        })

    # Execute all formatting
    if requests:
        sh.batch_update({"requests": requests})
        
    return sh.url
