from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import pickle
import os.path
import re
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("Dynamic Sheets - Payroll Pro")

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

CREDENTIALS_FILE = "cred.json"
TOKEN_FILE = "token.pickle"

# ==================== AUTHORIZED EMAIL ====================
AUTHORIZED_EMAIL = "lotusmissov@gmail.com"

_drive_service = None
_sheets_service = None

# ==================== HEADERS (11 COLUMNS) ====================
HEADERS = [
    "Employee ID",
    "Employee Name",
    "Department",
    "Monthly Salary",
    "Working Days",
    "Paid Leave Days",
    "Loss of Pay Days",
    "Per Day Salary",
    "Gross Salary",
    "LOP Amount",
    "Net Salary"
]

# ==================== FORMULAS ====================
DEFAULT_FORMULA_COLUMNS = {
    7: "=D{ROW}/30",                               # H: Per Day Salary
    8: "=D{ROW}",                                  # I: Gross Salary
    9: "=G{ROW}*H{ROW}",                           # J: LOP Amount
    10: "=I{ROW}-J{ROW}"                           # K: Net Salary
}

PROTECTED_COLUMNS = {"C"}

def get_services():
    global _drive_service, _sheets_service
    if _drive_service and _sheets_service:
        return _drive_service, _sheets_service
    
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            # Login hint - suggests the authorized email
            flow.oauth2session.params['login_hint'] = AUTHORIZED_EMAIL
            creds = flow.run_local_server(port=0)
        
        # Verify the logged-in email matches authorized email
        from google.oauth2.credentials import Credentials
        if hasattr(creds, 'id_token') and creds.id_token:
            import json
            import base64
            # Decode JWT to get email
            payload = creds.id_token.split('.')[1]
            payload += '=' * (4 - len(payload) % 4)  # Add padding
            decoded = json.loads(base64.urlsafe_b64decode(payload))
            logged_email = decoded.get('email', '')
            
            if logged_email.lower() != AUTHORIZED_EMAIL.lower():
                raise Exception(f"❌ Unauthorized! Only {AUTHORIZED_EMAIL} can use this system.\nYou logged in as: {logged_email}")
        
        with open(TOKEN_FILE, "wb") as token:
            pickle.dump(creds, token)
    
    _drive_service = build("drive", "v3", credentials=creds)
    _sheets_service = build("sheets", "v4", credentials=creds)
    return _drive_service, _sheets_service

def extract_spreadsheet_id(url_or_id: str) -> str:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url_or_id)
    return match.group(1) if match else url_or_id

def get_sheet_link(spreadsheet_id: str) -> str:
    return f"https://docs.google.com/spreadsheets/d/{extract_spreadsheet_id(spreadsheet_id)}/edit"

def column_index_to_letter(idx: int) -> str:
    letter = ""
    temp = idx
    while temp >= 0:
        letter = chr(65 + (temp % 26)) + letter
        temp = (temp // 26) - 1
        if temp < 0:
            break
    return letter

# ==================== CREATE SHEET ====================
@mcp.tool(description="Create payroll sheet with folder, headers, employee data and formulas")
def create_payroll_sheet(
    folder_name: str,
    sheet_name: str,
    data_rows: list[list]
) -> str:
    try:
        drive_service, sheets_service = get_services()
        
        folder = drive_service.files().create(body={
            "name": folder_name,
            "mimeType": "application/vnd.google-apps.folder"
        }, fields="id").execute()
        folder_id = folder["id"]
        
        sheet = drive_service.files().create(body={
            "name": sheet_name,
            "mimeType": "application/vnd.google-apps.spreadsheet",
            "parents": [folder_id]
        }, fields="id").execute()
        spreadsheet_id = sheet["id"]
        link = get_sheet_link(spreadsheet_id)
        
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range="Sheet1!A1",
            valueInputOption="RAW",
            body={"values": [HEADERS]}
        ).execute()
        
        if data_rows:
            data_range = f"Sheet1!A2:K{1 + len(data_rows)}"
            sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=data_range,
                valueInputOption="RAW",
                body={"values": data_rows}
            ).execute()
        
        sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_id = sheet_metadata["sheets"][0]["properties"]["sheetId"]
        
        requests = []
        for row_idx in range(2, len(data_rows) + 2):
            for col_idx, template in DEFAULT_FORMULA_COLUMNS.items():
                formula = template.replace("{ROW}", str(row_idx))
                requests.append({
                    "updateCells": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": row_idx - 1,
                            "endRowIndex": row_idx,
                            "startColumnIndex": col_idx,
                            "endColumnIndex": col_idx + 1
                        },
                        "rows": [{"values": [{"userEnteredValue": {"formulaValue": formula}}]}],
                        "fields": "userEnteredValue"
                    }
                })
        
        if requests:
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": requests}
            ).execute()
        
        return f"✓ Sheet created with {len(data_rows)} employees!\n📧 Owner: {AUTHORIZED_EMAIL}\n🔗 Live Link: {link}"
    
    except Exception as e:
        return f"✗ Error: {str(e)}"

# ==================== FIND ROW ====================
def find_row(spreadsheet_id: str, employee_id: str) -> str:
    try:
        _, sheets_service = get_services()
        spreadsheet_id = extract_spreadsheet_id(spreadsheet_id)
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range="Sheet1!A:A"
        ).execute()
        values = result.get("values", [])
        for i, cell in enumerate(values):
            if len(cell) > 0 and str(cell[0]).strip() == str(employee_id).strip():
                return f"✓ Found at row: {i + 1}"
        return "✗ Employee not found"
    except Exception as e:
        return f"✗ Error: {str(e)}"

# ==================== CHANGE SALARY ====================
@mcp.tool(description="Change employee salary - auto updates Per Day, Gross, LOP Amount, Net Salary")
def change_employee_salary(
    spreadsheet_id: str,
    employee_id: str,
    new_salary: float
) -> str:
    find_result = find_row(spreadsheet_id, employee_id)
    if not find_result.startswith("✓"):
        return find_result
    row_num = int(find_result.split("row: ")[1])
    
    result = update_row_smart(spreadsheet_id, row_num, {"D": str(new_salary)})
    link = get_sheet_link(spreadsheet_id)
    return f"{result}\n💰 New Salary: ₹{new_salary:,.0f}\n🔗 {link}"

# ==================== APPLY LEAVE ====================
@mcp.tool(description="Apply leave - Paid (no cut) or LOP (salary deducted)")
def apply_employee_leave(
    spreadsheet_id: str,
    employee_id: str,
    leave_days: int,
    leave_type: str = "LOP"
) -> str:
    if leave_days <= 0:
        return "✗ Leave days must be > 0"
    
    find_result = find_row(spreadsheet_id, employee_id)
    if not find_result.startswith("✓"):
        return find_result
    row_num = int(find_result.split("row: ")[1])
    
    _, sheets_service = get_services()
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id)
    
    values = sheets_service.spreadsheets().values().batchGet(
        spreadsheetId=spreadsheet_id,
        ranges=[f"Sheet1!E{row_num}", f"Sheet1!F{row_num}", f"Sheet1!G{row_num}"]
    ).execute().get("valueRanges", [])
    
    current_working = float(values[0].get("values", [[30]])[0][0] or 30)
    current_paid = float(values[1].get("values", [[0]])[0][0] or 0)
    current_lop = float(values[2].get("values", [[0]])[0][0] or 0)
    
    if current_working < leave_days:
        return f"✗ Cannot apply {leave_days} days - only {current_working} working days left"
    
    updates = {}
    updates["E"] = str(current_working - leave_days)
    
    if leave_type.upper() == "PAID":
        updates["F"] = str(current_paid + leave_days)
        impact = "✅ NO SALARY CUT (Paid Leave)"
    else:
        updates["G"] = str(current_lop + leave_days)
        impact = f"⚠️ SALARY CUT: {leave_days} days × Per Day Salary"
    
    result = update_row_smart(spreadsheet_id, row_num, updates)
    link = get_sheet_link(spreadsheet_id)
    return f"{result}\n🏖️ {leave_type.upper()} Leave: {leave_days} days\n{impact}\n🔗 {link}"

# ==================== SMART UPDATE WITH RECALC ====================
def update_row_smart(spreadsheet_id: str, row_number: int, column_updates: dict[str, str]) -> str:
    try:
        _, sheets_service = get_services()
        spreadsheet_id = extract_spreadsheet_id(spreadsheet_id)
        
        if column_updates:
            requests = [{"range": f"Sheet1!{col}{row_number}", "values": [[val]]} for col, val in column_updates.items()]
            sheets_service.spreadsheets().values().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"valueInputOption": "USER_ENTERED", "data": requests}
            ).execute()
        
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"Sheet1!A{row_number}:ZZ{row_number}",
            valueRenderOption="FORMULA"
        ).execute()
        current_row = result.get("values", [[]])[0]
        
        formula_cells = []
        for col_idx, cell in enumerate(current_row):
            col_letter = column_index_to_letter(col_idx)
            if (col_letter not in column_updates
                and col_letter not in PROTECTED_COLUMNS
                and cell and str(cell).startswith("=")):
                formula_cells.append({"col": col_letter, "formula": str(cell)})
        
        if formula_cells:
            temp = [{"range": f"Sheet1!{f['col']}{row_number}", "values": [[f"{f['formula']}&RAND()"]]} for f in formula_cells]
            sheets_service.spreadsheets().values().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"valueInputOption": "USER_ENTERED", "data": temp}
            ).execute()
            
            restore = [{"range": f"Sheet1!{f['col']}{row_number}", "values": [[f['formula']]]} for f in formula_cells]
            sheets_service.spreadsheets().values().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"valueInputOption": "USER_ENTERED", "data": restore}
            ).execute()
        
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"Sheet1!K{row_number}",
            valueInputOption="USER_ENTERED",
            body={"values": [["Updated"]]}
        ).execute()
        sheets_service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=f"Sheet1!K{row_number}"
        ).execute()
        
        return "✓ All calculations updated successfully!"
    
    except Exception as e:
        return f"✗ Error: {str(e)}"

if __name__ == "__main__":
    mcp.run(transport="sse")