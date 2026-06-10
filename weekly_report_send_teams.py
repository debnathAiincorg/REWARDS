# Downloads source file → sends performance report to Teams

import requests
from openpyxl import load_workbook
import os
import glob
import time
from datetime import datetime, timedelta
from dotenv import load_dotenv

# Load .env file if it exists
load_dotenv(r"D:\SHAREPOINT\.env")

# Load credentials from environment variables
SHAREPOINT_USERNAME = os.environ.get("SHAREPOINT_USERNAME")
SHAREPOINT_PASSWORD = os.environ.get("SHAREPOINT_PASSWORD")
SHAREPOINT_TENANT = os.environ.get("SHAREPOINT_TENANT")

# Teams Webhook URL (not sensitive - safe to hardcode)
WEBHOOK_URL = "https://defaultabe0ba584a8e4ee9985a85449c16df.58.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7a93c82bdb7b4095894f25309cbe4673/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=NymSHVyWZFGo_y0NQLWsEeapIiWl7f0L44B8z0zYAW4"

# File paths
TEMP_FILE = "/tmp/temp_source.xlsx"
SOURCE_FILE_LOCAL = r"D:\SHAREPOINT\Strict Employee Performance Analysis.xlsx"

# SharePoint API details (for authenticated download)
SHAREPOINT_DRIVE_ID = "b!ncIFrojL106h53P8D_qyinjcKBnN9CFGrFJ26Ac7DPP-kaV-qe0HS73NEabHbqpx"
SHAREPOINT_ITEM_ID = "01ZAECLIUEFXFIUSMHQNAZRWSUI2DSRS5G"


def get_access_token():
    """Get SharePoint access token using credentials from environment variables"""
    try:
        from msal import PublicClientApplication

        CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"
        AUTHORITY = f"https://login.microsoftonline.com/{SHAREPOINT_TENANT}"
        SCOPES = ["https://graph.microsoft.com/.default"]

        app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
        result = app.acquire_token_by_username_password(
            username=SHAREPOINT_USERNAME,
            password=SHAREPOINT_PASSWORD,
            scopes=SCOPES
        )

        if "access_token" in result:
            return result["access_token"]
        else:
            error_msg = result.get("error_description", result.get("error", "Unknown error"))
            print(f"Note: Authentication failed - {error_msg}")
            return None
    except Exception as e:
        print(f"Note: Could not authenticate ({type(e).__name__})")
        return None


def download_from_sharepoint_api(token):
    """Download file from SharePoint using Graph API"""
    try:
        graph_url = "https://graph.microsoft.com/v1.0"
        headers = {"Authorization": f"Bearer {token}"}

        file_url = f"{graph_url}/drives/{SHAREPOINT_DRIVE_ID}/items/{SHAREPOINT_ITEM_ID}/content"
        response = requests.get(file_url, headers=headers, timeout=30)

        if response.status_code == 200:
            with open(TEMP_FILE, 'wb') as f:
                f.write(response.content)
            print("[OK] Source file downloaded successfully")
            return True
        print(f"[ERROR] SharePoint API returned status {response.status_code}")
        return False
    except Exception as e:
        print(f"[ERROR] Could not download via API: {e}")
        return False


def download_source_file():
    """Download source file from SharePoint (local, public link, or authenticated)"""

    # Method 1: Check if local file exists
    print("Checking for local source file...")
    if os.path.exists(SOURCE_FILE_LOCAL):
        try:
            import shutil
            shutil.copy(SOURCE_FILE_LOCAL, TEMP_FILE)
            print("[OK] Source file downloaded successfully")
            return True
        except Exception as e:
            print(f"[ERROR] Failed to copy local file: {e}")
            return False

    # Method 2: Try public SharePoint link
    print("Attempting to download from public SharePoint link...")
    url = "https://cloudaiorg.sharepoint.com/:x:/r/sites/StrictEmployeePerformance/_layouts/15/Doc.aspx?sourcedoc=%7B4CB66907-495B-454F-BF9E-B119D4DA1461%7D&file=Strict%20Employee%20Performance%20Analysis.xlsx&fromShare=true&action=default&mobileredirect=true&download=1"

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream",
        "Accept-Language": "en-US,en;q=0.9"
    }

    try:
        session = requests.Session()
        response = session.get(url, headers=headers, allow_redirects=True, timeout=30)
        response.raise_for_status()

        # Check if we got a valid Excel file
        if response.headers.get('content-type', '').startswith('text/html') or b'<!DOCTYPE' in response.content[:100] or response.headers.get('content-type', '').startswith('text/'):
            print("Note: Public link requires authentication. Using environment variable credentials...")
        else:
            with open(TEMP_FILE, 'wb') as f:
                f.write(response.content)
            print("[OK] Source file downloaded successfully")
            return True
    except Exception as e:
        print(f"Note: Public link failed ({type(e).__name__})")

    # Method 3: Try authenticated download via SharePoint API (with retry)
    if SHAREPOINT_USERNAME and SHAREPOINT_PASSWORD and SHAREPOINT_TENANT:
        max_retries = 3
        retry_delay = 30

        for attempt in range(1, max_retries + 1):
            print(f"Attempting authenticated download via SharePoint API (attempt {attempt}/{max_retries})...")
            token = get_access_token()
            if token and download_from_sharepoint_api(token):
                # Clear credentials from memory after use
                globals()['SHAREPOINT_USERNAME'] = None
                globals()['SHAREPOINT_PASSWORD'] = None
                globals()['SHAREPOINT_TENANT'] = None
                return True

            if attempt < max_retries:
                print(f"Authenticated download failed, retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)

    # All methods failed
    print("")
    print("[ERROR] Failed to download source file from all methods:")
    print("  1. Local file check: " + SOURCE_FILE_LOCAL)
    print("  2. Public SharePoint link (requires accessibility)")
    print("  3. Authenticated download via environment variable credentials")
    print("")
    print("To fix, do ONE of:")
    print("  • Copy Excel file to: " + SOURCE_FILE_LOCAL)
    print("  • Set environment variables: SHAREPOINT_USERNAME, SHAREPOINT_PASSWORD, SHAREPOINT_TENANT")
    print("  • Ensure the public link is accessible (no login required)")
    print("")
    return False


def get_last_week_range():
    """Calculate last completed week (Monday to Sunday)"""
    today = datetime.now()

    # Find this week's Monday (0 = Monday, 6 = Sunday)
    days_since_monday = today.weekday()
    this_monday = today - timedelta(days=days_since_monday)

    # Last week is the previous Monday to Sunday
    last_sunday = this_monday - timedelta(days=1)
    last_monday = last_sunday - timedelta(days=6)

    # Generate week label
    week_number = last_monday.isocalendar()[1]
    week_label = f"Week #{week_number} - {last_monday.strftime('%b %d')} - {last_sunday.strftime('%b %d')}"

    return last_monday.date(), last_sunday.date(), week_label


def read_unique_employee_names_from_source():
    """Read unique employee names from source file (temp_source.xlsx)"""
    try:
        wb = load_workbook(TEMP_FILE)

        # Find sheet with actual data (has more than 3 rows or is Weekly Rewards)
        sheet_name = None
        if "Weekly Rewards" in wb.sheetnames:
            sheet_name = "Weekly Rewards"
        else:
            # Find sheet with most rows
            for sheet in wb.sheetnames:
                if sheet not in ["Sheet1 (2)", "Sheet1 (3)"]:  # Skip empty template sheets
                    if wb[sheet].max_row > 3:
                        sheet_name = sheet
                        break

        if not sheet_name:
            sheet_name = wb.sheetnames[0]

        ws = wb[sheet_name]

        # Find "Name" column (usually in row 1 or row 3)
        name_col = None
        name_row = None
        for row in range(1, min(5, ws.max_row + 1)):
            for col in range(1, ws.max_column + 1):
                header = ws.cell(row=row, column=col).value
                if header and 'name' in str(header).strip().lower():
                    name_col = col
                    name_row = row
                    break

            if name_col:
                break

        if not name_col or not name_row:
            print(f"ERROR: Could not find 'Name' column in {sheet_name}")
            return []

        # Read unique employee names starting from row after header
        employees = set()
        for row in range(name_row + 1, ws.max_row + 1):
            name = ws.cell(row=row, column=name_col).value
            if name and str(name).strip():
                employees.add(str(name).strip())

        if not employees:
            print(f"ERROR: No employee names found in {sheet_name}")
            return []

        employees = sorted(list(employees))
        return employees

    except Exception as e:
        print(f"ERROR reading employee names from source: {e}")
        import traceback
        traceback.print_exc()
        return []


def read_employee_data_from_source(employee_names):
    """Read employee data (name, points, amount) from source file"""
    try:
        wb = load_workbook(TEMP_FILE)

        # Find the sheet with most complete data (Weekly Rewards or Sheet1 (2))
        target_sheet = None
        if "Weekly Rewards" in wb.sheetnames:
            target_sheet = "Weekly Rewards"
        else:
            # Find sheet with "Points" column
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                for col in range(1, ws.max_column + 1):
                    header = ws.cell(row=1, column=col).value
                    if header and 'points' in str(header).strip().lower():
                        target_sheet = sheet
                        break
                if target_sheet:
                    break

        if not target_sheet:
            print("ERROR: Could not find data sheet with Points column")
            return []

        ws = wb[target_sheet]

        # Get week label from Row 1
        week_label = ws['A1'].value or "Performance Report"

        # Find column headers (in row 3 or row 1)
        name_col = None
        points_col = None
        amount_col = None
        header_row = None

        for row in range(1, min(5, ws.max_row + 1)):
            for col in range(1, ws.max_column + 1):
                header = ws.cell(row=row, column=col).value
                if header:
                    header_lower = str(header).strip().lower()
                    if 'name' in header_lower:
                        name_col = col
                        header_row = row
                    elif 'points' in header_lower:
                        points_col = col
                        header_row = row
                    elif 'amount' in header_lower:
                        amount_col = col
                        header_row = row

            if name_col and points_col:
                break

        if not name_col or not header_row:
            print("ERROR: Could not find Name/Points columns")
            return []

        # Read employee data
        employees = []
        employees_seen = set()

        for row in range(header_row + 1, ws.max_row + 1):
            name = ws.cell(row=row, column=name_col).value
            if not name or not str(name).strip():
                break

            name = str(name).strip()

            # Only include employees from the source list
            if name not in employee_names:
                continue

            if name in employees_seen:
                continue

            employees_seen.add(name)

            points = ws.cell(row=row, column=points_col).value if points_col else 0
            amount = ws.cell(row=row, column=amount_col).value if amount_col else 0

            # Convert to appropriate types
            try:
                points = float(points) if points else 0
            except:
                points = 0

            try:
                amount = float(amount) if amount else 0
            except:
                amount = 0

            # Convert points to integer and calculate amount
            points_int = int(points) if points else 0
            amount_calculated = points_int * 10

            employees.append({
                "name": name,
                "points": points_int,
                "amount": amount_calculated
            })

        return week_label, employees

    except Exception as e:
        print(f"ERROR reading employee data from source: {e}")
        import traceback
        traceback.print_exc()
        return None, []




def format_teams_message(week_label, employees):
    """Format data as Adaptive Card (same format as send_teams.py)"""

    # Build adaptive card body
    body = [
        {
            "type": "TextBlock",
            "text": "Weekly Performance Report",
            "weight": "Bolder",
            "size": "Large",
            "color": "Accent"
        },
        {
            "type": "TextBlock",
            "text": week_label,
            "weight": "Bolder",
            "size": "Medium",
            "spacing": "None"
        },
        {
            "type": "ColumnSet",
            "separator": True,
            "columns": [
                {"type": "Column", "width": "stretch", "items": [{"type": "TextBlock", "text": "Name", "weight": "Bolder"}]},
                {"type": "Column", "width": "100px", "items": [{"type": "TextBlock", "text": "Points", "weight": "Bolder", "horizontalAlignment": "Center"}]},
                {"type": "Column", "width": "100px", "items": [{"type": "TextBlock", "text": "Amount", "weight": "Bolder", "horizontalAlignment": "Right"}]}
            ]
        }
    ]

    # Add employee rows
    for emp in employees:
        name = emp["name"]
        points = int(emp["points"]) if emp["points"] is not None else 0
        amount = int(emp["amount"]) if emp["amount"] else 0
        amount_display = f"₹{amount}"

        row = {
            "type": "ColumnSet",
            "separator": True,
            "style": "default",
            "columns": [
                {"type": "Column", "width": "stretch", "items": [{"type": "TextBlock", "text": str(name)}]},
                {"type": "Column", "width": "100px", "items": [{"type": "TextBlock", "text": str(points), "horizontalAlignment": "Center"}]},
                {"type": "Column", "width": "100px", "items": [{"type": "TextBlock", "text": amount_display, "horizontalAlignment": "Right"}]}
            ]
        }

        body.append(row)

    # Build complete adaptive card
    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.2",
        "body": body
    }

    return card


def send_teams_webhook_message(webhook_url, adaptive_card):
    """Send Adaptive Card message to Teams using webhook"""
    headers = {"Content-Type": "application/json"}

    # Wrap adaptive card in message format
    payload = {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": adaptive_card
            }
        ]
    }

    try:
        response = requests.post(webhook_url, headers=headers, json=payload)

        if response.status_code in [200, 201, 202]:
            return True
        else:
            print(f"ERROR sending message: {response.status_code}")
            if response.text:
                print(f"Response: {response.text[:300]}")
            return False
    except Exception as e:
        print(f"ERROR: Could not send Teams message: {e}")
        return False


def cleanup_temp_file():
    """Delete temporary files created during execution"""
    temp_files = [
        "/tmp/temp_source.xlsx",
        r"D:\SHAREPOINT\Rewards.xlsx",
        r"D:\SHAREPOINT\Book2.xlsx"
    ]

    cleanup_count = 0

    # Delete specific temp files
    for temp_file in temp_files:
        try:
            if os.path.exists(temp_file):
                os.remove(temp_file)
                print(f"[OK] Deleted: {temp_file}")
                cleanup_count += 1
        except Exception as e:
            print(f"WARNING: Could not delete {temp_file}: {e}")

    # Delete wildcard backup files
    try:
        backup_pattern = r"D:\SHAREPOINT\Book2_backup_*.xlsx"
        for backup_file in glob.glob(backup_pattern):
            os.remove(backup_file)
            print(f"[OK] Deleted: {backup_file}")
            cleanup_count += 1
    except Exception as e:
        print(f"WARNING: Could not delete backup files: {e}")

    if cleanup_count > 0:
        print(f"[OK] Cleaned up {cleanup_count} temporary file(s)")
        return True
    else:
        print("[OK] No temporary files to clean up")
        return True


# MAIN EXECUTION
if __name__ == "__main__":
    print("=" * 60)
    print("WEEKLY PERFORMANCE REPORT - TEAMS SENDER")
    print("=" * 60)

    # Step 1: Download source file
    if not download_source_file():
        exit(1)

    # Step 2: Read unique employee names from source file
    print("Reading unique employee names from source file...")
    employee_names = read_unique_employee_names_from_source()

    if not employee_names:
        print("ERROR: No employee names found in source file")
        cleanup_temp_file()
        exit(1)

    print(f"[OK] Employees found in source file: {employee_names}")

    # Step 3: Read employee performance data
    print("Reading employee performance data...")
    week_label, employee_data = read_employee_data_from_source(employee_names)

    if not employee_data:
        print("ERROR: No employee data found")
        cleanup_temp_file()
        exit(1)

    print(f"[OK] Processing: {week_label}")
    print(f"[OK] Found {len(employee_data)} employees with data")

    # Step 4: Verify employee data before sending
    print("\nEmployee data verification:")
    for emp in employee_data:
        print(f"  {emp['name']}: {emp['points']} points = Rs.{emp['amount']}")
    print()

    # Step 5: Format and send Teams message
    print("Formatting Teams message...")
    message = format_teams_message(week_label, employee_data)

    print("Sending report to Teams...")
    if send_teams_webhook_message(WEBHOOK_URL, message):
        print("[OK] Report sent to Teams successfully!")
        print(f"Total employees processed: {len(employee_data)}")
    else:
        print("ERROR: Failed to send message to Teams")
        cleanup_temp_file()
        exit(1)

    # Step 6: Cleanup
    cleanup_temp_file()

    print("=" * 60)
