# Downloads source file → sends performance report to Teams

import requests
from openpyxl import load_workbook
from collections import defaultdict
import os
import time
from datetime import datetime, timedelta
from dotenv import load_dotenv

# Load .env file if it exists
load_dotenv(r"D:\SHAREPOINT\.env")

# Load credentials from environment variables (Azure App Registration)
AZURE_CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
AZURE_TENANT_ID = os.environ.get("AZURE_TENANT_ID")
AZURE_CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")

# Teams Webhook URL
WEBHOOK_URL = "https://defaultabe0ba584a8e4ee9985a85449c16df.58.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7a93c82bdb7b4095894f25309cbe4673/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=NymSHVyWZFGo_y0NQLWsEeapIiWl7f0L44B8z0zYAW4"

# File paths
TEMP_FILE = r"D:\sharepoint\temp_source.xlsx"

# SharePoint API details
SHAREPOINT_DRIVE_ID = "b!_Oj5AOOCqUa-6fnpgxmwM4Tmz3IIfOZIhM-bF3vfV8Q7o8oZ3WyrQ4ILTnuUDgHw"
SHAREPOINT_ITEM_ID = "01EUH7IGAHNG3EYW2JJ5C37HVRDHKNUFDB"


def get_access_token():
    try:
        from msal import ConfidentialClientApplication
        authority = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
        app = ConfidentialClientApplication(
            AZURE_CLIENT_ID,
            client_credential=AZURE_CLIENT_SECRET,
            authority=authority
        )
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" in result:
            return result["access_token"]
        else:
            print(f"Note: Authentication failed - {result.get('error_description', result.get('error', 'Unknown error'))}")
            return None
    except Exception as e:
        print(f"Note: Could not authenticate ({type(e).__name__})")
        return None


def download_from_sharepoint_api(token):
    try:
        headers = {"Authorization": f"Bearer {token}"}
        file_url = f"https://graph.microsoft.com/v1.0/drives/{SHAREPOINT_DRIVE_ID}/items/{SHAREPOINT_ITEM_ID}/content"
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
        if response.headers.get('content-type', '').startswith('text/html') or b'<!DOCTYPE' in response.content[:100]:
            print("Note: Public link requires authentication. Using environment variable credentials...")
        else:
            with open(TEMP_FILE, 'wb') as f:
                f.write(response.content)
            print("[OK] Source file downloaded successfully")
            return True
    except Exception as e:
        print(f"Note: Public link failed ({type(e).__name__})")

    if AZURE_CLIENT_ID and AZURE_TENANT_ID and AZURE_CLIENT_SECRET:
        for attempt in range(1, 4):
            print(f"Attempting authenticated download via Microsoft Graph API (attempt {attempt}/3)...")
            token = get_access_token()
            if token and download_from_sharepoint_api(token):
                return True
            if attempt < 3:
                print("Authenticated download failed, retrying in 30 seconds...")
                time.sleep(30)

    print("[ERROR] Failed to download source file from all methods.")
    return False


def get_cumulative_data():
    """Determine date range and read cumulative data from Excel."""
    today = datetime.now().date()
    weekday = today.weekday()  # Mon=0, Tue=1, Wed=2, Thu=3, Fri=4, Sat=5, Sun=6

    this_monday = today - timedelta(days=weekday)

    if weekday == 0:  # Monday → zero report, no date range needed
        week_number = this_monday.isocalendar()[1]
        yesterday = today - timedelta(days=1)
        week_label = f"Week #{week_number} - {yesterday.strftime('%b %d')} (Sunday - Week off)"
        start = None
        end = None
        prev_day_start = None
        prev_day_end = None
    elif weekday == 6:  # Sunday → this week Mon to yesterday (Saturday = complete week)
        start = this_monday
        end = today - timedelta(days=1)  # yesterday = Saturday
        week_number = start.isocalendar()[1]
        week_label = f"Week #{week_number} - {start.strftime('%b %d')} to {end.strftime('%b %d')} (Complete Week)"
        prev_day_start = end
        prev_day_end = end
    else:  # Tue to Sat → this Monday to yesterday
        start = this_monday
        end = today - timedelta(days=1)
        week_number = start.isocalendar()[1]
        if start == end:
            week_label = f"Week #{week_number} - {start.strftime('%b %d')}"
        else:
            week_label = f"Week #{week_number} - {start.strftime('%b %d')} to {end.strftime('%b %d')}"
        prev_day_start = today - timedelta(days=1)
        prev_day_end = today - timedelta(days=1)

    # Load Excel
    wb = load_workbook(TEMP_FILE)
    if "Daily Performance Bonus" in wb.sheetnames:
        ws = wb["Daily Performance Bonus"]
    else:
        ws = max(wb.worksheets, key=lambda s: s.max_row or 0)

    # Detect columns from row 1
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    date_col = name_col = None
    point_cols = []
    for i, h in enumerate(headers):
        if not h:
            continue
        hl = str(h).strip().lower()
        if hl == "date":
            date_col = i + 1
        elif hl == "name":
            name_col = i + 1
        elif hl != "index":
            point_cols.append(i + 1)

    if not name_col:
        print(f"[ERROR] Could not find Name column. Headers: {headers}")
        return None, None, None

    # Get ALL unique employee names from entire sheet
    all_names = []
    seen = set()
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row=row, column=name_col).value
        if val and str(val).strip() and str(val).strip() not in seen:
            seen.add(str(val).strip())
            all_names.append(str(val).strip())

    # Monday → return all employees with 0 points
    if weekday == 0:
        employees = sorted(
            [{"name": n, "points": 0, "amount": 0, "prev_day_points": 0, "prev_day_amount": 0} for n in all_names],
            key=lambda x: x["name"]
        )
        return week_label, employees, None

    # Aggregate points per employee within date range
    totals = defaultdict(int)
    prev_day_totals = defaultdict(int)
    for name in all_names:
        totals[name] = 0
        prev_day_totals[name] = 0

    # Read weekly data
    if date_col:
        for row in range(2, ws.max_row + 1):
            name_val = ws.cell(row=row, column=name_col).value
            date_val = ws.cell(row=row, column=date_col).value
            if not name_val or not date_val:
                continue
            row_date = date_val.date() if hasattr(date_val, 'date') else None
            if not row_date:
                continue
            if not (start <= row_date <= end):
                continue
            pts = sum(int(ws.cell(row=row, column=c).value or 0) for c in point_cols)
            totals[str(name_val).strip()] += pts

    # Read previous day data
    if date_col and prev_day_start and prev_day_end:
        for row in range(2, ws.max_row + 1):
            name_val = ws.cell(row=row, column=name_col).value
            date_val = ws.cell(row=row, column=date_col).value
            if not name_val or not date_val:
                continue
            row_date = date_val.date() if hasattr(date_val, 'date') else None
            if not row_date:
                continue
            if not (prev_day_start <= row_date <= prev_day_end):
                continue
            pts = sum(int(ws.cell(row=row, column=c).value or 0) for c in point_cols)
            prev_day_totals[str(name_val).strip()] += pts

    employees = sorted(
        [{"name": n, "points": p, "amount": p * 10, "prev_day_points": prev_day_totals.get(n, 0), "prev_day_amount": prev_day_totals.get(n, 0) * 10} for n, p in totals.items()],
        key=lambda x: x["points"], reverse=True
    )
    return week_label, employees, None


def format_teams_message(week_label, employees):
    body = [
        {"type": "TextBlock", "text": "Weekly Performance Report", "weight": "Bolder", "size": "Large", "color": "Accent", "wrap": True},
        {"type": "TextBlock", "text": week_label, "weight": "Bolder", "size": "Medium", "spacing": "None", "wrap": True},
        {
            "type": "ColumnSet",
            "separator": True,
            "columns": [
                {"type": "Column", "width": "3", "items": [{"type": "TextBlock", "text": "Name", "weight": "Bolder", "wrap": True}]},
                {"type": "Column", "width": "2", "items": [{"type": "TextBlock", "text": "Previous Day Point", "weight": "Bolder", "horizontalAlignment": "Center", "wrap": True}]},
                {"type": "Column", "width": "2", "items": [{"type": "TextBlock", "text": "Previous Day Amount", "weight": "Bolder", "horizontalAlignment": "Right", "wrap": True}]},
                {"type": "Column", "width": "2", "items": [{"type": "TextBlock", "text": "Weekly Total Point", "weight": "Bolder", "horizontalAlignment": "Center", "wrap": True}]},
                {"type": "Column", "width": "2", "items": [{"type": "TextBlock", "text": "Weekly Total Amount", "weight": "Bolder", "horizontalAlignment": "Right", "wrap": True}]}
            ]
        }
    ]

    for emp in employees:
        name = emp["name"]
        prev_day_points = int(emp["prev_day_points"]) if emp["prev_day_points"] is not None else 0
        prev_day_amount = int(emp["prev_day_amount"]) if emp["prev_day_amount"] is not None else 0
        weekly_points = int(emp["points"]) if emp["points"] is not None else 0
        weekly_amount = int(emp["amount"]) if emp["amount"] else 0
        body.append({
            "type": "ColumnSet",
            "separator": True,
            "style": "default",
            "columns": [
                {"type": "Column", "width": "3", "items": [{"type": "TextBlock", "text": str(name), "wrap": True}]},
                {"type": "Column", "width": "2", "items": [{"type": "TextBlock", "text": str(prev_day_points), "horizontalAlignment": "Center", "wrap": True}]},
                {"type": "Column", "width": "2", "items": [{"type": "TextBlock", "text": f"₹{prev_day_amount}", "horizontalAlignment": "Right", "wrap": True}]},
                {"type": "Column", "width": "2", "items": [{"type": "TextBlock", "text": str(weekly_points), "horizontalAlignment": "Center", "wrap": True}]},
                {"type": "Column", "width": "2", "items": [{"type": "TextBlock", "text": f"₹{weekly_amount}", "horizontalAlignment": "Right", "wrap": True}]}
            ]
        })

    total_prev_day_points = sum(int(emp["prev_day_points"]) if emp["prev_day_points"] else 0 for emp in employees)
    total_prev_day_amount = total_prev_day_points * 10
    total_weekly_points = sum(int(emp["points"]) if emp["points"] else 0 for emp in employees)
    total_weekly_amount = total_weekly_points * 10
    body.append({
        "type": "ColumnSet",
        "separator": True,
        "columns": [
            {"type": "Column", "width": "3", "items": [{"type": "TextBlock", "text": "TOTAL", "weight": "Bolder", "wrap": True}]},
            {"type": "Column", "width": "2", "items": [{"type": "TextBlock", "text": str(total_prev_day_points), "weight": "Bolder", "horizontalAlignment": "Center", "wrap": True}]},
            {"type": "Column", "width": "2", "items": [{"type": "TextBlock", "text": f"₹{total_prev_day_amount}", "weight": "Bolder", "horizontalAlignment": "Right", "wrap": True}]},
            {"type": "Column", "width": "2", "items": [{"type": "TextBlock", "text": str(total_weekly_points), "weight": "Bolder", "horizontalAlignment": "Center", "wrap": True}]},
            {"type": "Column", "width": "2", "items": [{"type": "TextBlock", "text": f"₹{total_weekly_amount}", "weight": "Bolder", "horizontalAlignment": "Right", "wrap": True}]}
        ]
    })

    return {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.2",
        "body": body
    }


def send_teams_webhook_message(webhook_url, adaptive_card):
    headers = {"Content-Type": "application/json"}
    payload = {
        "type": "message",
        "attachments": [{"contentType": "application/vnd.microsoft.card.adaptive", "content": adaptive_card}]
    }
    try:
        response = requests.post(webhook_url, headers=headers, json=payload)
        if response.status_code in [200, 201, 202]:
            return True
        print(f"ERROR sending message: {response.status_code}")
        if response.text:
            print(f"Response: {response.text[:300]}")
        return False
    except Exception as e:
        print(f"ERROR: Could not send Teams message: {e}")
        return False


def cleanup_temp_file():
    for f in [TEMP_FILE]:
        try:
            if os.path.exists(f):
                os.remove(f)
                print(f"[OK] Deleted: {f}")
        except Exception as e:
            print(f"WARNING: Could not delete {f}: {e}")
    print("[OK] Cleanup done")
    return True


# MAIN EXECUTION
if __name__ == "__main__":
    print("=" * 60)
    print("WEEKLY PERFORMANCE REPORT - TEAMS SENDER")
    print("=" * 60)

    if not download_source_file():
        exit(1)

    print("Determining date range and reading Excel data...")
    week_label, employee_data, _ = get_cumulative_data()

    if not week_label or not employee_data:
        print("[ERROR] No data to send.")
        cleanup_temp_file()
        exit(1)

    print(f"[OK] Report: {week_label}")
    print(f"[OK] Employees: {len(employee_data)}")
    print("\nVerification:")
    for emp in employee_data:
        print(f"  {emp['name']}: Prev Day {emp['prev_day_points']} pts (₹{emp['prev_day_amount']}) | Weekly {emp['points']} pts (₹{emp['amount']})")
    print()

    message = format_teams_message(week_label, employee_data)
    print("Sending report to Teams...")
    if send_teams_webhook_message(WEBHOOK_URL, message):
        print("[OK] Report sent to Teams successfully!")
    else:
        print("ERROR: Failed to send message to Teams")
        cleanup_temp_file()
        exit(1)

    cleanup_temp_file()
    print("=" * 60)

