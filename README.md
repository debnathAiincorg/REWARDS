# Weekly Performance Report — Teams Bot

A Python script that reads employee performance data from a SharePoint Excel file and sends a **Weekly Performance Report** as an Adaptive Card to Microsoft Teams via a Power Automate webhook.

---

## How It Works

1. **Downloads the Excel file from SharePoint**
   - Tries a public link first; falls back to Microsoft Graph API with 3 retries on failure.

2. **Reads the "Daily Performance Bonus" sheet**
   - Auto-detects the Date, Name, and point columns — no column positions are hardcoded.

3. **Calculates the date range based on today's weekday automatically**

   | Day of Run | Data Range | Label Style |
   |---|---|---|
   | Monday | All employees shown with 0 points | `Week #N - Jun 15 (Sunday - Week off)` |
   | Tuesday | Monday only | `Week #N - Jun 16` (single date) |
   | Wednesday – Saturday | Monday to yesterday | `Week #N - Jun 16 to Jun 20` |
   | Sunday | Full week Monday – Saturday | `Week #N - Jun 16 to Jun 21 (Complete Week)` |

4. **Calculates payout**: Points × ₹10

5. **Sends an Adaptive Card** to Teams via the Power Automate webhook URL.

6. **Deletes the temp file** after sending.

---

## Prerequisites

- Python 3.x
- Required packages:
  - `requests`
  - `openpyxl`
  - `msal`
  - `python-dotenv`

---

## Setup

### 1. Create a `.env` file

Create a file named `.env` in the project directory with the following variables:

```env
AZURE_CLIENT_ID=your-client-id
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_SECRET=your-client-secret
```

These are used to authenticate with Microsoft Graph API when the public SharePoint link is unavailable.

### 2. Update constants in the script

Open `weekly_report_send_teams.py` and update these values near the top of the file:

| Constant | Description |
|---|---|
| `WEBHOOK_URL` | Your Power Automate HTTP trigger URL |
| `SHAREPOINT_DRIVE_ID` | Drive ID of your SharePoint document library |
| `SHAREPOINT_ITEM_ID` | Item ID of the Excel file in SharePoint |
| `TEMP_FILE` | Local path where the Excel file is temporarily saved |

---

## Installation

```bash
pip install requests openpyxl msal python-dotenv
```

---

## Running the Script

```bash
python weekly_report_send_teams.py
```

The script is designed to run daily (Monday through Sunday). Schedule it with Windows Task Scheduler or any cron-compatible scheduler to run once per day at your preferred time.

---

## Excel File Format

- **Sheet name:** `Daily Performance Bonus`
- **Required columns:** `Date`, `Name`, and any number of point columns
- The script auto-detects all point columns — adding or renaming columns requires no code changes.
- New employees added as rows in the Excel file are picked up automatically in the next run.

---

## Notes

- No hardcoded dates or employee names anywhere — everything is derived from the Excel data and the current date.
- New employee names in the Excel file appear in the Teams report automatically.
- On **Tuesday**, the label shows a single date (Monday) instead of a redundant "Jun 16 to Jun 16" range.
