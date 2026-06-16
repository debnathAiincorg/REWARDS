# Weekly Performance Report — Teams Bot

A Python script that reads employee performance data from a SharePoint Excel file and sends a **Weekly Performance Report** as an Adaptive Card to Microsoft Teams via a Power Automate webhook.

---

## How It Works

### 1. Download the Excel file from SharePoint

- Tries the public SharePoint link first.
- If the public link redirects to an HTML login page, falls back to **Microsoft Graph API** using Azure credentials from `.env`.
- Graph API download retries up to **3 times**, with a 30-second wait between attempts.
- Saves the file temporarily to `TEMP_FILE` (default: `D:\sharepoint\temp_source.xlsx`).

### 2. Read the "Daily Performance Bonus" sheet

- Auto-detects the `Date`, `Name`, and point columns from row 1 — no column positions are hardcoded.
- Collects all unique employee names from the sheet automatically.
- New employees added to the Excel file appear in the report with no code changes.

### 3. Calculate date range based on today's weekday

| Day of Run | Data Range | Label Example |
|---|---|---|
| **Monday** | Previous Monday → Previous Saturday | `Week #24 - Jun 09 to Jun 14 (Complete Week)` |
| **Tuesday** | This Monday only (single date) | `Week #25 - Jun 16` |
| **Wednesday – Saturday** | This Monday → Yesterday | `Week #25 - Jun 16 to Jun 18` |
| **Sunday** | This Monday → Yesterday (Saturday) | `Week #25 - Jun 16 to Jun 21 (Complete Week)` |

> **Monday** shows the **previous week's complete data** (same format as Sunday) so no report day is skipped.
> **Tuesday** shows a single date because start and end are both Monday.

### 4. Aggregate points and calculate payout

- Sums all point columns per employee within the weekly date range.
- Also aggregates the **previous day** separately (shown as a "Previous Day" column).
- **Amount = Points × ₹10**
- Employees are sorted **by weekly points, highest first**.

### 5. Send Adaptive Card to Teams

The card columns are:

| Name | Previous Day Point | Previous Day Amount | Weekly Total Point | Weekly Total Amount |
|---|---|---|---|---|

Sent to Teams via the Power Automate webhook URL using an HTTP POST.

### 6. Cleanup

Deletes the temporary Excel file after the report is sent.

---

## Prerequisites

- Python 3.x
- Required packages:
  - `requests`
  - `openpyxl`
  - `msal`
  - `python-dotenv`

---

## Installation

```bash
pip install requests openpyxl msal python-dotenv
```

---

## Setup

### 1. Create a `.env` file

```env
AZURE_CLIENT_ID=your-client-id
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_SECRET=your-client-secret
```

These credentials are only used if the public SharePoint link requires authentication.

### 2. Update constants in the script

Open `weekly_report_send_teams.py` and update these values near the top:

| Constant | Description |
|---|---|
| `WEBHOOK_URL` | Your Power Automate HTTP trigger URL |
| `SHAREPOINT_DRIVE_ID` | Drive ID of your SharePoint document library |
| `SHAREPOINT_ITEM_ID` | Item ID of the Excel file in SharePoint |
| `TEMP_FILE` | Local path for the temporary Excel download |

---

## Running the Script

```bash
python weekly_report_send_teams.py
```

Schedule it with **Windows Task Scheduler** to run once daily (Monday–Sunday). The script auto-determines the correct date range from the current day — no arguments needed.

---

## Excel File Format

- **Sheet name:** `Daily Performance Bonus`
- **Required columns:** `Date`, `Name`, and any number of point columns
- Any column that is not named `Date`, `Name`, or `Index` is treated as a point column
- New employees (rows) are picked up automatically on the next run

---

## Notes

- No hardcoded dates or employee names — everything is derived from the Excel data and the current date.
- On **Monday**, the script reports the **previous complete week** instead of skipping or showing zeros, so every day produces a meaningful report.
- On **Tuesday**, the label shows a single date (Monday) rather than a redundant "Jun 16 to Jun 16".
