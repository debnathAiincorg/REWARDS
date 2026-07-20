# Weekly Performance Report — Teams Bot

A Python script that reads employee performance data from a SharePoint Excel file and sends **two separate Adaptive Cards** to Microsoft Teams via a Power Automate webhook:

1. **Previous Day Performance Breakdown** — Category-level scores (Punctuality, L&D, Fluency Compliance, Innovation, Extraordinary Performance) for the previous working day
2. **Weekly Performance Report** — Weekly totals with previous-day aggregates and TOTAL row

---

## How It Works

### 1. Download the Excel file from SharePoint

- Tries the public SharePoint link first.
- If the public link redirects to an HTML login page, falls back to **Microsoft Graph API** using Azure credentials from `.env` (locally) or GitHub Actions Secrets (in CI).
- Graph API download retries up to **3 times**, with a 30-second wait between attempts.
- Saves the file temporarily to `TEMP_FILE` — defaults to a path inside the OS temp directory (`tempfile.gettempdir()`), e.g. `C:\Users\<you>\AppData\Local\Temp\temp_source.xlsx` on Windows or `/tmp/temp_source.xlsx` on Linux. This makes the script work unmodified both locally on Windows and on the GitHub Actions `ubuntu-latest` runner — no path setup needed on either platform.

### 2. Read the "Daily Performance Bonus" sheet

- Auto-detects the `Date`, `Name`, and category columns from row 1 — no column positions are hardcoded.
- Identifies category columns: `Punctuality`, `L&D`, `Fluency Compliance`, `Innovation`, `Extraordinary Performance`
- Collects all unique employee names from the sheet automatically.
- New employees added to the Excel file appear in reports with no code changes.
- Deduplicates rows by (name, date) to prevent double-counting.

### 3. Calculate date range based on today's weekday

| Day of Run | Weekly Range | Previous Day (Weekly Card) | Breakdown Card Date |
|---|---|---|---|
| **Monday** | Previous Monday → Previous Sunday | Previous Saturday (with real data) | Previous Sunday (all 0s) |
| **Tuesday** | This Monday only | Yesterday (Monday) | Yesterday (Monday) |
| **Wednesday–Saturday** | This Monday → Yesterday | Yesterday | Yesterday |
| **Sunday** | This Monday → Yesterday (Saturday) | Yesterday (Saturday) | Yesterday (Saturday) |

**Monday Special Handling:**
- Weekly Report shows previous Saturday's actual performance data
- Breakdown card shows all employees with 0s dated Sunday (non-working day)
- Both cards are always sent on Monday, even though breakdown is all zeros

### 4. Aggregate points and calculate payout

- **Weekly Card:**
  - Sums all point columns per employee within the weekly date range
  - Also aggregates the **previous day** separately (for "Previous Day Point/Amount" columns)
  - **Amount = Points × ₹10**
  - Employees sorted by **previous day amount (descending), name (ascending) as tiebreaker**

- **Breakdown Card:**
  - Shows category-level scores from the "Daily Performance Bonus" sheet
  - One row per employee with actual values: Punctuality, L&D, Fluency Compliance, Innovation, Extraordinary Performance
  - On Monday: All employees shown with 0s (since Sunday is non-working)
  - On other days: Only employees with data for that previous day are shown; if no data exists, breakdown card is skipped

### 5. Change detection — skip sending when nothing changed

Before sending anything, the script compares fresh data against a baseline snapshot saved in `last_known_yesterday.json` (see [Snapshot File](#snapshot-file-last_known_yesterdayjson) below) across **3 independent conditions**. If **any one** differs, the run is treated as changed: both cards are sent and the snapshot is refreshed with the new baseline. If **all three** match, the script exits silently without sending to Teams — but still rewrites the snapshot file so the baseline stays current.

1. **Date changed** — does "yesterday" (today − 1 day) match the `date` stored in the snapshot? True on every new calendar day the script runs.
2. **Previous Day Performance Breakdown changed** — does the fresh per-employee, per-category breakdown for yesterday (`get_real_yesterday_data()`) match `snapshot["employees"]`? This catches someone correcting yesterday's category scores in the source Excel *after* the report already went out.
3. **Weekly Performance Report total changed** — does the aggregated weekly Points/Amount per employee (the numbers actually shown on the Weekly Report card) match what's stored under `snapshot["week"]["totals"]`? This is backed by **two** checks combined with OR:
   - `has_weekly_totals_changed()` — the **authoritative** check: compares `get_weekly_totals()` (derived straight from the same `get_cumulative_data()` call that builds the card) against `snapshot["week"]["totals"]`.
   - `has_weekly_data_changed()` — a secondary, more granular **per-day proxy**: compares raw per-day points (`get_weekly_daily_breakdown()`) against `snapshot["week"]["days"]`, but only for days present in both stored and fresh data — it can't detect an entire day's rows disappearing from the source file. That gap is why the totals check above is authoritative; this one is kept purely as an extra day-level signal, since either check firing is enough to trigger a send.

The console always prints which condition(s) fired:
```
Change detection:
  [1] Date changed (yesterday vs. snapshot date): False
  [2] Previous Day Performance Breakdown changed (per-category, yesterday): False
  [3] Weekly Performance Report total changed (per-employee weekly points/amount): True
[INFO] Sending — triggered by: Weekly Performance Report total changed
```

Card titles are always plain — **"Previous Day Performance Breakdown"** and **"Weekly Performance Report"** — whether or not the send was triggered by a correction. There is no "🔧 Corrected Report -" style prefix.

### 6. Send two separate Adaptive Cards to Teams

**Card 1: Previous Day Performance Breakdown** (sent first, if data exists)
```
Title: "Previous Day Performance Breakdown"
Date Label: (e.g., "Jun 21" for Monday's Sunday, or "Jun 22" for other days)

Columns: Name | Punctuality | L&D | Fluency Compliance | Innovation | Extraordinary Performance
```

**Card 2: Weekly Performance Report** (always sent)
```
Title: "Weekly Performance Report"
Week Label: (e.g., "Week #25 - Jun 15 to Jun 21 (Complete Week)")

Columns: Name | Previous Day Point | Previous Day Amount | Weekly Total Point | Weekly Total Amount
TOTAL Row: Aggregated sums
```

Both cards are sent via the same Power Automate webhook URL using HTTP POST, one after the other. After the Previous Day card is confirmed sent, the script waits **5 seconds** (`time.sleep(5)`) before sending the Weekly Report card. This exists because the webhook returns `202 Accepted` as soon as Power Automate *receives* the request, then posts to Teams asynchronously in its own flow run — two rapid webhook calls can occasionally have their Teams posts land out of order. The delay makes correct ordering (Previous Day card appearing before Weekly Report card) far more likely, but since both sends are independent async flow runs on Microsoft's side, **this reduces the chance of out-of-order delivery — it does not guarantee it**.

### 7. Cleanup

Deletes the temporary Excel file after both reports are sent.

---

## Snapshot File (`last_known_yesterday.json`)

This file is the baseline the change-detection system (step 5 above) compares against, and it's rewritten after every run — whether or not a send happened. Structure:

```json
{
  "date": "2026-07-19",
  "employees": {
    "Employee1": {
      "Punctuality": 1, "L&D": 1, "Fluency Compliance": 0,
      "Innovation": 0, "Extraordinary Performance": 0
    }
  },
  "week": {
    "week_start": "2026-07-13",
    "days": {
      "2026-07-13": {"Employee1": 2, "Employee2": 1},
      "2026-07-14": {"Employee1": 2, "Employee2": 0}
    },
    "totals": {
      "Employee1": {"points": 12, "amount": 120},
      "Employee2": {"points": 7, "amount": 70}
    }
  }
}
```

| Key | Set by | Backs condition |
|---|---|---|
| `date` | Calendar date of "yesterday" at the time of the last run | #1 Date changed |
| `employees` | Per-employee, per-category breakdown for yesterday (`get_real_yesterday_data()`) | #2 Previous Day Breakdown changed |
| `week.week_start` | Monday of the current report week | #3 (used to detect week rollover) |
| `week.days` | Per-day, per-employee raw points for the current week (`get_weekly_daily_breakdown()`) | #3, secondary per-day proxy |
| `week.totals` | Per-employee aggregated weekly Points/Amount (`get_weekly_totals()`) — the numbers shown on the Weekly Report card | #3, authoritative check |

`week.totals` is the newer of these keys, added alongside `week.days` (rather than replacing it) so both the day-level and totals-level checks can run — see step 5. An old snapshot file missing `week` or `week.totals` entirely is handled gracefully (treated as "no baseline yet" for that check, not a crash), so upgrading is automatic on the next run.

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
| `WEBHOOK_URL` | Your Power Automate HTTP trigger URL (sends both cards to same webhook). Reads the `TEAMS_WEBHOOK_URL` environment variable first (set it in `.env` locally, or as a GitHub Actions secret in CI) and falls back to this hardcoded value if the env var isn't set — so overriding it doesn't require a code change. |
| `SHAREPOINT_DRIVE_ID` | Drive ID of your SharePoint document library |
| `SHAREPOINT_ITEM_ID` | Item ID of the Excel file in SharePoint |
| `TEMP_FILE` | Path for the temporary Excel download. Defaults to the OS temp directory via `tempfile.gettempdir()` — works on Windows and Linux without changes; override only if you need a specific location. |

---

## Running the Script

```bash
python weekly_report_send_teams.py
```

The script auto-determines the correct date range from the current day — no arguments needed. It can be run locally on a schedule (e.g. Windows Task Scheduler) or via the GitHub Actions workflow below.

### Example Output — send triggered

```
============================================================
WEEKLY PERFORMANCE REPORT - TEAMS SENDER
============================================================
[OK] Source file downloaded successfully
Determining date range and reading Excel data...

Change detection:
  [1] Date changed (yesterday vs. snapshot date): False
  [2] Previous Day Performance Breakdown changed (per-category, yesterday): False
  [3] Weekly Performance Report total changed (per-employee weekly points/amount): True
[INFO] Sending — triggered by: Weekly Performance Report total changed
[OK] Report: Week #25 - Jun 15 to Jun 21 (Complete Week)
[OK] Employees: 10
[OK] Previous Day (Jun 21): 10 employees

Verification:
  Employee1: Prev Day 2 pts (₹20) | Weekly 8 pts (₹80)
  Employee2: Prev Day 1 pts (₹10) | Weekly 5 pts (₹50)
  ...

Previous Day Breakdown (Jun 21):
  Employee2: Punctuality=0, L&D=0, Fluency Compliance=0, Innovation=0, Extraordinary Performance=0
  Employee3: Punctuality=0, L&D=0, Fluency Compliance=0, Innovation=0, Extraordinary Performance=0
  ...

Sending Previous Day Performance Breakdown card to Teams...
[OK] Previous Day card sent successfully!
Sending Weekly Performance Report card to Teams...
[OK] Weekly Report card sent successfully!

[OK] Total cards sent: 2
[OK] Deleted: /tmp/temp_source.xlsx
[OK] Cleanup done
============================================================
```

### Example Output — nothing changed

```
============================================================
WEEKLY PERFORMANCE REPORT - TEAMS SENDER
============================================================
[OK] Source file downloaded successfully
Determining date range and reading Excel data...

Change detection:
  [1] Date changed (yesterday vs. snapshot date): False
  [2] Previous Day Performance Breakdown changed (per-category, yesterday): False
  [3] Weekly Performance Report total changed (per-employee weekly points/amount): False
[OK] Snapshot saved for 2026-07-19
[INFO] No change detected since last check (date, Previous Day Breakdown, and Weekly Report totals all match). Exiting silently.
[INFO] Baseline snapshot updated (no send needed)
[OK] Deleted: /tmp/temp_source.xlsx
[OK] Cleanup done
```

---

## GitHub Actions Workflow

[`.github/workflows/weekly_report.yml`](.github/workflows/weekly_report.yml) runs this script on GitHub's `ubuntu-latest` runner instead of (or in addition to) running it locally.

**Trigger:** `workflow_dispatch` only — the workflow has no `schedule:` (cron) trigger. Scheduling is handled **externally, via Power Automate**, which previously called this workflow through a `repository_dispatch` event; that trigger has since been removed from the workflow, and Power Automate is expected to instead call the GitHub REST API's [workflow dispatch endpoint](https://docs.github.com/en/rest/actions/workflows#create-a-workflow-dispatch-event) (`POST /repos/debnathAiincorg/REWARDS/actions/workflows/weekly_report.yml/dispatches`) with a token that has `actions: write` permission. **Open item:** confirm that Power Automate has actually been repointed to call this endpoint — until then, this workflow only runs when triggered manually.

**To run manually:** GitHub repo → **Actions** tab → **Weekly Performance Report** → **Run workflow** button → select branch `main` → **Run workflow**.

**Required GitHub Secrets** (repo → Settings → Secrets and variables → Actions → New repository secret):

| Secret | Purpose |
|---|---|
| `AZURE_CLIENT_ID` | Azure App Registration client ID (Graph API fallback download) |
| `AZURE_TENANT_ID` | Azure App Registration tenant ID |
| `AZURE_CLIENT_SECRET` | Azure App Registration client secret |
| `TEAMS_WEBHOOK_URL` | *(optional)* overrides the hardcoded Power Automate webhook URL in the script — omit to use the built-in default |

**What the workflow does:**
1. Checks out the repo, sets up Python 3.11, installs `requirements.txt`
2. Runs `weekly_report_send_teams.py` with the secrets above passed in as environment variables
3. Commits `last_known_yesterday.json` back to `main` (as `github-actions[bot]`), but **only if it actually changed** — the runner is ephemeral and doesn't persist files between runs, so this step keeps the correction-detection baseline current for the next run

---

## Excel File Format

- **Sheet name:** `Daily Performance Bonus`
- **Required columns:** `Date`, `Name`, `Punctuality`, `L&D`, `Fluency Compliance`, `Innovation`, `Extraordinary Performance`
- Additional columns (like `Index`) are recognized and skipped
- Any unknown numeric column is treated as a point column (summed in weekly total)
- New employees (rows) are picked up automatically on the next run

### Example Sheet Structure

| Index | Date | Name | Punctuality | L&D | Fluency Compliance | Innovation | Extraordinary Performance |
|---|---|---|---|---|---|---|---|
| 1 | 2026-06-22 | Employee1 | 1 | 1 | 0 | 0 | 0 |
| 2 | 2026-06-22 | Employee2 | 1 | 1 | 0 | 0 | 0 |
| 3 | 2026-06-22 | Employee3 | 0 | 0 | 0 | 0 | 0 |

---

## Key Features

✓ **Automatic column detection** — No hardcoded column positions  
✓ **Automatic employee discovery** — New rows/names picked up dynamically  
✓ **Deduplication** — Prevents double-counting on duplicate (name, date) rows  
✓ **Two-card split** — Breakdown card sent independently from weekly summary  
✓ **Monday special case** — Previous day breakdown shows all employees with 0s for Sunday  
✓ **Fallback authentication** — Public link → Graph API with retries  
✓ **3-condition change detection** — skips sending when date, Previous Day Breakdown, and Weekly Report totals all match the last known snapshot, avoiding duplicate/unchanged Teams messages while still catching later corrections  
✓ **Cross-platform temp path** — works unmodified on Windows (local) and Linux (GitHub Actions runner)  
✓ **GitHub Actions workflow** — runs on demand via `workflow_dispatch`, with the snapshot baseline committed back to the repo automatically  
✓ **Idempotent** — safe to run multiple times; only sends when something actually changed

---

## Notes

- No hardcoded dates or employee names — everything derived from Excel data and current date.
- **Monday** sends breakdown card with all employees + 0s (Sunday is non-working), then weekly card with Saturday's real data.
- **Tuesday** single-day report (Monday only) since start == end.
- **Sorting:** Employees ordered by previous-day amount (highest first), name (A–Z) as tiebreaker.
- **Rows with no date or name columns are skipped** — handles sparse or malformed Excel data gracefully.
- **Card send order** (Previous Day before Weekly Report) relies on a 5-second delay, not a hard guarantee — see step 6 in How It Works.
- **Secrets:** never commit real Azure or webhook credentials. `.env` is gitignored; `.env.example` holds placeholders only. In CI, the same values live in GitHub Actions Secrets (see GitHub Actions Workflow above).
