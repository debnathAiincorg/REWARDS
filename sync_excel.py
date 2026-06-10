#!/usr/bin/env python3
"""Sync Excel file from source to repo and push to GitHub"""

import shutil
import os
import subprocess
from datetime import datetime

SOURCE_FILE = r"D:\SHAREPOINT\Strict Employee Performance Analysis.xlsx"
REPO_FILE = "Strict Employee Performance Analysis.xlsx"

def sync_file():
    """Copy Excel file from source to repo folder"""
    if not os.path.exists(SOURCE_FILE):
        print(f"[ERROR] Source file not found: {SOURCE_FILE}")
        return False

    try:
        shutil.copy(SOURCE_FILE, REPO_FILE)
        print(f"[OK] Copied {SOURCE_FILE} to {REPO_FILE}")
        return True
    except Exception as e:
        print(f"[ERROR] Failed to copy file: {e}")
        return False

def git_operations():
    """Perform git add, commit, and push"""
    try:
        # Stage the file
        subprocess.run(["git", "add", REPO_FILE], check=True, capture_output=True)
        print("[OK] Staged file with git add")

        # Check if there are changes to commit
        result = subprocess.run(["git", "diff", "--cached", "--quiet"], capture_output=True)
        if result.returncode != 0:
            # Create commit with timestamp
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            commit_message = f"Sync Excel file: {timestamp}"
            subprocess.run(["git", "commit", "-m", commit_message], check=True, capture_output=True)
            print(f"[OK] Created commit: {commit_message}")

            # Push to remote
            subprocess.run(["git", "push"], check=True, capture_output=True)
            print("[OK] Pushed to GitHub")
        else:
            print("[OK] No changes to commit")

        return True
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] Git operation failed: {e}")
        return False
    except Exception as e:
        print(f"[ERROR] Unexpected error: {e}")
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("SYNC EXCEL FILE TO REPO")
    print("=" * 60)

    if sync_file() and git_operations():
        print("=" * 60)
        print("[OK] Sync completed successfully")
        print("=" * 60)
    else:
        print("=" * 60)
        print("[ERROR] Sync failed")
        print("=" * 60)
        exit(1)
