"""
google_tasks_sync.py
====================
Core engine for synchronizing flagged device audits from CSV to Google Tasks.
Features:
  - Official, secure Google OAuth 2.0 authentication flow with token caching.
  - Automatic creation and reuse of Clinic-based Task Lists.
  - Smart Deduplication:
      * Prevents duplicate parent tasks for the same computer.
      * Prevents duplicate subtasks if an issue is already open.
      * Re-opens completed tasks if new or recurring issues are detected.
      * Appends newly discovered subtasks to existing device tasks.
  - Dry Run / Preview support.
"""

import os
import sys
import csv
import re
import time
from datetime import datetime
from collections import defaultdict

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/tasks"]

def execute_with_retry(request, max_retries=5, initial_delay=1.0, log_callback=None):
    """
    Executes a Google API request with exponential backoff on rate-limit (429)
    or transient server errors (500, 502, 503, 504).
    """
    delay = initial_delay
    for attempt in range(max_retries):
        try:
            return request.execute()
        except HttpError as e:
            status = getattr(e.resp, 'status', None)
            if status in [429, 500, 502, 503, 504]:
                if attempt == max_retries - 1:
                    raise
                msg = f"  [!] Google API Throttled/Server Error ({status}). Retrying in {delay:.1f}s (Attempt {attempt+1}/{max_retries})..."
                if log_callback:
                    log_callback(msg)
                else:
                    print(msg)
                time.sleep(delay)
                delay *= 2
            else:
                raise

def _normalize_subtask_key(title):
    """
    Extract a normalized key from a subtask title so we can match
    issues across audits (e.g., 'Missing Huntress', 'Offline > 7 Days', 'Missing AD').
    """
    t = str(title).lower().strip()
    if "huntress" in t:
        return "missing_huntress"
    if "offline" in t:
        return "offline_over_7"
    if "disabled" in t:
        return "ad_disabled"
    if "missing ad" in t or "active directory" in t:
        return "missing_ad"
    if "missing ninja" in t or "ninjaone" in t:
        return "missing_ninja"
    return re.sub(r'[^a-z0-9]', '_', t)[:30]

def _extract_computer_name(task_title):
    """Extract computer hostname from 'Check Device: <NAME>' or raw name."""
    t = str(task_title).strip()
    if t.upper().startswith("CHECK DEVICE:"):
        return t[13:].strip().upper()
    return t.upper()

def _resolve_config_path(path_str):
    if not path_str:
        return path_str
    if os.path.exists(path_str):
        return path_str
    repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
    root_candidate = os.path.join(repo_root, path_str)
    if os.path.exists(root_candidate):
        return root_candidate
    return path_str

def get_auth_status(credentials_path="credentials.json", token_path="token.json"):
    """Check current Google Tasks authentication status."""
    resolved_creds = _resolve_config_path(credentials_path)
    resolved_token = _resolve_config_path(token_path)
    has_creds = os.path.exists(resolved_creds)
    has_token = os.path.exists(resolved_token)
    is_valid = False

    if has_token:
        try:
            creds = Credentials.from_authorized_user_file(resolved_token, SCOPES)
            if creds and creds.valid:
                is_valid = True
            elif creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
                is_valid = creds.valid
        except Exception:
            is_valid = False

    return {
        "has_credentials": has_creds,
        "has_token": has_token,
        "is_valid": is_valid
    }

def authenticate_google_tasks(credentials_path="credentials.json", token_path="token.json", port=0):
    """
    Authenticates with Google Tasks using OAuth 2.0.
    Opens the default browser for official Google Sign-In on first run.
    Caches token in token_path for subsequent automatic access.
    """
    resolved_creds = _resolve_config_path(credentials_path)
    resolved_token = _resolve_config_path(token_path)
    creds = None
    if os.path.exists(resolved_token):
        try:
            creds = Credentials.from_authorized_user_file(resolved_token, SCOPES)
        except Exception:
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(credentials_path):
                raise FileNotFoundError(
                    f"Google OAuth credentials file '{credentials_path}' not found.\n"
                    "Please download OAuth 2.0 Client Credentials (Desktop app) from Google Cloud Console "
                    f"and save it as '{credentials_path}'."
                )
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=port)

        with open(token_path, "w", encoding="utf-8") as token_file:
            token_file.write(creds.to_json())

    service = build("tasks", "v1", credentials=creds)
    return service

def sync_flagged_csv_to_google_tasks(
    csv_path,
    service=None,
    selected_clinics=None,
    dry_run=False,
    log_callback=None
):
    """
    Reads flaggedDevices_MMDDYYYY.csv and synchronizes it to Google Tasks.
    
    Args:
        csv_path (str): Path to the flagged devices CSV.
        service (Resource): Google Tasks service object (or None to auto-authenticate).
        selected_clinics (list): Optional list of clinic names to sync. None = all.
        dry_run (bool): If True, previews operations without writing to Google Tasks API.
        log_callback (func): Optional callback for logging string messages to a UI or console.

    Returns:
        dict: Detailed statistics of the sync operation.
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"Flagged devices CSV not found: {csv_path}")

    if not dry_run and service is None:
        service = authenticate_google_tasks()

    # Read and group rows by Clinic / Task List
    rows_by_clinic = defaultdict(list)
    with open(csv_path, mode="r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for r in reader:
            clinic = (r.get("Google Task List (Clinic)") or r.get("Clinic / Location") or "Unassigned").strip()
            if selected_clinics and clinic not in selected_clinics:
                continue
            rows_by_clinic[clinic].append(r)

    total_devices = sum(len(rows) for rows in rows_by_clinic.values())
    log(f"Loaded {total_devices} flagged devices across {len(rows_by_clinic)} clinic task lists.")
    if dry_run:
        log(">>> RUNNING IN DRY RUN MODE (No changes will be written to Google Tasks) <<<\n")

    stats = {
        "lists_created": 0,
        "lists_reused": 0,
        "tasks_created": 0,
        "tasks_reopened": 0,
        "tasks_unchanged": 0,
        "subtasks_created": 0,
        "subtasks_reopened": 0,
        "subtasks_skipped": 0,
        "details": []
    }

    # Fetch existing task lists if connected
    existing_lists = {}
    existing_lists_lower = {}
    if not dry_run and service:
        page_token = None
        while True:
            req = service.tasklists().list(maxResults=100, pageToken=page_token)
            res = execute_with_retry(req, log_callback=log)
            for item in res.get("items", []):
                t_name = (item.get("title") or "").strip()
                if t_name:
                    existing_lists[t_name] = item["id"]
                    existing_lists_lower[t_name.lower()] = item["id"]
            page_token = res.get("nextPageToken")
            if not page_token:
                break

    # Process each clinic
    for clinic_name, devices in sorted(rows_by_clinic.items()):
        log(f"\n--- Processing Clinic List: '{clinic_name}' ({len(devices)} flagged devices) ---")
        
        list_id = existing_lists.get(clinic_name) or existing_lists_lower.get(clinic_name.lower())
        if not list_id:
            if dry_run:
                log(f"[DRY RUN] Would create new Google Task List: '{clinic_name}'")
                list_id = f"mock_list_{clinic_name}"
                stats["lists_created"] += 1
            else:
                req = service.tasklists().insert(body={"title": clinic_name})
                new_list = execute_with_retry(req, log_callback=log)
                list_id = new_list["id"]
                existing_lists[clinic_name] = list_id
                existing_lists_lower[clinic_name.lower()] = list_id
                stats["lists_created"] += 1
                log(f"[OK] Created new Google Task List: '{clinic_name}'")
        else:
            stats["lists_reused"] += 1
            log(f"[*] Reusing existing Google Task List: '{clinic_name}'")

        # Fetch existing tasks in this list to prevent duplicates
        existing_parent_tasks = {}  # computer_name -> task_obj
        existing_subtasks = defaultdict(dict)  # parent_id -> {norm_subtask_key: task_obj}

        if not dry_run and service and not list_id.startswith("mock_"):
            page_token = None
            while True:
                req = service.tasks().list(
                    tasklist=list_id,
                    showCompleted=True,
                    showHidden=True,
                    maxResults=100,
                    pageToken=page_token
                )
                tasks_res = execute_with_retry(req, log_callback=log)
                items = tasks_res.get("items", [])
                
                # Separate parents and subtasks
                for t in items:
                    parent_ref = t.get("parent")
                    t_title = (t.get("title") or "").strip()
                    if not parent_ref:
                        # Parent device task
                        c_name = _extract_computer_name(t_title)
                        if c_name:
                            existing_parent_tasks[c_name] = t
                    else:
                        # Subtask
                        sub_key = _normalize_subtask_key(t_title)
                        existing_subtasks[parent_ref][sub_key] = t

                page_token = tasks_res.get("nextPageToken")
                if not page_token:
                    break

        # Sync each device in this clinic
        for dev in devices:
            comp_name = (dev.get("Computer Name") or dev.get("Device Name") or "").strip().upper()
            initial_title = (dev.get("Initial Task (Computer Name)") or f"Check Device: {comp_name}").strip()
            notes = (dev.get("Task Notes (Details)") or dev.get("Google Tasks Formatted Notes") or "").strip()
            
            # Extract subtasks
            subtask_list = []
            pipe_raw = dev.get("All Subtasks (Pipe-Delimited)")
            if pipe_raw:
                subtask_list = [s.strip() for s in pipe_raw.split("|") if s.strip()]
            else:
                for k in ["Subtask 1", "Subtask 2", "Subtask 3", "Subtask 4"]:
                    if dev.get(k):
                        subtask_list.append(dev[k].strip())

            existing_parent = existing_parent_tasks.get(comp_name)

            # -------------------------------------------------------------
            # Case 1: Device Task does NOT exist in Google Tasks yet
            # -------------------------------------------------------------
            if not existing_parent:
                if dry_run:
                    log(f"  [DRY RUN] Would create Parent Task: '{initial_title}'")
                    for st in subtask_list:
                        log(f"    └─ [DRY RUN] Would create Subtask: '{st}'")
                    stats["tasks_created"] += 1
                    stats["subtasks_created"] += len(subtask_list)
                else:
                    # Insert parent task
                    parent_body = {
                        "title": initial_title,
                        "notes": notes,
                        "status": "needsAction"
                    }
                    req = service.tasks().insert(tasklist=list_id, body=parent_body)
                    new_parent = execute_with_retry(req, log_callback=log)
                    parent_id = new_parent["id"]
                    stats["tasks_created"] += 1
                    log(f"  [+] Created Parent Task: '{initial_title}'")
                    time.sleep(0.04)

                    # Insert subtasks
                    for st in subtask_list:
                        sub_body = {
                            "title": st,
                            "status": "needsAction"
                        }
                        req_sub = service.tasks().insert(tasklist=list_id, parent=parent_id, body=sub_body)
                        execute_with_retry(req_sub, log_callback=log)
                        stats["subtasks_created"] += 1
                        log(f"      └─ [+] Added Subtask: '{st}'")
                        time.sleep(0.04)

            # -------------------------------------------------------------
            # Case 2: Device Task ALREADY exists (Deduplication & Smart Merge)
            # -------------------------------------------------------------
            else:
                parent_id = existing_parent["id"]
                parent_status = existing_parent.get("status", "needsAction")
                parent_reopened = False

                # If parent task was previously completed, but still appears in audit -> Re-open!
                if parent_status == "completed":
                    if dry_run:
                        log(f"  [DRY RUN] Would RE-OPEN completed Parent Task: '{initial_title}'")
                        stats["tasks_reopened"] += 1
                    else:
                        existing_parent["status"] = "needsAction"
                        # Append audit note
                        orig_notes = existing_parent.get("notes", "")
                        reopen_stamp = f"\n[Re-opened on {datetime.now().strftime('%m/%d/%Y %H:%M')}: Unresolved audit issues remain]\n"
                        existing_parent["notes"] = reopen_stamp + orig_notes
                        req = service.tasks().update(tasklist=list_id, task=parent_id, body=existing_parent)
                        execute_with_retry(req, log_callback=log)
                        parent_reopened = True
                        stats["tasks_reopened"] += 1
                        log(f"  [!] RE-OPENED Completed Parent Task: '{initial_title}' (Issues still active)")
                        time.sleep(0.04)
                else:
                    stats["tasks_unchanged"] += 1

                # Check subtasks under this parent
                subs_map = existing_subtasks.get(parent_id, {})
                for st in subtask_list:
                    st_key = _normalize_subtask_key(st)
                    existing_sub = subs_map.get(st_key)

                    if not existing_sub:
                        # New subtask not previously recorded
                        if dry_run:
                            log(f"    └─ [DRY RUN] Would append new Subtask: '{st}'")
                            stats["subtasks_created"] += 1
                        else:
                            sub_body = {
                                "title": st,
                                "status": "needsAction"
                            }
                            req_sub = service.tasks().insert(tasklist=list_id, parent=parent_id, body=sub_body)
                            execute_with_retry(req_sub, log_callback=log)
                            stats["subtasks_created"] += 1
                            log(f"      └─ [+] Appended New Subtask: '{st}'")
                            time.sleep(0.04)

                    elif existing_sub.get("status") == "completed":
                        # Subtask was marked completed, but failed again in new audit -> Re-open subtask!
                        if dry_run:
                            log(f"    └─ [DRY RUN] Would RE-OPEN completed Subtask: '{st}'")
                            stats["subtasks_reopened"] += 1
                        else:
                            existing_sub["status"] = "needsAction"
                            req_sub = service.tasks().update(tasklist=list_id, task=existing_sub["id"], body=existing_sub)
                            execute_with_retry(req_sub, log_callback=log)
                            stats["subtasks_reopened"] += 1
                            log(f"      └─ [!] RE-OPENED Subtask: '{st}' (Failed new audit)")
                            time.sleep(0.04)

                    else:
                        # Subtask already open -> Skip duplicate!
                        stats["subtasks_skipped"] += 1
                        if dry_run:
                            log(f"    └─ [DRY RUN] Skipping already open Subtask: '{st}'")

    log("\n" + "=" * 60)
    log("SYNCHRONIZATION COMPLETED SUMMARY")
    log("=" * 60)
    log(f"Task Lists Created:      {stats['lists_created']}")
    log(f"Task Lists Reused:       {stats['lists_reused']}")
    log(f"Parent Tasks Created:    {stats['tasks_created']}")
    log(f"Parent Tasks Re-opened:  {stats['tasks_reopened']}")
    log(f"Subtasks Created:        {stats['subtasks_created']}")
    log(f"Subtasks Re-opened:      {stats['subtasks_reopened']}")
    log(f"Subtasks Skipped (Dup):  {stats['subtasks_skipped']}")
    log("=" * 60)

    return stats
