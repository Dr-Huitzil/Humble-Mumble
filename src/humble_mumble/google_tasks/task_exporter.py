"""
task_exporter.py
================
Generates an actionable Google Tasks-ready CSV export for endpoints requiring IT investigation.
Deduplicates devices to 1 row per physical device, formatted directly for Google Tasks:
  - Google Task List: Clinic / Location (each clinic becomes a separate Task List in Google Tasks)
  - Initial Task: Check Device: <Computer Name>
  - Subtasks: Specific status checks (Missing Huntress, Missing AD, Missing NinjaOne, Offline > 7 Days, etc.)
  - Task Notes: Rich metadata including Last Seen, Last User, Serial Number, Model, and Fleet Presence

Audit Conditions Flagged:
  1. Offline > 7 Days: Endpoint offline/silent in NinjaOne for more than 7 days.
  2. Missing Huntress: Endpoint active in Ninja or AD, but missing Huntress EDR agent.
  3. Missing NinjaOne: Computer enabled in AD or active in Huntress, but missing NinjaOne RMM.
  4. Missing Active Directory: Device in Ninja/Huntress, but missing domain computer object.
  5. AD Disabled Account: Device active in Ninja/Huntress, but computer account is disabled in AD.

Output Destination:
  - Single Master CSV: reports/Tasks/flaggedDevices_MMDDYYYY.csv (sorted by Clinic & Priority)
"""

import os
import sys
import csv
import re
from datetime import datetime
from collections import defaultdict
import openpyxl

from humble_mumble.core import config
from humble_mumble.core.utils import is_valid_serial, normalize_hostname, safe_float, safe_int

def _calculate_priority(subtasks, days_inactive, presence):
    """
    Assign priority (P1 - High, P2 - Medium, P3 - Low) based on security risk.
    - P1: Active online device missing EDR, AD Disabled account, or missing RMM on active machine.
    - P2: Offline 8-30 days, missing domain object, or Huntress lagging.
    - P3: Dormant / offline > 90 days (likely decommission candidate).
    """
    days = days_inactive if days_inactive is not None else 0
    st_text = " ".join(subtasks).lower()
    
    if "disabled" in st_text:
        return "P1 - High"
    if "missing huntress" in st_text and days <= 30:
        return "P1 - High"
    if "missing ninjaone" in st_text and days <= 30:
        return "P1 - High"
    if days > 90:
        return "P3 - Low"
    if days > 7:
        return "P2 - Medium"
    return "P2 - Medium"

def parse_flagged_devices_from_workbook(wb, ref_date=None):
    """
    Extracts all device discrepancy data from an openpyxl Workbook
    containing 'Ninja Inventory', 'AD Inventory', and optionally 'Huntress Inventory'.
    """
    ws_ninja = next((wb[s] for s in ["Ninja Inventory", "Inventory", "Master Inventory"] if s in wb.sheetnames), None)
    ws_ad = next((wb[s] for s in ["AD Inventory", "AD Computers", "Active Directory"] if s in wb.sheetnames), None)
    ws_hunt = next((wb[s] for s in ["Huntress Inventory", "Huntress Agents", "Huntress"] if s in wb.sheetnames), None)

    if not ws_ninja and not ws_ad:
        raise ValueError("Workbook must contain at least 'Ninja Inventory' or 'AD Inventory' worksheets.")

    # Header indexes
    ninja_headers = {ws_ninja.cell(1, c).value: c for c in range(1, ws_ninja.max_column + 1)} if ws_ninja else {}
    ad_headers = {ws_ad.cell(1, c).value: c for c in range(1, ws_ad.max_column + 1)} if ws_ad else {}
    hunt_headers = {ws_hunt.cell(1, c).value: c for c in range(1, ws_hunt.max_column + 1)} if ws_hunt else {}

    # 1. Parse Ninja Devices
    ninja_devs = {}
    ninja_by_serial = {}
    if ws_ninja:
        for r in range(2, ws_ninja.max_row + 1):
            dev = ws_ninja.cell(r, ninja_headers.get("Device", 1)).value
            if not dev:
                continue
            dev_str = str(dev).strip()
            up_days_val = ws_ninja.cell(r, ninja_headers.get("Last Uptime - Days", 8)).value
            try:
                up_days = float(up_days_val) if up_days_val not in [None, ""] else None
            except (ValueError, TypeError):
                up_days = None

            sn_val = str(ws_ninja.cell(r, ninja_headers.get("Serial Number", 5)).value or "N/A").strip()
            dev_entry = {
                "name": dev_str,
                "location": config.canonicalize_clinic_name(ws_ninja.cell(r, ninja_headers.get("Location", 2)).value),
                "role": str(ws_ninja.cell(r, ninja_headers.get("Device Role", 3)).value or "").strip(),
                "model": str(ws_ninja.cell(r, ninja_headers.get("Device Model", 4)).value or "Unknown").strip(),
                "serial": sn_val,
                "user": str(ws_ninja.cell(r, ninja_headers.get("Last Login", 6)).value or "N/A").strip(),
                "last_uptime": str(ws_ninja.cell(r, ninja_headers.get("Last Uptime_formatted", 7)).value or "").strip(),
                "uptime_days": up_days,
                "os": str(ws_ninja.cell(r, ninja_headers.get("OS", 12)).value or "").strip()
            }
            ninja_devs[dev_str.upper()] = dev_entry
            if is_valid_serial(sn_val):
                ninja_by_serial[sn_val.upper()] = dev_entry

    # 2. Parse AD Devices
    ad_devs = {}
    if ws_ad:
        c_ad_name = ad_headers.get("Name", 1)
        c_ad_status = ad_headers.get("Status", 2)
        c_ad_clinic = ad_headers.get("Clinic", 3)
        c_ad_dept = ad_headers.get("Department", 4)
        c_ad_type = ad_headers.get("Device Type", 5)
        c_ad_os = ad_headers.get("Operating System", 6)
        c_ad_logon_date = ad_headers.get("Last Logon Date") or ad_headers.get("LastLogonDate", 8)
        c_ad_logon_days = ad_headers.get("Last Logon (Days)") or ad_headers.get("Last Logon - Days", 9)

        for r in range(2, ws_ad.max_row + 1):
            name = ws_ad.cell(r, c_ad_name).value
            if not name:
                continue
            name_str = str(name).strip()
            logon_days_val = ws_ad.cell(r, c_ad_logon_days).value if c_ad_logon_days else None
            try:
                logon_days = float(logon_days_val) if logon_days_val not in [None, ""] else None
            except (ValueError, TypeError):
                logon_days = None

            ad_devs[name_str.upper()] = {
                "name": name_str,
                "status": str(ws_ad.cell(r, c_ad_status).value or "Enabled").strip(),
                "clinic": config.canonicalize_clinic_name(ws_ad.cell(r, c_ad_clinic).value),
                "dept": str(ws_ad.cell(r, c_ad_dept).value or "").strip(),
                "type": str(ws_ad.cell(r, c_ad_type).value or "").strip(),
                "os": str(ws_ad.cell(r, c_ad_os).value or "").strip(),
                "last_logon_date": str(ws_ad.cell(r, c_ad_logon_date).value or "").strip(),
                "last_logon_days": logon_days
            }

    # 3. Parse Huntress Devices
    hunt_devs = {}
    hunt_by_serial = {}
    if ws_hunt:
        c_h_agent = hunt_headers.get("Agent", 1)
        c_h_org = hunt_headers.get("Organization", 2)
        c_h_os = hunt_headers.get("Operating System", 3)
        c_h_seen = hunt_headers.get("Last Seen", 10)
        c_h_sn = hunt_headers.get("Serial Number", 11)
        c_h_inact = hunt_headers.get("Last Seen - Days") or hunt_headers.get("Days Inactive") or hunt_headers.get("Last Seen (Days)", 14)

        for r in range(2, ws_hunt.max_row + 1):
            agent = ws_hunt.cell(r, c_h_agent).value
            if not agent:
                continue
            agent_str = str(agent).strip()
            inact_val = ws_hunt.cell(r, c_h_inact).value if c_h_inact else None
            try:
                inact_days = float(inact_val) if inact_val not in [None, ""] else None
            except (ValueError, TypeError):
                inact_days = None

            sn_val = str(ws_hunt.cell(r, c_h_sn).value or "").strip() if c_h_sn else ""
            h_obj = {
                "agent": agent_str,
                "org": config.canonicalize_clinic_name(ws_hunt.cell(r, c_h_org).value),
                "os": str(ws_hunt.cell(r, c_h_os).value or "").strip(),
                "last_seen": str(ws_hunt.cell(r, c_h_seen).value or "").strip(),
                "days_inactive": inact_days,
                "serial": sn_val
            }
            hunt_devs[agent_str.upper()] = h_obj
            if is_valid_serial(sn_val):
                hunt_by_serial[sn_val.upper()] = h_obj

    # -------------------------------------------------------------------------
    # Reconcile & Build Flagged Devices Master Map
    # Target Active Fleet: Ninja fleet + Huntress fleet + Active (Enabled) AD objects
    # -------------------------------------------------------------------------
    flagged = {}

    active_ad_keys = {k for k, a in ad_devs.items() if a["status"].lower() in ["enabled", "true"]}
    fleet_keys = set(ninja_devs.keys()) | set(hunt_devs.keys()) | active_ad_keys

    hunt_by_norm = {normalize_hostname(k): v for k, v in hunt_devs.items()}
    ad_by_norm = {normalize_hostname(k): v for k, v in ad_devs.items()}
    ninja_by_norm = {normalize_hostname(k): v for k, v in ninja_devs.items()}

    processed_serials = set()

    for key in fleet_keys:
        norm_k = normalize_hostname(key)
        n_item = ninja_devs.get(key) or ninja_by_norm.get(norm_k)
        h_item = hunt_devs.get(key) or hunt_by_norm.get(norm_k)
        a_item = ad_devs.get(key) or ad_by_norm.get(norm_k)
        
        # Cross-reference Ninja & Huntress by Serial Number primarily
        n_sn = n_item.get("serial", "").upper() if n_item and is_valid_serial(n_item.get("serial")) else ""
        h_sn = h_item.get("serial", "").upper() if h_item and is_valid_serial(h_item.get("serial")) else ""

        # Avoid processing the same physical hardware multiple times under different hostnames
        if n_sn and n_sn in processed_serials:
            continue
        if h_sn and h_sn in processed_serials:
            continue

        if n_sn:
            processed_serials.add(n_sn)
            h_match = hunt_by_serial.get(n_sn)
            if h_match:
                h_item = h_match
            elif h_item and is_valid_serial(h_item.get("serial")) and h_item.get("serial").upper() != n_sn:
                # Disconnect candidate hostname match if it belongs to a different valid serial
                h_item = None
        elif h_sn:
            processed_serials.add(h_sn)
            n_match = ninja_by_serial.get(h_sn)
            if n_match:
                n_item = n_match
            elif n_item and is_valid_serial(n_item.get("serial")) and n_item.get("serial").upper() != h_sn:
                n_item = None

        if n_item and not a_item:
            a_item = ad_devs.get(n_item["name"].upper()) or ad_by_norm.get(normalize_hostname(n_item["name"]))

        # Resolve primary display name & clinic
        dev_name = (n_item["name"] if n_item else None) or (a_item["name"] if a_item else None) or (h_item["agent"] if h_item else key)
        raw_clinic = (n_item["location"] if n_item and n_item["location"] not in ["Unassigned", ""] else None) or \
                     (a_item["clinic"] if a_item and a_item["clinic"] not in ["Unassigned", ""] else None) or \
                     (h_item["org"] if h_item and h_item["org"] not in ["Unassigned", ""] else "Unassigned")
        clinic = config.canonicalize_clinic_name(raw_clinic)

        subtasks = []
        action_items = []
        
        # Hardware & user metadata
        model = (n_item["model"] if n_item else "Unknown")
        serial = (n_item["serial"] if n_item and n_item["serial"] != "N/A" else None) or (h_item["serial"] if h_item else "N/A")
        user = (n_item["user"] if n_item else "N/A")
        os_name = (n_item["os"] if n_item and n_item["os"] else None) or (a_item["os"] if a_item else None) or (h_item["os"] if h_item else "Unknown")

        # Inactivity & last seen
        days_inactive = None
        last_seen = ""
        if n_item:
            days_inactive = n_item["uptime_days"]
            last_seen = n_item["last_uptime"]
        elif h_item:
            days_inactive = h_item["days_inactive"]
            last_seen = h_item["last_seen"]
        elif a_item:
            days_inactive = a_item["last_logon_days"]
            last_seen = a_item["last_logon_date"]

        # Presence Flags
        in_ninja = bool(n_item)
        in_huntress = bool(h_item)
        in_ad = bool(a_item)

        # ---------------------------------------------------------------------
        # Status Subtask 1: Offline > 7 Days in NinjaOne
        # ---------------------------------------------------------------------
        if in_ninja:
            if days_inactive is not None and days_inactive > 7:
                d_int = int(days_inactive)
                subtasks.append(f"Offline > 7 Days ({d_int} days since last seen)")
                if d_int > 90:
                    action_items.append("Audit physical hardware: if retired, decommission AD account & cancel licenses")
                else:
                    action_items.append("Verify endpoint power, network cable/Wi-Fi, and NinjaOne service")

        # ---------------------------------------------------------------------
        # Status Subtask 2: Missing Huntress (in Ninja fleet or active AD)
        # ---------------------------------------------------------------------
        if (in_ninja or (in_ad and a_item["status"].lower() in ["enabled", "true"])) and not in_huntress:
            subtasks.append("Missing Huntress")
            action_items.append("Deploy Huntress agent via NinjaOne policy or installer script")

        # ---------------------------------------------------------------------
        # Status Subtask 3: Missing NinjaOne (In AD or In Huntress)
        # ---------------------------------------------------------------------
        if not in_ninja:
            if in_ad and str(a_item["status"]).lower() in ["enabled", "true"]:
                subtasks.append("Missing NinjaOne (Active in AD)")
                action_items.append("Deploy NinjaOne agent via GPO/script or verify computer rename")
            elif in_huntress:
                subtasks.append("Missing NinjaOne (Active in Huntress)")
                action_items.append("Install NinjaOne RMM agent to bring endpoint under management")

        # ---------------------------------------------------------------------
        # Status Subtask 4: In Ninja (or Huntress), but Missing from Active Directory
        # ---------------------------------------------------------------------
        if in_ninja and not in_ad:
            subtasks.append("Missing AD (Not found in Active Directory)")
            action_items.append("Join computer to domain or verify if workgroup/kiosk")

        # ---------------------------------------------------------------------
        # Status Subtask 5: Account Disabled in Active Directory
        # ---------------------------------------------------------------------
        if in_ad and str(a_item["status"]).lower() in ["disabled", "false"] and (in_ninja or in_huntress):
            subtasks.append("AD Account Disabled")
            action_items.append("Security check: Re-enable AD computer account or decommission hardware")

        # ---------------------------------------------------------------------
        # Only retain devices with at least 1 actionable subtask
        # ---------------------------------------------------------------------
        if subtasks:
            presence_str = f"Ninja: {'Yes' if in_ninja else 'No'} | Huntress: {'Yes' if in_huntress else 'No'} | AD: {'Yes' if in_ad else 'No'}"
            priority = _calculate_priority(subtasks, days_inactive, presence_str)

            # Build Google Tasks Formatted Notes Block
            checklist_lines = "\n".join([f"[ ] {st}" for st in subtasks])
            remediation_lines = "\n".join([f"- {act}" for act in action_items])
            
            notes_block = (
                f"Clinic List: {clinic}\n"
                f"Model: {model} | Serial: {serial}\n"
                f"Last User: {user} | OS: {os_name}\n"
                f"Inactivity: {int(days_inactive) if days_inactive is not None else 'Unknown'} days (Last seen: {last_seen or 'N/A'})\n"
                f"Fleet Presence: {presence_str}\n\n"
                f"Status Checklist:\n{checklist_lines}\n\n"
                f"Recommended Actions:\n{remediation_lines}"
            )

            flagged[key] = {
                "clinic": clinic,
                "device": dev_name,
                "task_title": f"Check Device: {dev_name}",
                "priority": priority,
                "subtask_count": len(subtasks),
                "subtasks": subtasks,
                "subtasks_pipe": " | ".join(subtasks),
                "days_inactive": int(days_inactive) if days_inactive is not None else "",
                "last_seen": last_seen,
                "last_user": user,
                "model": model,
                "serial": serial,
                "os": os_name,
                "presence": presence_str,
                "remediation": "; ".join(action_items),
                "checklist_text": checklist_lines,
                "notes_block": notes_block
            }

    return flagged

def export_flagged_devices_tasks(wb_or_path, output_dir=None, ref_date=None):
    """
    Exports deduplicated flagged device tasks to a single master CSV.
    Creates: reports/Tasks/flaggedDevices_MMDDYYYY.csv (sorted by Clinic & Priority)
    
    Structure directly maps to Google Tasks:
      - Google Task List: Clinic / Location
      - Initial Task: Check Device: <Computer Name>
      - Subtasks: Missing Huntress, Missing AD, Offline > 7 Days, etc.
      - Task Notes: Formatted context block

    Returns a summary dictionary with counts and generated file paths.
    """
    if not output_dir:
        output_dir = os.path.join("reports", "Tasks")

    # Load workbook if path given
    if isinstance(wb_or_path, str):
        if not os.path.exists(wb_or_path):
            raise FileNotFoundError(f"Workbook not found at: {wb_or_path}")
        wb = openpyxl.load_workbook(wb_or_path, data_only=True)
    else:
        wb = wb_or_path

    flagged_map = parse_flagged_devices_from_workbook(wb, ref_date=ref_date)
    
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    today_str = datetime.now().strftime("%m%d%Y")
    master_csv_path = os.path.join(output_dir, f"flaggedDevices_{today_str}.csv")

    # Priority sorting helper
    p_order = {"P1 - High": 1, "P2 - Medium": 2, "P3 - Low": 3}

    # Sort devices by: Clinic, Priority, Days Inactive (descending), Device Name
    sorted_devices = sorted(
        flagged_map.values(),
        key=lambda x: (
            x["clinic"].lower() if x["clinic"] != "Unassigned" else "zzz",
            p_order.get(x["priority"], 99),
            -(x["days_inactive"] if isinstance(x["days_inactive"], (int, float)) else -1),
            x["device"].upper()
        )
    )

    # Determine max subtask count across entire fleet to size headers dynamically
    max_subtasks = max((d["subtask_count"] for d in sorted_devices), default=1)
    max_subtasks = max(max_subtasks, 3)  # Guarantee at least Subtask 1, 2, 3 columns

    headers = [
        "Google Task List (Clinic)",
        "Initial Task (Computer Name)",
        "Computer Name",
        "Priority",
        "Subtask Count"
    ]
    for s_idx in range(1, max_subtasks + 1):
        headers.append(f"Subtask {s_idx}")
    headers.extend([
        "All Subtasks (Pipe-Delimited)",
        "All Subtasks (Checklist)",
        "Days Since Last Seen",
        "Last Seen Date",
        "Last User",
        "Device Model",
        "Serial Number",
        "Operating System",
        "Fleet Presence",
        "Recommended Actions",
        "Task Notes (Details)"
    ])

    def build_row(d):
        row = [
            d["clinic"],
            d["task_title"],
            d["device"],
            d["priority"],
            d["subtask_count"]
        ]
        # Append subtasks into individual columns
        for s_idx in range(max_subtasks):
            if s_idx < len(d["subtasks"]):
                row.append(d["subtasks"][s_idx])
            else:
                row.append("")
        row.extend([
            d["subtasks_pipe"],
            d["checklist_text"],
            d["days_inactive"],
            d["last_seen"],
            d["last_user"],
            d["model"],
            d["serial"],
            d["os"],
            d["presence"],
            d["remediation"],
            d["notes_block"]
        ])
        return row

    # Write Single Master CSV
    with open(master_csv_path, mode="w", newline="", encoding="utf-8-sig") as f_master:
        writer = csv.writer(f_master)
        writer.writerow(headers)
        for dev in sorted_devices:
            writer.writerow(build_row(dev))

    # Summary Metrics
    clinic_counts = defaultdict(int)
    for dev in sorted_devices:
        clinic_counts[dev["clinic"]] += 1

    summary = {
        "total_flagged": len(sorted_devices),
        "total_clinics": len(clinic_counts),
        "master_file": master_csv_path,
        "clinic_counts": dict(clinic_counts),
        "priority_counts": Counter_priority(sorted_devices),
        "subtask_distribution": Counter_subtasks(sorted_devices)
    }

    return summary

def Counter_priority(devices):
    counts = defaultdict(int)
    for d in devices:
        counts[d["priority"]] += 1
    return dict(counts)

def Counter_subtasks(devices):
    counts = defaultdict(int)
    for d in devices:
        counts[d["subtask_count"]] += 1
    return dict(counts)


def main(argv=None):
    if argv is None:
        argv = sys.argv[1:]

    print("=" * 72)
    print("TASK EXPORTER — GOOGLE TASKS CSV GENERATOR")
    print("=" * 72)

    # Locate report workbook
    target_report = None
    if len(argv) > 0 and os.path.exists(argv[0]):
        target_report = argv[0]
    else:
        # Auto-detect latest in reports/Extensive Reports/
        repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
        search_dirs = [
            os.path.join(os.getcwd(), "reports", "Extensive Reports"),
            os.path.join(repo_root, "reports", "Extensive Reports")
        ]
        candidates = []
        for d in search_dirs:
            if os.path.exists(d):
                candidates.extend([os.path.join(d, f) for f in os.listdir(d) if f.endswith(".xlsx") and not f.startswith("~$")])
        if candidates:
            candidates.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            target_report = candidates[0]

    if not target_report:
        print("ERROR: No Extensive Report workbook found to extract tasks from.")
        print("Usage: python -m humble_mumble.google_tasks.task_exporter [path_to_extensive_report.xlsx]")
        sys.exit(1)

    print(f"Reading Extensive Report from: {target_report}...")
    res = export_flagged_devices_tasks(target_report)

    print(f"\nSUCCESS!")
    print(f"Total Actionable Endpoints Flagged: {res['total_flagged']}")
    print(f"Total Clinic Task Lists:            {res['total_clinics']}")
    print(f"Priority Breakdown:                 {res['priority_counts']}")
    print(f"Subtask Distribution:               {res['subtask_distribution']}")
    print(f"\nSingle Master Tasks CSV generated at:")
    print(f"  -> {res['master_file']}")
    print("=" * 72)

if __name__ == "__main__":
    main()
