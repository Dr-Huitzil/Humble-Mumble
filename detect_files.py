"""
detect_files.py
===============
Smart CSV detection utility.
Inspects CSV headers and timestamps instead of relying on fragile filenames like '(1)' or '(2)'.
Automatically finds the newest 'Device Details', 'Custom Fields', 'Huntress', and 'AD' exports across
the working folder, samples directory, and the user's Downloads folder.
"""

import os
import csv
import glob
import time

def classify_csv(filepath):
    """Inspects the header row of a CSV file to determine its export type."""
    if not filepath or not os.path.exists(filepath) or not os.path.isfile(filepath):
        return "NOT_FOUND"
    
    first_row = None
    for enc in ["utf-8-sig", "utf-8", "cp1252", "latin-1"]:
        try:
            with open(filepath, "r", encoding=enc, errors="replace") as f:
                reader = csv.reader(f)
                first_row = next(reader, None)
                if first_row:
                    break
        except Exception:
            continue

    if not first_row:
        return "EMPTY"

    header = [str(c).replace("\ufeff", "").strip() for c in first_row]
    
    # Check for Huntress Agents
    if "Agent" in header and any(k in header for k in ["Agent Version", "EDR Version", "Last Seen", "Registered"]):
        return "HUNTRESS_AGENTS"
    # Check for Device Details
    elif any(k in header for k in ["Volumes", "Memory Capacity GiB", "Processors Max Clock Speed"]):
        return "DEVICE_DETAILS"
    # Check for Custom Fields (AD Install Date)
    elif any(k in header for k in ["AD Install Date", "AD Install date", "Install Date"]):
        return "CUSTOM_FIELDS"
    # Check for AD Computers Raw Export
    elif ("OperatingSystemVersion" in header and "Column1" in header) or any(k in header for k in ["Last Logon - Days", "Created - Days", "Modified- Days"]):
        return "AD_COMPUTERS_RAW"
    # Check for Redundant Visible Columns
    elif "Last Uptime_formatted" in header and len(header) < 10:
        return "VISIBLE_COLUMNS_REDUNDANT"
    else:
        return "OTHER_CSV"

def find_best_ninja_csvs(search_dirs=None, max_age_days=7):
    """
    Scans specified directories (defaults to current dir + samples + Downloads)
    and returns the newest file for DEVICE_DETAILS, CUSTOM_FIELDS, and HUNTRESS_AGENTS.
    
    If the newest file for a category is older than max_age_days (e.g. 7 days / a week),
    it is marked as stale, but still returned under 'newest_*' for fallback use.
    """
    if search_dirs is None:
        search_dirs = [
            os.getcwd(),
            os.path.join(os.getcwd(), "data"),
            os.path.join(os.getcwd(), "samples"),
            os.path.join(os.getcwd(), "archive"),
            os.path.expanduser("~/Downloads")
        ]

    candidates = []
    for d in search_dirs:
        if os.path.exists(d):
            patterns = [
                "*Devices*.csv", "*devices*.csv",
                "*agents*.csv", "*Agents*.csv",
                "*huntress*.csv", "*Huntress*.csv",
                "*custom*.csv", "*Custom*.csv",
                "*ad_install*.csv", "*export*.csv",
                "*computers*.csv", "*Computers*.csv",
                "*sample*.csv"
            ]
            for pattern in patterns:
                for f in glob.glob(os.path.join(d, pattern)):
                    if os.path.isfile(f):
                        candidates.append(f)

    # Remove duplicates
    candidates = list(dict.fromkeys(candidates))

    details_files = []
    custom_files = []
    huntress_files = []
    redundant_files = []

    now = time.time()

    for c in candidates:
        ctype = classify_csv(c)
        try:
            mtime = os.path.getmtime(c)
        except OSError:
            continue

        if ctype == "DEVICE_DETAILS":
            details_files.append((c, mtime))
        elif ctype == "CUSTOM_FIELDS":
            custom_files.append((c, mtime))
        elif ctype == "HUNTRESS_AGENTS":
            huntress_files.append((c, mtime))
        elif ctype == "VISIBLE_COLUMNS_REDUNDANT":
            redundant_files.append((c, mtime))

    # Sort newest first
    details_files.sort(key=lambda x: x[1], reverse=True)
    custom_files.sort(key=lambda x: x[1], reverse=True)
    huntress_files.sort(key=lambda x: x[1], reverse=True)

    def pick_within_window(file_list):
        if not file_list:
            return None, None, None
        newest_file, newest_mtime = file_list[0]
        age_days = round((now - newest_mtime) / 86400.0, 1)
        if age_days <= max_age_days:
            return newest_file, None, (newest_file, age_days)
        else:
            return None, (newest_file, age_days), (newest_file, age_days)

    best_details, stale_details, newest_details = pick_within_window(details_files)
    best_custom, stale_custom, newest_custom = pick_within_window(custom_files)
    best_huntress, stale_huntress, newest_huntress = pick_within_window(huntress_files)

    return {
        "details_file": best_details,
        "custom_file": best_custom,
        "huntress_file": best_huntress,
        "stale_details": stale_details,
        "stale_custom": stale_custom,
        "stale_huntress": stale_huntress,
        "newest_details": newest_details,
        "newest_custom": newest_custom,
        "newest_huntress": newest_huntress,
        "all_details": [f[0] for f in details_files],
        "all_custom": [f[0] for f in custom_files],
        "all_huntress": [f[0] for f in huntress_files],
        "redundant_found": [f[0] for f in redundant_files]
    }

def find_best_ad_csv(search_dirs=None, max_age_days=7):
    """
    Scans for the newest Active Directory raw export (Computers_Raw_*.csv).
    """
    if search_dirs is None:
        search_dirs = [
            os.getcwd(),
            os.path.join(os.getcwd(), "data"),
            os.path.join(os.getcwd(), "samples"),
            os.path.join(os.getcwd(), "archive"),
            os.path.expanduser("~/Downloads")
        ]

    candidates = []
    for d in search_dirs:
        if os.path.exists(d):
            patterns = [
                "Computers_Raw_*.csv", "*Computers_Raw*.csv",
                "*ad_computers*.csv", "*Computers*.csv",
                "*sample*ad*.csv"
            ]
            for pattern in patterns:
                for f in glob.glob(os.path.join(d, pattern)):
                    if os.path.isfile(f):
                        candidates.append(f)

    candidates = list(dict.fromkeys(candidates))
    ad_files = []
    now = time.time()

    for c in candidates:
        ctype = classify_csv(c)
        try:
            mtime = os.path.getmtime(c)
        except OSError:
            continue
        if ctype == "AD_COMPUTERS_RAW" or "Computers_Raw" in os.path.basename(c):
            ad_files.append((c, mtime))

    ad_files.sort(key=lambda x: x[1], reverse=True)
    if not ad_files:
        return None, None
    newest_file, newest_mtime = ad_files[0]
    age_days = round((now - newest_mtime) / 86400.0, 1)
    if age_days <= max_age_days:
        return newest_file, None
    return None, (newest_file, age_days)

if __name__ == "__main__":
    result = find_best_ninja_csvs()
    print("Best Device Details CSV:", result["details_file"])
    print("Best Custom Fields CSV: ", result["custom_file"])
    print("Best Huntress CSV:      ", result["huntress_file"])
    if result["stale_details"]:
        print(f"Skipped older Device Details CSV: {result['stale_details'][0]} ({result['stale_details'][1]} days old)")
    if result["stale_custom"]:
        print(f"Skipped older Custom Fields CSV: {result['stale_custom'][0]} ({result['stale_custom'][1]} days old)")
    if result["stale_huntress"]:
        print(f"Skipped older Huntress CSV: {result['stale_huntress'][0]} ({result['stale_huntress'][1]} days old)")
    if result["redundant_found"]:
        print("Ignored Redundant Files:", [os.path.basename(f) for f in result["redundant_found"]])
