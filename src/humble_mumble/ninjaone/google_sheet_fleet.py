"""
google_sheet_fleet.py
=====================
Google Sheets Laptop Fleet Integration for NinjaOne Reporting.

Generates a shared Google Sheets-optimized multi-tab laptop fleet report with:
  1. 'All Laptops' (Master Tab) containing 7 required fields:
     Device | Location | Device Role | Device Model | Serial Number | Last Uptime - Days | Refresh Year
  2. Individual Clinic Tabs at the bottom (one tab per clinic) dynamically powered by:
     =IFERROR(SORT(FILTER('All Laptops'!A2:G, 'All Laptops'!B2:B = "<Clinic Name>"), 1, TRUE), "No laptops found")

Excludes non-clinic and administrative locations by default:
  - Administration-Airport
  - Administration-Plaza
  - newWindowsWorkstation (Staging)

Provides three frictionless update workflows for the existing shared Google Sheet:
  - Workflow 1: 1-Click Clipboard Copy (copies TSV data straight to Windows clipboard for Ctrl+V into A2).
  - Workflow 2: Clean CSV export (reports/Google Sheets/Ninja_Laptop_Master.csv) for File -> Import.
  - Workflow 3: Direct API Synchronization via Google Sheets API (updates 'All Laptops' in-place).
"""

import os
import sys
import csv
import glob
import json
import subprocess
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from humble_mumble.core import config
from humble_mumble.core.utils import parse_date as parse_ninja_date, safe_float, normalize_hostname
from humble_mumble.core.detect_files import find_best_ninja_csvs

FLEET_HEADERS = [
    'Device',
    'Location',
    'Device Role',
    'Device Model',
    'Serial Number',
    'Last Uptime - Days',
    'Refresh Year'
]

CONFIG_FILE_NAME = "google_sheets_config.json"
DEFAULT_SPREADSHEET_ID = "11CJq7U7RrZgrkJNoNyxnaW8C-Jn9om_4mrCJuT_dV4E"


def is_laptop_device(role, model="", device=""):
    """Determines whether a device is a laptop based on Device Role or Model."""
    r = str(role or "").strip().lower()
    m = str(model or "").strip().lower()
    if "laptop" in r or "notebook" in r:
        return True
    if r == "windows desktop" and ("latitude" in m or "thinkpad" in m or "surface book" in m or "surface laptop" in m):
        return True
    return False


def get_clinic_tab_name(existing_sheets, clinic_name):
    """Finds the best matching existing tab for a clinic, supporting '<Clinic> Laptops' conventions."""
    if clinic_name in existing_sheets:
        return clinic_name
    candidates = [
        f"{clinic_name} Laptops",
        f"{clinic_name.replace(' Clinic', '')} Laptops",
        clinic_name.replace(" Clinic", "")
    ]
    for c in candidates:
        if c in existing_sheets:
            return c
    return clinic_name


def hex_to_rgb(hex_str):
    h = hex_str.lstrip('#')
    return {
        'red': int(h[0:2], 16) / 255.0,
        'green': int(h[2:4], 16) / 255.0,
        'blue': int(h[4:6], 16) / 255.0
    }

def apply_google_sheet_formatting(service, spreadsheet_id):
    """
    Applies executive Navy header styling, column widths, freeze panes,
    and pastel conditional formatting across all sheets:
      - Refresh Year 2026 (Due this year): Soft Red (#FCE4D6)
      - Refresh Year < 2026 (Past due): Lavender (#E1D5E7)
      - Uptime > 90 days (Really bad): Lavender (#E1D5E7)
      - Uptime 31-90 days: Soft Red (#FCE4D6)
      - Uptime 8-30 days: Soft Yellow (#FFF2CC)
    """
    ss = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = ss.get('sheets', [])

    c_lavender_bg = hex_to_rgb('E1D5E7')
    c_lavender_fg = hex_to_rgb('5B2C6F')
    c_red_bg      = hex_to_rgb('FCE4D6')
    c_red_fg      = hex_to_rgb('C00000')
    c_yellow_bg   = hex_to_rgb('FFF2CC')
    c_yellow_fg   = hex_to_rgb('B25900')
    c_navy_bg     = hex_to_rgb('1F4E79')
    c_white_fg    = hex_to_rgb('FFFFFF')

    col_widths = [150, 180, 150, 180, 140, 150, 120]
    requests = []

    for s in sheets:
        s_id = s['properties']['sheetId']

        # Clear existing conditional formats
        for idx in reversed(range(len(s.get('conditionalFormats', [])))):
            requests.append({
                'deleteConditionalFormatRule': {
                    'sheetId': s_id,
                    'index': idx
                }
            })

        # Reset all data rows (wipe any lingering manual cell background or font colors)
        requests.append({
            'repeatCell': {
                'range': {'sheetId': s_id, 'startRowIndex': 1, 'endRowIndex': 1000, 'startColumnIndex': 0, 'endColumnIndex': 10},
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColorStyle': {'rgbColor': {'red': 1.0, 'green': 1.0, 'blue': 1.0}},
                        'textFormat': {
                            'foregroundColorStyle': {'rgbColor': {'red': 0.12, 'green': 0.16, 'blue': 0.22}},
                            'bold': False, 'italic': False, 'fontSize': 10
                        },
                        'horizontalAlignment': 'LEFT',
                        'verticalAlignment': 'MIDDLE'
                    }
                },
                'fields': 'userEnteredFormat(backgroundColorStyle,textFormat,horizontalAlignment,verticalAlignment)'
            }
        })

        # Header formatting
        requests.append({
            'repeatCell': {
                'range': {'sheetId': s_id, 'startRowIndex': 0, 'endRowIndex': 1, 'startColumnIndex': 0, 'endColumnIndex': 7},
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColorStyle': {'rgbColor': c_navy_bg},
                        'textFormat': {'foregroundColorStyle': {'rgbColor': c_white_fg}, 'bold': True, 'fontSize': 10},
                        'horizontalAlignment': 'CENTER',
                        'verticalAlignment': 'MIDDLE'
                    }
                },
                'fields': 'userEnteredFormat(backgroundColorStyle,textFormat,horizontalAlignment,verticalAlignment)'
            }
        })

        # Align center for Uptime Days (col 5) and Refresh Year (col 6)
        requests.append({
            'repeatCell': {
                'range': {'sheetId': s_id, 'startRowIndex': 1, 'startColumnIndex': 5, 'endColumnIndex': 7},
                'cell': {'userEnteredFormat': {'horizontalAlignment': 'CENTER', 'verticalAlignment': 'MIDDLE'}},
                'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment)'
            }
        })

        # Freeze header row
        requests.append({
            'updateSheetProperties': {
                'properties': {'sheetId': s_id, 'gridProperties': {'frozenRowCount': 1}},
                'fields': 'gridProperties.frozenRowCount'
            }
        })

        # Column widths
        for c_idx, width in enumerate(col_widths):
            requests.append({
                'updateDimensionProperties': {
                    'range': {'sheetId': s_id, 'dimension': 'COLUMNS', 'startIndex': c_idx, 'endIndex': c_idx + 1},
                    'properties': {'pixelSize': width},
                    'fields': 'pixelSize'
                }
            })

        # Conditional Rules:
        # 1. Past Refresh Years (< 2026) -> Critical Lavender
        requests.append({
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': [{'sheetId': s_id, 'startRowIndex': 1, 'startColumnIndex': 6, 'endColumnIndex': 7}],
                    'booleanRule': {
                        'condition': {'type': 'NUMBER_LESS', 'values': [{'userEnteredValue': '2026'}]},
                        'format': {'backgroundColorStyle': {'rgbColor': c_lavender_bg}, 'textFormat': {'foregroundColorStyle': {'rgbColor': c_lavender_fg}, 'bold': True}}
                    }
                },
                'index': 0
            }
        })
        # 2. Refresh Year 2026 -> Soft Red
        requests.append({
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': [{'sheetId': s_id, 'startRowIndex': 1, 'startColumnIndex': 6, 'endColumnIndex': 7}],
                    'booleanRule': {
                        'condition': {'type': 'NUMBER_EQ', 'values': [{'userEnteredValue': '2026'}]},
                        'format': {'backgroundColorStyle': {'rgbColor': c_red_bg}, 'textFormat': {'foregroundColorStyle': {'rgbColor': c_red_fg}, 'bold': True}}
                    }
                },
                'index': 1
            }
        })
        # 3. Uptime > 90 days (Really bad) -> Critical Lavender
        requests.append({
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': [{'sheetId': s_id, 'startRowIndex': 1, 'startColumnIndex': 5, 'endColumnIndex': 6}],
                    'booleanRule': {
                        'condition': {'type': 'NUMBER_GREATER', 'values': [{'userEnteredValue': '90'}]},
                        'format': {'backgroundColorStyle': {'rgbColor': c_lavender_bg}, 'textFormat': {'foregroundColorStyle': {'rgbColor': c_lavender_fg}, 'bold': True}}
                    }
                },
                'index': 2
            }
        })
        # 4. Uptime 31-90 days -> Soft Red
        requests.append({
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': [{'sheetId': s_id, 'startRowIndex': 1, 'startColumnIndex': 5, 'endColumnIndex': 6}],
                    'booleanRule': {
                        'condition': {'type': 'NUMBER_GREATER', 'values': [{'userEnteredValue': '30'}]},
                        'format': {'backgroundColorStyle': {'rgbColor': c_red_bg}, 'textFormat': {'foregroundColorStyle': {'rgbColor': c_red_fg}, 'bold': True}}
                    }
                },
                'index': 3
            }
        })
        # 5. Uptime 8-30 days -> Soft Yellow
        requests.append({
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': [{'sheetId': s_id, 'startRowIndex': 1, 'startColumnIndex': 5, 'endColumnIndex': 6}],
                    'booleanRule': {
                        'condition': {'type': 'NUMBER_GREATER', 'values': [{'userEnteredValue': '7'}]},
                        'format': {'backgroundColorStyle': {'rgbColor': c_yellow_bg}, 'textFormat': {'foregroundColorStyle': {'rgbColor': c_yellow_fg}, 'bold': True}}
                    }
                },
                'index': 4
            }
        })

    # Execute in chunks of 100
    for i in range(0, len(requests), 100):
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': requests[i:i + 100]}
        ).execute()


def is_excluded_location(location_name):
    """
    Excludes administrative and staging locations:
    - Administration-Airport
    - Administration-Plaza
    - newWindowsWorkstation
    """
    if not location_name:
        return True
    loc_clean = str(location_name).strip().lower()
    if "admin" in loc_clean or "newwindowsworkstation" in loc_clean or loc_clean in ["unassigned", "staging"]:
        return True
    return False


def get_config_path():
    """Returns the path to google_sheets_config.json in the current working dir or repo root."""
    for d in [os.getcwd(), os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))]:
        p = os.path.join(d, CONFIG_FILE_NAME)
        if os.path.exists(p):
            return p
    return os.path.join(os.getcwd(), CONFIG_FILE_NAME)


def load_google_sheet_config():
    """Loads Google Sheets configuration (such as saved spreadsheet_id)."""
    cfg_path = get_config_path()
    data = {}
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}
    if not data.get("spreadsheet_id"):
        data["spreadsheet_id"] = DEFAULT_SPREADSHEET_ID
    return data


def save_google_sheet_config(data):
    """Saves Google Sheets configuration."""
    cfg_path = get_config_path()
    try:
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        return True
    except Exception as e:
        print(f"[Warning] Could not save config to {cfg_path}: {e}")
        return False


def find_latest_inventory_report():
    """
    Finds the newest generated Excel report (Extensive Report or Ninja Report)
    that contains a 'Ninja Inventory' worksheet.
    """
    search_dirs = [
        os.path.join(os.getcwd(), "reports", "Extensive Reports"),
        os.path.join(os.getcwd(), "reports", "Ninja Reports"),
        os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "reports", "Extensive Reports")),
        os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "reports", "Ninja Reports")),
        r"E:\Humble Mumble CSVS\reports\Extensive Reports",
        r"E:\Humble Mumble CSVS\reports\Ninja Reports",
    ]

    candidates = []
    for d in search_dirs:
        if os.path.exists(d):
            for f in glob.glob(os.path.join(d, "*.xlsx")):
                if not os.path.basename(f).startswith("~$"):
                    try:
                        candidates.append((os.path.getmtime(f), f))
                    except OSError:
                        pass

    candidates.sort(key=lambda x: x[0], reverse=True)
    for _, fpath in candidates:
        try:
            wb = openpyxl.load_workbook(fpath, read_only=True)
            if "Ninja Inventory" in wb.sheetnames:
                wb.close()
                return fpath
            wb.close()
        except Exception:
            continue
    return None


def extract_fleet_from_excel(excel_path, laptops_only=True, exclude_admin=True):
    """
    Extracts the 7 target columns from an existing Excel report's 'Ninja Inventory' tab,
    filtered to clinic laptops only by default.
    """
    wb = openpyxl.load_workbook(excel_path, read_only=True)
    ws = wb["Ninja Inventory"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        raise ValueError(f"No rows found in {excel_path} sheet 'Ninja Inventory'")

    header = [str(c).strip() if c else "" for c in rows[0]]
    indices = []
    for col_name in FLEET_HEADERS:
        if col_name in header:
            indices.append(header.index(col_name))
        else:
            found = -1
            for idx, h in enumerate(header):
                if col_name.lower() in h.lower():
                    found = idx
                    break
            indices.append(found)

    data_rows = []
    locations = set()
    for r in rows[1:]:
        if not r or not any(r):
            continue
        vals = []
        for idx in indices:
            val = r[idx] if idx != -1 and idx < len(r) else ""
            vals.append(val if val is not None else "")
        if vals[0]:
            dev_name = str(vals[0]).strip()
            loc = str(vals[1]).strip() if vals[1] else "Unassigned"
            role = str(vals[2]).strip()
            model = str(vals[3]).strip()

            if laptops_only and not is_laptop_device(role, model, dev_name):
                continue

            if exclude_admin and is_excluded_location(loc):
                continue

            data_rows.append(vals)
            locations.add(loc)

    sorted_locations = sorted(list(locations), key=lambda s: str(s).lower())
    return FLEET_HEADERS, data_rows, sorted_locations


def extract_fleet_from_raw_csvs(laptops_only=True, exclude_admin=True):
    """
    Extracts the 7 target columns directly from raw NinjaOne CSVs (details + custom fields).
    """
    best = find_best_ninja_csvs()
    csv_details = best.get("details_file")
    csv_custom = best.get("custom_file")

    if not csv_details or not os.path.exists(csv_details):
        raise FileNotFoundError("Could not find a valid NinjaOne Device Details CSV export.")

    ad_dates = {}
    if csv_custom and os.path.exists(csv_custom):
        with open(csv_custom, mode="r", encoding="utf-8-sig", errors="ignore") as f:
            reader = csv.reader(f)
            header = [c.strip() for c in next(reader, [])]
            dev_idx = header.index("Device") if "Device" in header else -1
            if dev_idx == -1 and "Display Name" in header:
                dev_idx = header.index("Display Name")
            ad_idx = header.index("AD Install Date") if "AD Install Date" in header else -1
            if ad_idx == -1:
                for idx, col in enumerate(header):
                    if "install date" in col.lower():
                        ad_idx = idx
                        break
            for row in reader:
                if row and dev_idx != -1 and ad_idx != -1 and len(row) > max(dev_idx, ad_idx):
                    dev_name = row[dev_idx].strip()
                    ad_val = row[ad_idx].strip()
                    if ad_val:
                        parsed = parse_ninja_date(ad_val)
                        if parsed:
                            ad_dates[dev_name.upper()] = parsed
                            ad_dates[normalize_hostname(dev_name)] = parsed

    devices = []
    parsed_online_dates = []

    with open(csv_details, mode="r", encoding="utf-8-sig", errors="ignore") as f:
        reader = csv.reader(f)
        header = [c.strip() for c in next(reader, [])]
        col_map = {name: i for i, name in enumerate(header)}

        for row in reader:
            if not row or not any(row):
                continue
            def get_val(col_name):
                idx = col_map.get(col_name)
                return row[idx].strip() if idx is not None and idx < len(row) else ""

            dev_name = get_val("Display Name") or get_val("Device") or get_val("Device Name") or get_val("Name")
            if not dev_name:
                continue

            role = get_val("Device Role")
            model = get_val("Device Model")
            loc = get_val("Location") or "Unassigned"

            if laptops_only and not is_laptop_device(role, model, dev_name):
                continue

            if exclude_admin and is_excluded_location(loc):
                continue

            last_online_str = get_val("Last Online_formatted")
            if last_online_str:
                p_date = parse_ninja_date(last_online_str)
                if p_date:
                    parsed_online_dates.append(p_date)

            devices.append({
                "Device": dev_name,
                "Location": loc,
                "Device Role": role,
                "Device Model": model,
                "Serial Number": get_val("Serial Number"),
                "Last Uptime_formatted": last_online_str,
                "Warranty End Date_formatted": get_val("Warranty End Date_formatted")
            })

    today_date = datetime.now().date()
    max_device_date = max(parsed_online_dates) if parsed_online_dates else today_date
    ref_date = max(today_date, max_device_date)

    data_rows = []
    locations = set()
    for d in devices:
        dev_name = d["Device"]
        up_date = parse_ninja_date(d["Last Uptime_formatted"])
        up_days = max(0, (ref_date - up_date).days) if up_date else ""

        ad_date = ad_dates.get(dev_name.upper()) or ad_dates.get(normalize_hostname(dev_name)) or ad_dates.get(dev_name)
        if ad_date:
            rf_year = ad_date.year + 3
        else:
            w_end_p = parse_ninja_date(d["Warranty End Date_formatted"])
            rf_year = (w_end_p.year + 3) if w_end_p else ""

        loc = d["Location"]
        locations.add(loc)
        data_rows.append([
            dev_name,
            loc,
            d["Device Role"],
            d["Device Model"],
            d["Serial Number"],
            up_days,
            rf_year
        ])

    sorted_locations = sorted(list(locations), key=lambda s: str(s).lower())
    return FLEET_HEADERS, data_rows, sorted_locations


def get_fleet_data(preferred_excel=None, laptops_only=True, exclude_admin=True):
    """
    Extracts the 7 fleet columns from the best available source (newest report or raw CSVs).
    """
    if preferred_excel and os.path.exists(preferred_excel):
        return extract_fleet_from_excel(preferred_excel, laptops_only=laptops_only, exclude_admin=exclude_admin)

    latest_report = find_latest_inventory_report()
    if latest_report:
        try:
            return extract_fleet_from_excel(latest_report, laptops_only=laptops_only, exclude_admin=exclude_admin)
        except Exception as e:
            print(f"[Notice] Could not parse {latest_report}: {e}. Falling back to raw CSVs...")

    return extract_fleet_from_raw_csvs(laptops_only=laptops_only, exclude_admin=exclude_admin)


def generate_google_sheet_template(output_path=None, data_source=None, laptops_only=True, exclude_admin=True, include_all_tab=False):
    """
    Builds the multi-tab Google Sheet template workbook:
      - Individual Clinic Tabs for each clinic with static, sortable laptop rows
      - (Optional) 'All Laptops' master tab if include_all_tab=True
    """
    headers, rows, locations = get_fleet_data(data_source, laptops_only=laptops_only, exclude_admin=exclude_admin)

    master_tab_title = "All Laptops" if laptops_only else "All Devices"

    if not output_path:
        out_dir = os.path.join(os.getcwd(), "reports", "Google Sheets")
        os.makedirs(out_dir, exist_ok=True)
        today_str = datetime.now().strftime("%m%d%Y")
        output_path = os.path.join(out_dir, f"Ninja_Clinic_Laptop_Inventory_{today_str}.xlsx")
    else:
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    wb = openpyxl.Workbook()

    # Styling Palette
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='E5E7EB'),
        right=Side(style='thin', color='E5E7EB'),
        top=Side(style='thin', color='E5E7EB'),
        bottom=Side(style='thin', color='E5E7EB')
    )
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    col_widths = {
        'A': 20,  # Device
        'B': 26,  # Location
        'C': 20,  # Device Role
        'D': 24,  # Device Model
        'E': 18,  # Serial Number
        'F': 18,  # Last Uptime - Days
        'G': 14   # Refresh Year
    }

    def populate_sheet(ws, title, sheet_rows):
        ws.title = title
        ws.append(headers)
        ws.row_dimensions[1].height = 26
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        for r_idx, row_vals in enumerate(sheet_rows, start=2):
            ws.append(row_vals)
            ws.row_dimensions[r_idx].height = 20
            use_zebra = (r_idx % 2 == 1)
            for c_idx in range(1, len(row_vals) + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.border = thin_border
                if use_zebra:
                    cell.fill = zebra_fill
                if c_idx in [6, 7]:  # Days, Year
                    cell.alignment = center_align
                else:
                    cell.alignment = left_align

        ws.freeze_panes = "A2"
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width
        max_r = max(len(sheet_rows) + 1, 2)
        ws.auto_filter.ref = f"A1:G{max_r}"

    is_first = True
    if include_all_tab:
        populate_sheet(wb.active, master_tab_title, rows)
        is_first = False

    for loc in locations:
        clean_title = loc[:31]
        for ch in [":", "\\", "/", "?", "*", "[", "]"]:
            clean_title = clean_title.replace(ch, "")
        clinic_rows = [r for r in rows if r[1] == loc]
        clinic_rows.sort(key=lambda x: str(x[0]).lower())

        if is_first:
            ws_clinic = wb.active
            is_first = False
        else:
            ws_clinic = wb.create_sheet(title=clean_title)

        populate_sheet(ws_clinic, clean_title, clinic_rows)

    wb.save(output_path)
    print(f"[Success] Generated Google Sheets template with {len(rows)} clinic laptops across {len(locations)} clinic tabs:")
    print(f"          -> {output_path}")
    return output_path, len(rows), len(locations)


def copy_fleet_to_clipboard(include_header=False, data_source=None, laptops_only=True, exclude_admin=True):
    """
    Formats the 7 fleet columns as Tab-Separated Values (TSV) and copies them
    to the Windows clipboard for instant pasting into cell A2 of 'All Laptops'.
    """
    headers, rows, _ = get_fleet_data(data_source, laptops_only=laptops_only, exclude_admin=exclude_admin)

    lines = []
    if include_header:
        lines.append("\t".join(headers))
    for r in rows:
        lines.append("\t".join(str(c) if c is not None else "" for c in r))

    tsv_data = "\r\n".join(lines)

    # Copy to Windows Clipboard using clip.exe
    try:
        p = subprocess.Popen(["clip.exe"], stdin=subprocess.PIPE, shell=True)
        p.communicate(input=tsv_data.encode("utf-8"))
        if p.returncode == 0:
            print(f"[Success] Copied {len(rows)} clinic laptop records directly to your Windows clipboard!")
            return len(rows)
    except Exception as e:
        print(f"[Notice] clip.exe failed: {e}. Trying PowerShell Set-Clipboard...")

    # Fallback to PowerShell Set-Clipboard
    try:
        ps_cmd = "Set-Clipboard -Value ([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($args[0])))"
        import base64
        b64 = base64.b64encode(tsv_data.encode("utf-8")).decode("ascii")
        subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd, b64], check=True)
        print(f"[Success] Copied {len(rows)} clinic laptop records to clipboard via PowerShell!")
        return len(rows)
    except Exception as e:
        print(f"[Error] Failed to copy to clipboard: {e}")
        return 0


def export_master_csv(output_path=None, data_source=None, laptops_only=True, exclude_admin=True):
    """
    Exports the 7 columns of all clinic laptops to a clean CSV file
    for manual Google Sheets import (File -> Import -> Replace current sheet).
    """
    headers, rows, _ = get_fleet_data(data_source, laptops_only=laptops_only, exclude_admin=exclude_admin)

    if not output_path:
        out_dir = os.path.join(os.getcwd(), "reports", "Google Sheets")
        os.makedirs(out_dir, exist_ok=True)
        output_path = os.path.join(out_dir, "Ninja_Clinic_Laptop_Master.csv")

    with open(output_path, mode="w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)

    print(f"[Success] Exported master clinic laptop CSV to: {output_path}")
    return output_path, len(rows)


def sync_to_google_sheet_api(spreadsheet_id, data_source=None, credentials_path=None, token_path=None, laptops_only=True, exclude_admin=True, include_all_tab=False):
    """
    Directly updates 'All Laptops'!A2:G in the user's existing shared Google Sheet
    using the Google Sheets API.
    """
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

    search_dirs = [
        os.getcwd(),
        os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..")),
        r"E:\Humble Mumble CSVS"
    ]

    if not credentials_path:
        for d in search_dirs:
            p = os.path.join(d, "credentials.json")
            if os.path.exists(p):
                credentials_path = p
                break

    if not token_path:
        for d in search_dirs:
            p = os.path.join(d, "token_sheets.json")
            if os.path.exists(p):
                token_path = p
                break
        if not token_path:
            token_path = os.path.join(os.getcwd(), "token_sheets.json")

    creds = None
    if os.path.exists(token_path):
        try:
            creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        except Exception:
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not credentials_path or not os.path.exists(credentials_path):
                raise FileNotFoundError(
                    "Google OAuth credentials.json not found. Please place credentials.json in project root."
                )
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w", encoding="utf-8") as f:
            f.write(creds.to_json())

    service = build("sheets", "v4", credentials=creds)

    headers, rows, locations = get_fleet_data(data_source, laptops_only=laptops_only, exclude_admin=exclude_admin)
    master_tab_title = "All Laptops" if laptops_only else "All Devices"

    # Inspect current sheets in the spreadsheet
    try:
        ss_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        existing_sheets = {s["properties"]["title"]: s["properties"]["sheetId"] for s in ss_metadata.get("sheets", [])}
    except Exception as e:
        print(f"[Error] Could not read spreadsheet metadata: {e}")
        existing_sheets = {}

    batch_requests = []

    # 1. Manage master tab and clinic tabs
    if not include_all_tab:
        # User requested clinic tabs only: delete 'All Laptops' / 'All Devices' if present
        for old_t in ["All Laptops", "All Devices"]:
            if old_t in existing_sheets and len(existing_sheets) > 1:
                print(f"[*] Removing '{old_t}' tab as requested...")
                batch_requests.append({
                    "deleteSheet": {
                        "sheetId": existing_sheets[old_t]
                    }
                })
                del existing_sheets[old_t]
    else:
        # Ensure master tab exists
        if master_tab_title not in existing_sheets:
            print(f"[*] Adding '{master_tab_title}' tab...")
            batch_requests.append({
                "addSheet": {
                    "properties": {
                        "title": master_tab_title
                    }
                }
            })

    # 2. Ensure each clinic tab exists
    for loc in locations:
        matched_tab = get_clinic_tab_name(existing_sheets, loc)
        if matched_tab not in existing_sheets:
            clean_loc = loc[:31]
            for ch in [":", "\\", "/", "?", "*", "[", "]"]:
                clean_loc = clean_loc.replace(ch, "")
            print(f"[*] Adding clinic tab: '{clean_loc}'...")
            batch_requests.append({
                "addSheet": {
                    "properties": {
                        "title": clean_loc
                    }
                }
            })
            existing_sheets[clean_loc] = True

    if batch_requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": batch_requests}
        ).execute()

    # 3. Clear existing data rows across clinic tabs (and master tab if included)
    clear_ranges = []
    if include_all_tab:
        clear_ranges.append(f"'{master_tab_title}'!A2:G1000")
    for loc in locations:
        target_tab = get_clinic_tab_name(existing_sheets, loc)
        clear_ranges.append(f"'{target_tab}'!A2:G1000")

    try:
        service.spreadsheets().values().batchClear(
            spreadsheetId=spreadsheet_id,
            body={"ranges": clear_ranges}
        ).execute()
        print(f"[*] Cleared previous data rows across {len(clear_ranges)} sheets.")
    except Exception as e:
        print(f"[Notice] batchClear notice: {e}")

    # 4. Prepare data updates with static rows for master tab and every clinic tab
    value_updates = []

    # Master Tab (if included)
    if include_all_tab:
        value_updates.append({
            "range": f"'{master_tab_title}'!A1:G1",
            "values": [headers]
        })
        value_updates.append({
            "range": f"'{master_tab_title}'!A2:G{len(rows) + 1}",
            "values": rows
        })

    # Individual Clinic Tabs (Direct static values - fully sortable without breaking!)
    for loc in locations:
        target_tab = get_clinic_tab_name(existing_sheets, loc)
        clinic_rows = [r for r in rows if r[1] == loc]
        clinic_rows.sort(key=lambda x: str(x[0]).lower())

        value_updates.append({
            "range": f"'{target_tab}'!A1:G1",
            "values": [headers]
        })
        if clinic_rows:
            value_updates.append({
                "range": f"'{target_tab}'!A2:G{len(clinic_rows) + 1}",
                "values": clinic_rows
            })

    # 5. Execute batchUpdate for values
    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": value_updates
        }
    ).execute()

    # 6. Apply Navy headers, freeze panes, column widths, and 5 pastel conditional formatting rules
    print("[*] Applying styling, freeze panes, and conditional formatting rules...")
    apply_google_sheet_formatting(service, spreadsheet_id)

    print(f"[Success] Directly updated {len(rows)} clinic laptops in Google Sheet (ID: {spreadsheet_id}).")
    print(f"          Populated {len(locations)} clinic tabs with static sortable rows. No formulas to break!")

    # Save spreadsheet_id in config for reuse
    cfg = load_google_sheet_config()
    cfg["spreadsheet_id"] = spreadsheet_id
    save_google_sheet_config(cfg)

    return len(rows)


def extract_spreadsheet_id(input_str):
    """Extracts spreadsheet ID from a full Google Sheets URL or raw ID string."""
    if not input_str:
        return ""
    s = str(input_str).strip()
    if "/d/" in s:
        part = s.split("/d/")[1]
        if "/" in part:
            return part.split("/")[0]
        return part
    return s


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Google Sheets Clinic Laptop Reporting for NinjaOne")
    parser.add_argument("--template", action="store_true", help="Generate the multi-tab Excel template for Google Sheets")
    parser.add_argument("--copy", action="store_true", help="Copy the 7-column clinic laptop data directly to Windows clipboard")
    parser.add_argument("--csv", action="store_true", help="Export clean CSV for Google Sheets import")
    parser.add_argument("--sync", nargs="?", const="default", default=None, metavar="SPREADSHEET_ID_OR_URL", help="Sync directly to Google Sheet via API (uses configured spreadsheet if omitted)")
    parser.add_argument("--source", metavar="FILE", help="Specific Excel report or folder to read from")
    parser.add_argument("--all-devices", action="store_true", help="Include all devices instead of laptops only")
    parser.add_argument("--include-admin", action="store_true", help="Include admin and staging locations")
    parser.add_argument("--include-all", action="store_true", help="Include 'All Laptops' master tab (default is False, clinic tabs only)")

    args = parser.parse_args()
    laptops_only = not args.all_devices
    exclude_admin = not args.include_admin

    if not (args.template or args.copy or args.csv or (args.sync is not None)):
        cfg = load_google_sheet_config()
        sheet_id = cfg.get("spreadsheet_id", DEFAULT_SPREADSHEET_ID)
        print("=" * 70)
        print("  GOOGLE SHEETS CLINIC LAPTOP FLEET SYNC — NINJAONE")
        print(f"  Target: https://docs.google.com/spreadsheets/d/{sheet_id}")
        print("=" * 70)
        print("[*] Running default action: Sync Directly to Google Sheets via API...")
        sync_to_google_sheet_api(
            sheet_id,
            data_source=args.source,
            laptops_only=laptops_only,
            exclude_admin=exclude_admin,
            include_all_tab=args.include_all
        )
        return

    if args.template:
        generate_google_sheet_template(
            data_source=args.source,
            laptops_only=laptops_only,
            exclude_admin=exclude_admin,
            include_all_tab=args.include_all
        )

    if args.csv:
        export_master_csv(
            data_source=args.source,
            laptops_only=laptops_only,
            exclude_admin=exclude_admin
        )

    if args.copy:
        copy_fleet_to_clipboard(
            include_header=False,
            data_source=args.source,
            laptops_only=laptops_only,
            exclude_admin=exclude_admin
        )

    if args.sync is not None:
        sheet_id = ""
        if isinstance(args.sync, str) and args.sync.lower() not in ["default", "none", ""]:
            sheet_id = extract_spreadsheet_id(args.sync)
        if not sheet_id:
            cfg = load_google_sheet_config()
            sheet_id = cfg.get("spreadsheet_id", DEFAULT_SPREADSHEET_ID)

        sync_to_google_sheet_api(
            sheet_id,
            data_source=args.source,
            laptops_only=laptops_only,
            exclude_admin=exclude_admin,
            include_all_tab=args.include_all
        )


if __name__ == "__main__":
    main()
