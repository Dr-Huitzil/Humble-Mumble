"""
build_ad_report.py
==================
Generates an executive-grade Excel report from the Active Directory raw CSV export.
Worksheet order:
  Tab 1: AD Computers (Master Inventory Table with Auto-Filters, Clean Typography & Styling)
  Tab 2: Executive Summary & Audit Dashboard (KPIs, Clinic Breakdown, Stale/Unassigned Alerts)
"""

import os
import sys
import csv
import glob
import re
from datetime import datetime
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

import config
from utils import clean_ou_token, create_kpi_card

def parse_ad_record(row):
    name = str(row.get("Name", "") or "").strip()
    status = str(row.get("Status", "") or "").strip()
    c1_raw = str(row.get("Column1", "") or "").strip()
    c2_raw = str(row.get("Column2", "") or "").strip()
    c3_raw = str(row.get("Column3", "") or "").strip()
    os_name = str(row.get("OperatingSystem", "") or "").strip()
    os_ver = str(row.get("OperatingSystemVersion", "") or "").strip()
    
    t1 = clean_ou_token(c1_raw)
    t2 = clean_ou_token(c2_raw)
    t3 = clean_ou_token(c3_raw)
    
    # Determine Clinic / Site, Department, and Device Type
    clinic = "Unknown"
    dept = ""
    device_type = "Desktop"
    
    # Check Server
    if "SERVER" in os_name.upper() or "SERVERS" in (c1_raw + c2_raw).upper():
        device_type = "Server"
        clinic = "Servers / Infrastructure"
    elif "Domain Controllers" in c1_raw:
        device_type = "Domain Controller"
        clinic = "Domain Controllers"
    elif "CN=Computers" in c1_raw:
        clinic = "Default Container (Unassigned)"
        device_type = "Laptop" if re.search(r'(LAPTOP|LAP|LPT|-L\d+|-L-)', name.upper()) else "Desktop"
    elif "Disabled Computers" in c1_raw:
        prefix = name.upper().split('-')[0]
        mapped = config.CLINIC_PREFIX_MAP.get(prefix, "Disabled Computers")
        clinic = config.canonicalize_clinic_name(mapped)
        device_type = "Laptop" if re.search(r'(LAPTOP|LAP|LPT|-L\d+|-L-)', name.upper()) else "Desktop"
    else:
        # Standard or Department OU
        if t3.upper() in ["ADMIN", "CORPORATE", "HQ", "HEADQUARTERS"]:
            clinic = config.canonicalize_clinic_name(t3.upper())
            dept = t2
        elif t3.upper() in ["ALL CLINICS", "ALL SITES", "LOCATIONS", "BRANCHES"]:
            clinic = config.canonicalize_clinic_name(t2.upper())
        else:
            raw_c = t2.upper() if t2 else t1.upper()
            clinic = config.canonicalize_clinic_name(raw_c)
            
        if "LAPTOP" in (c1_raw + name).upper() or re.search(r'(LAPTOP|LAP|LPT|-L\d+|-L-)', name.upper()):
            device_type = "Laptop"
        else:
            device_type = "Desktop"
            
    # Parse days as integers if possible
    def to_int_or_blank(val):
        try:
            return int(val.strip())
        except (ValueError, AttributeError):
            return ""

    return {
        "Name": name,
        "Status": status,
        "Clinic": clinic,
        "Department": dept,
        "Device Type": device_type,
        "OperatingSystem": os_name,
        "OperatingSystemVersion": os_ver,
        "LastLogonDate": row.get("LastLogonDate", ""),
        "Last Logon - Days": to_int_or_blank(row.get("Last Logon - Days", "")),
        "CreatedDate": row.get("CreatedDate", ""),
        "Created - Days": to_int_or_blank(row.get("Created - Days", "")),
        "ModifiedDate": row.get("ModifiedDate", ""),
        "Modified- Days": to_int_or_blank(row.get("Modified- Days", "")),
        "OU_Token1": t1,
        "OU_Token2": t2,
        "OU_Token3": t3
    }

def parse_ad_csv(csv_path):
    records = []
    with open(csv_path, mode="r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for r in reader:
            records.append(parse_ad_record(r))
    return records

def build_ad_inventory_tab(wb, records, title="AD Computers"):
    """
    Builds the AD master inventory tab into the provided workbook.
    """
    if title in wb.sheetnames:
        ws_master = wb[title]
    elif len(wb.sheetnames) == 1 and wb.active.title == "Sheet":
        ws_master = wb.active
        ws_master.title = title
    else:
        ws_master = wb.create_sheet(title=title)
        
    ws_master.views.sheetView[0].showGridLines = True
    master_headers = config.AD_HEADERS
    
    for col_idx, h in enumerate(master_headers, start=1):
        cell = ws_master.cell(row=1, column=col_idx, value=h)
        cell.font = config.FONT_TBL_HEADER
        cell.fill = config.FILL_NAVY_HEADER
        cell.alignment = config.ALIGN_HEADER
        cell.border = config.BORDER_THIN
        
    for r_idx, r in enumerate(records, start=2):
        row_vals = [
            r["Name"], r["Status"], r["Clinic"], r["Department"], r["Device Type"],
            r["OperatingSystem"], r["OperatingSystemVersion"], r["LastLogonDate"], r["Last Logon - Days"],
            r["CreatedDate"], r["Created - Days"], r["ModifiedDate"], r["Modified- Days"],
            r["OU_Token1"], r["OU_Token2"], r["OU_Token3"]
        ]
        fill = config.FILL_ZEBRA_ROW if (r_idx % 2 == 1) else config.FILL_WHITE_ROW
        
        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws_master.cell(row=r_idx, column=col_idx, value=val)
            cell.font = config.FONT_TBL_DATA
            cell.fill = fill
            cell.border = config.BORDER_THIN
            
            # Alignments & conditional styling
            if col_idx == 1:  # Name
                cell.font = config.FONT_TBL_BOLD
                cell.alignment = config.ALIGN_LEFT
            elif col_idx == 2: # Status
                cell.alignment = config.ALIGN_CENTER
                if val == "Enabled":
                    cell.fill = config.FILL_ALERT_GREEN
                    cell.font = config.FONT_ALERT_GREEN
                else:
                    cell.fill = config.FILL_ALERT_RED
                    cell.font = config.FONT_ALERT_RED
            elif col_idx in [8, 10, 12]: # Dates
                cell.alignment = config.ALIGN_CENTER
            elif col_idx in [9, 11, 13]: # Days
                cell.alignment = config.ALIGN_RIGHT
                if isinstance(val, int):
                    cell.number_format = '#,##0'
                    if col_idx == 9 and r["Status"] == "Enabled" and val > 90:
                        cell.fill = config.FILL_ALERT_WARN
                        cell.font = config.FONT_ALERT_WARN
            else:
                cell.alignment = config.ALIGN_LEFT
                
    # Freeze Header Row and Enable Auto-Filters
    ws_master.freeze_panes = "A2"
    ws_master.auto_filter.ref = f"A1:{get_column_letter(len(master_headers))}{len(records)+1}"
    
    # Set Master Column Widths
    for col_idx, width in config.AD_COL_WIDTHS.items():
        ws_master.column_dimensions[get_column_letter(col_idx)].width = width
    return ws_master

def build_ad_dashboard_tab(wb, records, title="Executive Summary"):
    """
    Builds the AD Executive Summary & Audit Dashboard tab into the provided workbook.
    """
    total_count = len(records)
    enabled_count = sum(1 for r in records if r["Status"] == "Enabled")
    disabled_count = total_count - enabled_count
    
    w11_count = sum(1 for r in records if "Windows 11" in r["OperatingSystem"])
    w11_pct = f"{(w11_count / total_count * 100):.1f}%" if total_count else "0%"
    
    stale_active = [r for r in records if r["Status"] == "Enabled" and isinstance(r["Last Logon - Days"], int) and r["Last Logon - Days"] > 90]
    unassigned = [r for r in records if r["Clinic"] == "Default Container (Unassigned)"]

    if title in wb.sheetnames:
        ws_summary = wb[title]
    else:
        ws_summary = wb.create_sheet(title=title)
        
    ws_summary.views.sheetView[0].showGridLines = True
    
    # Title Banner
    ws_summary.cell(row=2, column=2, value="ACTIVE DIRECTORY ENDPOINT AUDIT & INVENTORY").font = config.FONT_TITLE
    gen_time = datetime.now().strftime("%B %d, %Y")
    ws_summary.cell(row=3, column=2, value=f"Domain: {config.DOMAIN_NAME} | Total Discovered Endpoints: {total_count} | Report Date: {gen_time}").font = config.FONT_SUBTITLE
    
    # KPI Cards Row (Row 5-6)
    create_kpi_card(ws_summary, 2, 5, "TOTAL COMPUTER OBJECTS", total_count)
    create_kpi_card(ws_summary, 4, 5, "ACTIVE (ENABLED)", enabled_count)
    create_kpi_card(ws_summary, 6, 5, "DISABLED (ARCHIVE)", disabled_count)
    create_kpi_card(ws_summary, 8, 5, "WIN 11 ADOPTION", w11_pct)
    create_kpi_card(ws_summary, 10, 5, "STALE ACTIVE (>90d)", len(stale_active))
    create_kpi_card(ws_summary, 12, 5, "UNASSIGNED (CN=COMP)", len(unassigned))
    
    # Clinic Breakdown Table (Row 9)
    ws_summary.cell(row=9, column=2, value="SITE & WORKSTATION ALLOCATION").font = config.FONT_SECTION
    
    clinic_headers = ["Clinic / Location", "Total Devices", "Enabled", "Disabled", "Laptops", "Desktops", "Active %"]
    for col_idx, h in enumerate(clinic_headers, start=2):
        cell = ws_summary.cell(row=10, column=col_idx, value=h)
        cell.font = config.FONT_TBL_HEADER
        cell.fill = config.FILL_NAVY_HEADER
        cell.alignment = config.ALIGN_HEADER
        cell.border = config.BORDER_THIN
        
    # Aggregate by Clinic
    clinic_data = {}
    for r in records:
        c = r["Clinic"]
        if c not in clinic_data:
            clinic_data[c] = {"total": 0, "enabled": 0, "disabled": 0, "laptops": 0, "desktops": 0}
        clinic_data[c]["total"] += 1
        if r["Status"] == "Enabled":
            clinic_data[c]["enabled"] += 1
        else:
            clinic_data[c]["disabled"] += 1
        if r["Device Type"] == "Laptop":
            clinic_data[c]["laptops"] += 1
        elif r["Device Type"] == "Desktop":
            clinic_data[c]["desktops"] += 1
            
    cur_row = 11
    # Sort clinics alphabetically, putting special categories at end
    sorted_clinics = sorted(clinic_data.keys(), key=lambda x: (x.startswith("Disabled") or x.startswith("Default") or x.startswith("Servers"), x))
    for c_name in sorted_clinics:
        d = clinic_data[c_name]
        pct = f"{(d['enabled'] / d['total'] * 100):.0f}%" if d['total'] else "0%"
        row_vals = [c_name, d["total"], d["enabled"], d["disabled"], d["laptops"], d["desktops"], pct]
        fill = config.FILL_ZEBRA_ROW if (cur_row % 2 == 1) else config.FILL_WHITE_ROW
        
        for col_idx, val in enumerate(row_vals, start=2):
            cell = ws_summary.cell(row=cur_row, column=col_idx, value=val)
            cell.font = config.FONT_TBL_DATA
            cell.fill = fill
            cell.border = config.BORDER_THIN
            cell.alignment = config.ALIGN_LEFT if col_idx == 2 else config.ALIGN_CENTER
        cur_row += 1
        
    # Total Row
    total_pct = f"{(enabled_count / total_count * 100):.0f}%" if total_count else "0%"
    total_row_vals = ["TOTAL DOMAIN ASSETS", total_count, enabled_count, disabled_count, 
                      sum(d["laptops"] for d in clinic_data.values()), 
                      sum(d["desktops"] for d in clinic_data.values()),
                      total_pct]
    for col_idx, val in enumerate(total_row_vals, start=2):
        cell = ws_summary.cell(row=cur_row, column=col_idx, value=val)
        cell.font = config.FONT_TBL_TOTAL
        cell.fill = config.FILL_TOTAL_ROW
        cell.border = config.BORDER_TOTAL
        cell.alignment = config.ALIGN_LEFT if col_idx == 2 else config.ALIGN_CENTER
        
    # High-Priority Action Items (Row cur_row + 3)
    alert_row = cur_row + 3
    ws_summary.cell(row=alert_row, column=2, value="HIGH-PRIORITY HYGIENE & AUDIT TARGETS").font = config.FONT_SECTION
    alert_headers = ["Computer Name", "Status", "Clinic / Site", "Operating System", "Last Logon Date", "Days Offline", "Audit Concern"]
    for col_idx, h in enumerate(alert_headers, start=2):
        cell = ws_summary.cell(row=alert_row + 1, column=col_idx, value=h)
        cell.font = config.FONT_TBL_HEADER
        cell.fill = config.FILL_SUB_HEADER
        cell.alignment = config.ALIGN_HEADER
        cell.border = config.BORDER_THIN
        
    alert_cur = alert_row + 2
    # Add top stale enabled devices
    for st in sorted(stale_active, key=lambda x: -(x["Last Logon - Days"] if isinstance(x["Last Logon - Days"], int) else 0))[:10]:
        days = st["Last Logon - Days"]
        row_vals = [st["Name"], st["Status"], st["Clinic"], st["OperatingSystem"], st["LastLogonDate"], days, f"Active endpoint offline {days} days (Ghost asset)"]
        fill = config.FILL_ALERT_WARN
        for col_idx, val in enumerate(row_vals, start=2):
            cell = ws_summary.cell(row=alert_cur, column=col_idx, value=val)
            cell.font = config.FONT_ALERT_WARN if col_idx == 8 else config.FONT_TBL_DATA
            cell.fill = fill
            cell.border = config.BORDER_THIN
            cell.alignment = config.ALIGN_LEFT if col_idx in [2, 4, 8] else config.ALIGN_CENTER
        alert_cur += 1
        
    # Add unassigned devices in CN=Computers
    for un in unassigned[:10]:
        row_vals = [un["Name"], un["Status"], "Default Container", un["OperatingSystem"], un["LastLogonDate"], un["Last Logon - Days"], "In CN=Computers (Missing Site OU & GPOs)"]
        fill = config.FILL_ALERT_RED
        for col_idx, val in enumerate(row_vals, start=2):
            cell = ws_summary.cell(row=alert_cur, column=col_idx, value=val)
            cell.font = config.FONT_ALERT_RED if col_idx == 8 else config.FONT_TBL_DATA
            cell.fill = fill
            cell.border = config.BORDER_THIN
            cell.alignment = config.ALIGN_LEFT if col_idx in [2, 4, 8] else config.ALIGN_CENTER
        alert_cur += 1

    # Widths for Summary
    ws_summary.column_dimensions['A'].width = 4
    ws_summary.column_dimensions['B'].width = 30
    ws_summary.column_dimensions['C'].width = 16
    ws_summary.column_dimensions['D'].width = 16
    ws_summary.column_dimensions['E'].width = 16
    ws_summary.column_dimensions['F'].width = 16
    ws_summary.column_dimensions['G'].width = 16
    ws_summary.column_dimensions['H'].width = 42
    for col_c in ['I', 'J', 'K', 'L', 'M', 'N']:
        ws_summary.column_dimensions[col_c].width = 18

    return ws_summary

def build_report(csv_path, output_xlsx=None, inv_title="Inventory", dash_title="Dashboard"):
    if not output_xlsx:
        timestamp = datetime.now().strftime("%Y%m%d")
        output_xlsx = f"AD_Computer_Inventory_Report_{timestamp}.xlsx"
        
    records = parse_ad_csv(csv_path)
    wb = Workbook()
    build_ad_inventory_tab(wb, records, title=inv_title)
    build_ad_dashboard_tab(wb, records, title=dash_title)

    out_dir = os.path.dirname(os.path.abspath(output_xlsx))
    os.makedirs(out_dir, exist_ok=True)
    
    try:
        wb.save(output_xlsx)
        print(f"Report successfully generated!")
        print(f"Output File: {os.path.abspath(output_xlsx)}")
    except PermissionError:
        raise PermissionError(
            f"Cannot save '{os.path.basename(output_xlsx)}' because the file is currently open in Microsoft Excel or another program.\n"
            "Please close the workbook in Excel and re-run the report."
        )
    return output_xlsx

if __name__ == "__main__":
    target_csv = None
    if len(sys.argv) > 1:
        target_csv = sys.argv[1]
    else:
        raw_files = sorted(glob.glob("Computers_Raw_*.csv"), key=os.path.getmtime, reverse=True)
        if not raw_files:
            raw_files = sorted(glob.glob(os.path.join("data", "Computers_Raw_*.csv")) + glob.glob(os.path.join("archive", "Computers_Raw_*.csv")), key=os.path.getmtime, reverse=True)
        if raw_files:
            target_csv = raw_files[0]
            
    if not target_csv or not os.path.exists(target_csv):
        print("Error: No Computers_Raw CSV file found.")
        sys.exit(1)
        
    out_file = sys.argv[2] if len(sys.argv) > 2 else None
    print(f"Reading data from: {target_csv}")
    build_report(target_csv, output_xlsx=out_file)
