"""
generate_dashboard_tabs.py
===========================
Generates multi-tab executive dashboards, KPI summary cards, and native Excel
charts from master inventory data.

Can be called directly as an imported Python function:
    generate_dashboards(wb_path, huntress_csv=None, skip_lifecycle=False)

Or run from the command line:
    python generate_dashboard_tabs.py [path_to_workbook.xlsx] [--huntress path] [--no-ad]
"""

import sys
import os
import glob
from collections import Counter, defaultdict

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList

from humble_mumble.core import config
from humble_mumble.core.utils import draw_kpi_cards, is_valid_serial, normalize_hostname, safe_float, safe_int

# Use centralized styles from config
from humble_mumble.core.config import (
    font_title, font_subtitle, font_section, font_card_num, font_card_label, font_card_sub,
    font_tbl_header, font_tbl_data, font_tbl_bold, font_tbl_total,
    fill_navy_header, fill_sub_header, fill_accent_header, fill_kpi_card,
    fill_zebra, fill_white, fill_total,
    fill_red, font_red, fill_yellow, font_yellow, fill_green, font_green,
    fill_purple, font_purple, fill_staging, font_staging,
    thin_border, total_border, align_left, align_center, align_right, align_header
)

def create_tab(wb, sheet_name):
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws_new = wb.create_sheet(title=sheet_name)
    ws_new.views.sheetView[0].showGridLines = True
    return ws_new

def _load_inventory_data(master_ws):
    headers = {master_ws.cell(1, c).value: c for c in range(1, master_ws.max_column + 1)}
    c_dev = headers.get("Device", 1)
    c_loc = headers.get("Location", 2)
    c_role = headers.get("Device Role", 3)
    c_model = headers.get("Device Model", 4)
    c_serial = headers.get("Serial Number", 5)
    c_user = headers.get("Last Login", 6)
    c_uptime = headers.get("Last Uptime - Days", 8)
    c_boot = headers.get("Last Boot Time - Days", 10)
    c_update = headers.get("Last Update - Days", 14)
    c_ram = headers.get("RAM", 15)
    c_cap = headers.get("Primary Drive Capacity (GB)", 16)
    c_free = headers.get("Primary Drive Free (GB)", 17)
    c_usage = headers.get("Primary Drive Usage %", 18)
    c_rf = headers.get("Refresh Year", 22)
    c_w_status = headers.get("Warranty Status", 25)
    c_w_days = headers.get("Warranty Days Remaining", 26)
    c_os = headers.get("OS", 12)

    inventory_data = []
    for r in range(2, master_ws.max_row + 1):
        dev_val = master_ws.cell(r, c_dev).value
        if not dev_val:
            continue
        row_dict = {
            "Device": dev_val,
            "Location": str(master_ws.cell(r, c_loc).value or "Unassigned"),
            "Role": str(master_ws.cell(r, c_role).value or "Windows Desktop"),
            "Model": str(master_ws.cell(r, c_model).value or "Unknown"),
            "Serial": master_ws.cell(r, c_serial).value,
            "User": str(master_ws.cell(r, c_user).value or "N/A"),
            "Uptime": master_ws.cell(r, c_uptime).value or 0,
            "Boot": master_ws.cell(r, c_boot).value or 0,
            "Update": master_ws.cell(r, c_update).value or 0,
            "RAM": master_ws.cell(r, c_ram).value,
            "Capacity": master_ws.cell(r, c_cap).value,
            "Free": master_ws.cell(r, c_free).value,
            "Usage": master_ws.cell(r, c_usage).value,
            "Refresh": master_ws.cell(r, c_rf).value,
            "WarrantyStatus": str(master_ws.cell(r, c_w_status).value or "Unknown"),
            "WarrantyDays": master_ws.cell(r, c_w_days).value or 0,
            "OS": str(master_ws.cell(r, c_os).value or "Unknown")
        }
        inventory_data.append(row_dict)
    return inventory_data

def _compute_clinic_map(inventory_data, base_year=None):
    if base_year is None:
        valid_years = [d["Refresh"] for d in inventory_data if isinstance(d.get("Refresh"), int) and d["Refresh"] > 2000]
        base_year = min(valid_years) if valid_years else datetime.now().year
    years = [base_year, base_year + 1, base_year + 2, base_year + 3]

    clinic_map = defaultdict(lambda: {y: 0 for y in years} | {"Laptop": 0, "Desktop": 0, "Dental": 0, "Total": 0})
    for d in inventory_data:
        loc = config.canonicalize_clinic_name(d["Location"])
        yr = d["Refresh"]
        role = d["Role"]
        clinic_map[loc]["Total"] += 1
        if yr in years:
            clinic_map[loc][yr] += 1
        if "Laptop" in role:
            clinic_map[loc]["Laptop"] += 1
        elif "Dental" in role:
            clinic_map[loc]["Dental"] += 1
        else:
            clinic_map[loc]["Desktop"] += 1
    return clinic_map, years

def _build_lifecycle_tab(wb, inventory_data, total_fleet, clinic_map, years=None):
    print("Building Tab 1: Lifecycle...")
    if years is None:
        valid_years = [d["Refresh"] for d in inventory_data if isinstance(d.get("Refresh"), int) and d["Refresh"] > 2000]
        base_year = min(valid_years) if valid_years else datetime.now().year
        years = [base_year, base_year + 1, base_year + 2, base_year + 3]
    base_year = years[0]

    ws1 = create_tab(wb, "Lifecycle")
    ws1.cell(1, 1).value = "HARDWARE LIFECYCLE & REFRESH FORECAST"
    ws1.cell(1, 1).font = font_title
    ws1.cell(2, 1).value = f"Replacement Waves by Clinic (3-Year AD Lifecycle Policy) | {config.CURRENT_MONTH_YEAR}"
    ws1.cell(2, 1).font = font_subtitle

    # Counts
    rf_counts = Counter(d["Refresh"] for d in inventory_data)
    due_base = rf_counts.get(base_year, 0)
    w_counts = Counter(d["WarrantyStatus"] for d in inventory_data)
    w_expired = w_counts.get("Expired", 0)
    w_soon = w_counts.get("Expiring Soon", 0)

    exp_pct = f"{(w_expired/total_fleet)*100:.1f}% of Fleet" if total_fleet else "0%"
    kpis1 = [
        ("TOTAL FLEET SIZE", f"{total_fleet:,} Units", "Active RMM Inventory", False),
        (f"DUE NOW ({base_year})", f"{due_base:,} Units", "Immediate Refresh", True),
        ("WARRANTY EXPIRED", f"{w_expired:,} Units", exp_pct, True),
        ("EXPIRING SOON (<90D)", f"{w_soon:,} Units", "Review Warranties", True)
    ]
    draw_kpi_cards(ws1, kpis1)

    # Clinic Refresh Breakdown Table
    ws1.cell(7, 1).value = "Clinic-by-Clinic Refresh Schedule"
    ws1.cell(7, 1).font = font_section

    headers_t1 = [
        "Clinic / Location", f"{years[0]} (Due Now)", str(years[1]), str(years[2]), str(years[3]),
        "Total Fleet", "Laptops", "Desktops", "Dental AIO"
    ]
    ws1.row_dimensions[8].height = 25
    for c_idx, h in enumerate(headers_t1, 1):
        cell = ws1.cell(8, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_navy_header
        cell.alignment = align_header
        cell.border = thin_border

    sorted_clinics = sorted(clinic_map.keys())
    if "newWindowsWorkstation" in sorted_clinics:
        sorted_clinics.remove("newWindowsWorkstation")
        sorted_clinics.append("newWindowsWorkstation")

    r1 = 9
    for loc in sorted_clinics:
        ws1.row_dimensions[r1].height = 19
        cdata = clinic_map[loc]
        is_odd = (r1 % 2 == 1)
        bg = fill_zebra if is_odd else fill_white

        due_val = cdata.get(years[0], 0)
        row_vals = [
            (loc, align_left, font_tbl_bold if loc != "newWindowsWorkstation" else font_tbl_data, bg),
            (due_val, align_center, font_red if due_val > 0 else font_tbl_data, fill_red if due_val > 0 else bg),
            (cdata.get(years[1], 0), align_center, font_tbl_data, bg),
            (cdata.get(years[2], 0), align_center, font_tbl_data, bg),
            (cdata.get(years[3], 0), align_center, font_tbl_data, bg),
            (cdata["Total"], align_center, font_tbl_bold, bg),
            (cdata["Laptop"], align_center, font_tbl_data, bg),
            (cdata["Desktop"], align_center, font_tbl_data, bg),
            (cdata["Dental"], align_center, font_tbl_data, bg)
        ]
        for c_idx, (val, aln, fnt, fill) in enumerate(row_vals, 1):
            cell = ws1.cell(r1, c_idx, val)
            cell.alignment = aln
            cell.font = fnt
            cell.fill = fill
            cell.border = thin_border
        r1 += 1

    # Total row
    ws1.row_dimensions[r1].height = 22
    total_cells = [
        "TOTAL",
        f"=SUM(B9:B{r1-1})", f"=SUM(C9:C{r1-1})", f"=SUM(D9:D{r1-1})", f"=SUM(E9:E{r1-1})",
        f"=SUM(F9:F{r1-1})", f"=SUM(G9:G{r1-1})", f"=SUM(H9:H{r1-1})", f"=SUM(I9:I{r1-1})"
    ]
    for c_idx, v in enumerate(total_cells, 1):
        cell = ws1.cell(r1, c_idx, v)
        cell.font = font_tbl_total
        cell.fill = fill_total
        cell.border = total_border
        cell.alignment = align_left if c_idx == 1 else align_center
        if c_idx >= 2:
            cell.number_format = '#,##0'

    # Mini-Table 2: Warranty Status Summary (Placed at Column L)
    ws1.cell(7, 12).value = "Warranty Status Summary"
    ws1.cell(7, 12).font = font_section

    headers_w = ["Warranty Status", "Count", "% of Fleet"]
    ws1.row_dimensions[8].height = 25
    for c_idx, h in enumerate(headers_w, 12):
        cell = ws1.cell(8, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_sub_header
        cell.alignment = align_header
        cell.border = thin_border

    w_summary_data = [
        ("Active", w_counts.get("Active", 0), font_green, fill_green),
        ("Expiring Soon", w_counts.get("Expiring Soon", 0), font_yellow, fill_yellow),
        ("Expired", w_counts.get("Expired", 0), font_red, fill_red)
    ]

    r_w = 9
    for status, cnt, fnt, fill in w_summary_data:
        ws1.cell(r_w, 12, status).font = font_tbl_bold
        ws1.cell(r_w, 12).fill = fill
        ws1.cell(r_w, 12).border = thin_border
        ws1.cell(r_w, 12).alignment = align_left

        c_cnt = ws1.cell(r_w, 13, cnt)
        c_cnt.font = fnt
        c_cnt.fill = fill
        c_cnt.border = thin_border
        c_cnt.alignment = align_center
        c_cnt.number_format = '#,##0'

        c_pct = ws1.cell(r_w, 14, cnt / total_fleet if total_fleet else 0)
        c_pct.font = fnt
        c_pct.fill = fill
        c_pct.border = thin_border
        c_pct.alignment = align_center
        c_pct.number_format = '0.0%'
        r_w += 1

    # Warranty Total
    ws1.cell(r_w, 12, "TOTAL").font = font_tbl_total
    ws1.cell(r_w, 12).fill = fill_total
    ws1.cell(r_w, 12).border = total_border
    ws1.cell(r_w, 13, f"=SUM(M9:M{r_w-1})").font = font_tbl_total
    ws1.cell(r_w, 13).fill = fill_total
    ws1.cell(r_w, 13).border = total_border
    ws1.cell(r_w, 13).alignment = align_center
    ws1.cell(r_w, 13).number_format = '#,##0'
    ws1.cell(r_w, 14, f"=SUM(N9:N{r_w-1})").font = font_tbl_total
    ws1.cell(r_w, 14).fill = fill_total
    ws1.cell(r_w, 14).border = total_border
    ws1.cell(r_w, 14).alignment = align_center
    ws1.cell(r_w, 14).number_format = '0.0%'

    # Add Native Excel Doughnut Chart for Warranty
    chart_w = DoughnutChart()
    chart_w.title = "Fleet Warranty Status"
    chart_w.style = 10
    chart_w.width = 14
    chart_w.height = 7
    data_ref = Reference(ws1, min_col=13, min_row=8, max_row=11)
    cats_ref = Reference(ws1, min_col=12, min_row=9, max_row=11)
    chart_w.add_data(data_ref, titles_from_data=True)
    chart_w.set_categories(cats_ref)
    ws1.add_chart(chart_w, "L14")

    # Mini-Table 3: Refresh Year Summary for Chart
    ws1.cell(r1 + 2, 1).value = "Annual Refresh Distribution"
    ws1.cell(r1 + 2, 1).font = font_section

    headers_rf = ["Refresh Year", "Units Scheduled", "Pct of Fleet"]
    for c_idx, h in enumerate(headers_rf, 1):
        cell = ws1.cell(r1 + 3, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_navy_header
        cell.alignment = align_header
        cell.border = thin_border

    rf_table_row = r1 + 4
    for yr in years:
        cnt = rf_counts.get(yr, 0)
        ws1.cell(rf_table_row, 1, str(yr)).alignment = align_center
        ws1.cell(rf_table_row, 1).font = font_tbl_bold
        ws1.cell(rf_table_row, 1).border = thin_border

        c_units = ws1.cell(rf_table_row, 2, cnt)
        c_units.alignment = align_center
        c_units.font = font_red if yr == years[0] else font_tbl_data
        c_units.fill = fill_red if yr == years[0] else fill_white
        c_units.border = thin_border
        c_units.number_format = '#,##0'

        c_pct = ws1.cell(rf_table_row, 3, cnt / total_fleet if total_fleet else 0)
        c_pct.alignment = align_center
        c_pct.font = font_tbl_data
        c_pct.border = thin_border
        c_pct.number_format = '0.0%'
        rf_table_row += 1

    # Add Bar Chart for Refresh Forecast
    chart_rf = BarChart()
    chart_rf.type = "col"
    chart_rf.style = 10
    chart_rf.title = "Hardware Refresh Wave Forecast (Units)"
    chart_rf.y_axis.title = "Device Count"
    chart_rf.x_axis.title = "Replacement Year"
    chart_rf.width = 16
    chart_rf.height = 7
    chart_rf.legend = None
    chart_rf.dLbls = DataLabelList()
    chart_rf.dLbls.showVal = True
    chart_rf.x_axis.delete = False
    chart_rf.y_axis.delete = False
    rf_data = Reference(ws1, min_col=2, min_row=r1 + 3, max_row=rf_table_row - 1)
    rf_cats = Reference(ws1, min_col=1, min_row=r1 + 4, max_row=rf_table_row - 1)
    chart_rf.add_data(rf_data, titles_from_data=True)
    chart_rf.set_categories(rf_cats)
    ws1.add_chart(chart_rf, f"E{r1 + 2}")

    # Auto column widths
    ws1_widths = {1: 26, 2: 15, 3: 11, 4: 11, 5: 11, 6: 13, 7: 11, 8: 11, 9: 12, 10: 4, 11: 4, 12: 18, 13: 12, 14: 13}
    for c_idx, w in ws1_widths.items():
        ws1.column_dimensions[get_column_letter(c_idx)].width = w

def _build_patch_reboot_tab(wb, inventory_data, total_fleet):
    print("Building Tab 2: Patch & Reboot Health...")
    ws2 = create_tab(wb, "Patch & Reboot Health")
    ws2.cell(1, 1).value = "ENDPOINT SECURITY, PATCHING & REBOOT HYGIENE"
    ws2.cell(1, 1).font = font_title
    ws2.cell(2, 1).value = f"Tracking Windows 11 Build Compliance, Reboot Frequency & Patch Latency | {config.CURRENT_MONTH_YEAR}"
    ws2.cell(2, 1).font = font_subtitle

    # Calculations
    boot_over_7 = sum(1 for d in inventory_data if d["Boot"] > 7)
    boot_over_30 = sum(1 for d in inventory_data if d["Boot"] > 30)
    update_over_30 = sum(1 for d in inventory_data if d["Update"] > 30)
    os_counts = Counter(d["OS"] for d in inventory_data)
    legacy_os_23h2 = os_counts.get("23H2", 0)
    reboot_compliance_pct = (1 - (boot_over_7 / total_fleet)) * 100 if total_fleet else 100.0

    kpis2 = [
        ("OVERDUE REBOOTS (>7D)", f"{boot_over_7:,} Units", "Missed Weekly Reboot", True),
        ("CRITICAL UPTIME (>30D)", f"{boot_over_30:,} Units", "Skipping Patch Restarts", True),
        ("OVERDUE UPDATES (>30D)", f"{update_over_30:,} Units", "Unpatched Vulnerability", True),
        ("LEGACY OS (23H2)", f"{legacy_os_23h2:,} Units", "End of Support Upgrade", True),
        ("WEEKLY REBOOT COMPL.", f"{reboot_compliance_pct:.1f}%", "Fleet Meeting SLA", False)
    ]
    draw_kpi_cards(ws2, kpis2)

    # Table 1: OS Version Breakdown
    ws2.cell(7, 1).value = "Windows Operating System Build Adoption"
    ws2.cell(7, 1).font = font_section

    headers_os = ["Windows Build", "Device Count", "% of Fleet", "Support / Security Status"]
    ws2.row_dimensions[8].height = 25
    for c_idx, h in enumerate(headers_os, 1):
        cell = ws2.cell(8, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_navy_header
        cell.alignment = align_header
        cell.border = thin_border

    os_order = [
        ("25H2", "Mainstream Standard / Active Support", font_green, fill_green),
        ("24H2", "Previous Release / Active Support", font_tbl_data, fill_white),
        ("26H2", "Next-Gen", font_purple, fill_purple),
        ("23H2", "Deprecated / Immediate Upgrade Target", font_red, fill_red)
    ]

    r_os = 9
    for build, status, fnt, fill in os_order:
        cnt = os_counts.get(build, 0)
        ws2.row_dimensions[r_os].height = 19
        ws2.cell(r_os, 1, build).alignment = align_center
        ws2.cell(r_os, 1).font = font_tbl_bold
        ws2.cell(r_os, 1).border = thin_border
        ws2.cell(r_os, 1).fill = fill_zebra if r_os % 2 == 1 else fill_white
        
        c_cnt = ws2.cell(r_os, 2, cnt)
        c_cnt.alignment = align_center
        c_cnt.font = fnt
        c_cnt.fill = fill
        c_cnt.border = thin_border
        c_cnt.number_format = '#,##0'
        
        c_pct = ws2.cell(r_os, 3, cnt / total_fleet if total_fleet else 0)
        c_pct.alignment = align_center
        c_pct.font = fnt
        c_pct.fill = fill
        c_pct.border = thin_border
        c_pct.number_format = '0.0%'
        
        c_st = ws2.cell(r_os, 4, status)
        c_st.alignment = align_left
        c_st.font = fnt
        c_st.fill = fill
        c_st.border = thin_border
        r_os += 1

    # Total OS
    ws2.cell(r_os, 1, "TOTAL").font = font_tbl_total
    ws2.cell(r_os, 1).fill = fill_total
    ws2.cell(r_os, 1).border = total_border
    ws2.cell(r_os, 1).alignment = align_center

    ws2.cell(r_os, 2, f"=SUM(B9:B{r_os-1})").font = font_tbl_total
    ws2.cell(r_os, 2).fill = fill_total
    ws2.cell(r_os, 2).border = total_border
    ws2.cell(r_os, 2).alignment = align_center
    ws2.cell(r_os, 2).number_format = '#,##0'

    ws2.cell(r_os, 3, f"=SUM(C9:C{r_os-1})").font = font_tbl_total
    ws2.cell(r_os, 3).fill = fill_total
    ws2.cell(r_os, 3).border = total_border
    ws2.cell(r_os, 3).alignment = align_center
    ws2.cell(r_os, 3).number_format = '0.0%'

    ws2.cell(r_os, 4, "Enterprise Baseline").font = font_tbl_total
    ws2.cell(r_os, 4).fill = fill_total
    ws2.cell(r_os, 4).border = total_border

    # Add OS Doughnut Chart
    chart_os = DoughnutChart()
    chart_os.title = "Operating System Distribution"
    chart_os.style = 10
    chart_os.width = 14
    chart_os.height = 7
    os_data = Reference(ws2, min_col=2, min_row=8, max_row=r_os - 1)
    os_cats = Reference(ws2, min_col=1, min_row=9, max_row=r_os - 1)
    chart_os.add_data(os_data, titles_from_data=True)
    chart_os.set_categories(os_cats)
    ws2.add_chart(chart_os, "G7")

    # Table 2: Reboot Hygiene / Uptime Tiers
    r_uptime_start = r_os + 2
    ws2.cell(r_uptime_start, 1).value = "Endpoint Reboot Hygiene (Last Boot Time)"
    ws2.cell(r_uptime_start, 1).font = font_section

    headers_up = ["Uptime Tier", "Threshold", "Device Count", "% of Fleet", "Compliance Assessment"]
    for c_idx, h in enumerate(headers_up, 1):
        cell = ws2.cell(r_uptime_start + 1, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_navy_header
        cell.alignment = align_header
        cell.border = thin_border

    boot_tiers = [
        ("Tier 1: Compliant", "<= 7 Days", sum(1 for d in inventory_data if d["Boot"] <= 7), font_green, fill_green, "Normal - Compliant with weekly reboot"),
        ("Tier 2: Minor Warning", "8 - 14 Days", sum(1 for d in inventory_data if 8 <= d["Boot"] <= 14), font_yellow, fill_yellow, "Missed 1 automated maintenance reboot"),
        ("Tier 3: Moderate Risk", "15 - 30 Days", sum(1 for d in inventory_data if 15 <= d["Boot"] <= 30), font_yellow, fill_yellow, "Missed multiple reboot windows"),
        ("Tier 4: High Neglect", "31 - 90 Days", sum(1 for d in inventory_data if 31 <= d["Boot"] <= 90), font_red, fill_red, "Severely delayed - Force reboot recommended"),
        ("Tier 5: Critical / Stale", "> 90 Days", sum(1 for d in inventory_data if d["Boot"] > 90), font_purple, fill_purple, "Critical - Huntress drop risk / Physical check")
    ]

    r_tier = r_uptime_start + 2
    for name, thresh, cnt, fnt, fill, note in boot_tiers:
        ws2.row_dimensions[r_tier].height = 19
        ws2.cell(r_tier, 1, name).font = font_tbl_bold
        ws2.cell(r_tier, 1).border = thin_border
        ws2.cell(r_tier, 1).fill = fill_zebra if r_tier % 2 == 1 else fill_white
        
        ws2.cell(r_tier, 2, thresh).font = font_tbl_data
        ws2.cell(r_tier, 2).alignment = align_center
        ws2.cell(r_tier, 2).border = thin_border
        ws2.cell(r_tier, 2).fill = fill_zebra if r_tier % 2 == 1 else fill_white
        
        c_cnt = ws2.cell(r_tier, 3, cnt)
        c_cnt.font = fnt
        c_cnt.fill = fill
        c_cnt.alignment = align_center
        c_cnt.border = thin_border
        c_cnt.number_format = '#,##0'
        
        c_pct = ws2.cell(r_tier, 4, cnt / total_fleet if total_fleet else 0)
        c_pct.font = fnt
        c_pct.fill = fill
        c_pct.alignment = align_center
        c_pct.border = thin_border
        c_pct.number_format = '0.0%'
        
        c_note = ws2.cell(r_tier, 5, note)
        c_note.font = fnt
        c_note.fill = fill
        c_note.border = thin_border
        r_tier += 1

    # Total Reboot Table
    ws2.cell(r_tier, 1, "TOTAL").font = font_tbl_total
    ws2.cell(r_tier, 1).fill = fill_total
    ws2.cell(r_tier, 1).border = total_border
    ws2.cell(r_tier, 2, "-").alignment = align_center
    ws2.cell(r_tier, 2).fill = fill_total
    ws2.cell(r_tier, 2).border = total_border
    ws2.cell(r_tier, 3, f"=SUM(C{r_uptime_start+2}:C{r_tier-1})").font = font_tbl_total
    ws2.cell(r_tier, 3).fill = fill_total
    ws2.cell(r_tier, 3).border = total_border
    ws2.cell(r_tier, 3).alignment = align_center
    ws2.cell(r_tier, 3).number_format = '#,##0'
    ws2.cell(r_tier, 4, f"=SUM(D{r_uptime_start+2}:D{r_tier-1})").font = font_tbl_total
    ws2.cell(r_tier, 4).fill = fill_total
    ws2.cell(r_tier, 4).border = total_border
    ws2.cell(r_tier, 4).alignment = align_center
    ws2.cell(r_tier, 4).number_format = '0.0%'
    ws2.cell(r_tier, 5, "Total Fleet Status").font = font_tbl_total
    ws2.cell(r_tier, 5).fill = fill_total
    ws2.cell(r_tier, 5).border = total_border

    # Add Reboot Bar Chart
    chart_boot = BarChart()
    chart_boot.type = "col"
    chart_boot.style = 10
    chart_boot.title = "Fleet Uptime & Reboot Compliance"
    chart_boot.y_axis.title = "Number of Devices"
    chart_boot.x_axis.title = "Uptime Bracket"
    chart_boot.width = 15
    chart_boot.height = 7
    chart_boot.legend = None
    chart_boot.dLbls = DataLabelList()
    chart_boot.dLbls.showVal = True
    chart_boot.x_axis.delete = False
    chart_boot.y_axis.delete = False
    b_data = Reference(ws2, min_col=3, min_row=r_uptime_start + 1, max_row=r_tier - 1)
    b_cats = Reference(ws2, min_col=2, min_row=r_uptime_start + 2, max_row=r_tier - 1)
    chart_boot.add_data(b_data, titles_from_data=True)
    chart_boot.set_categories(b_cats)
    ws2.add_chart(chart_boot, "G19")

    # Auto column widths
    ws2_widths = {1: 24, 2: 18, 3: 15, 4: 33, 5: 39, 6: 4, 7: 18, 8: 15, 9: 15}
    for c_idx, w in ws2_widths.items():
        ws2.column_dimensions[get_column_letter(c_idx)].width = w

def _build_performance_capacity_tab(wb, inventory_data, total_fleet):
    print("Building Tab 3: Performance & Capacity...")
    ws3 = create_tab(wb, "Performance & Capacity")
    ws3.cell(1, 1).value = "STORAGE CAPACITY & MEMORY BOTTLENECK AUDIT"
    ws3.cell(1, 1).font = font_title
    ws3.cell(2, 1).value = f"Surfacing Disk Space Shortages (>=75%) and Sub-16GB RAM Workstations in Clinical Workflows | {config.CURRENT_MONTH_YEAR}"
    ws3.cell(2, 1).font = font_subtitle

    # Counts
    disk_crit = sum(1 for d in inventory_data if d["Usage"] is not None and float(d["Usage"]) >= 0.75)
    disk_warn = sum(1 for d in inventory_data if d["Usage"] is not None and 0.65 <= float(d["Usage"]) < 0.75)
    ram_8gb = sum(1 for d in inventory_data if d["RAM"] is not None and float(d["RAM"]) < 10)
    ram_16gb = sum(1 for d in inventory_data if d["RAM"] is not None and 10 <= float(d["RAM"]) < 20)
    ram_32plus = sum(1 for d in inventory_data if d["RAM"] is not None and float(d["RAM"]) >= 20)

    kpis3 = [
        ("CRITICAL STORAGE (>=75%)", f"{disk_crit:,} Units", "Immediate Disk Cleanup", True),
        ("STORAGE WARNING (65-74%)", f"{disk_warn:,} Units", "Approaching Threshold", True),
        ("8 GB RAM DEVICES", f"{ram_8gb:,} Units", f"{(ram_8gb/total_fleet)*100:.1f}% of Fleet" if total_fleet else "0%", True),
        ("16 GB STANDARD RAM", f"{ram_16gb:,} Units", "Optimal Fleet Standard", False),
        ("HIGH PERF (32 GB+)", f"{ram_32plus:,} Units", "Imaging / Heavy Admin", False)
    ]
    draw_kpi_cards(ws3, kpis3)

    # Table 1: Storage Tiers
    ws3.cell(7, 1).value = "Fleet Storage Capacity Distribution"
    ws3.cell(7, 1).font = font_section

    headers_disk = ["Drive Usage Tier", "Utilization Range", "Endpoints", "% of Fleet", "Recommended Action"]
    ws3.row_dimensions[8].height = 25
    for c_idx, h in enumerate(headers_disk, 1):
        cell = ws3.cell(8, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_navy_header
        cell.alignment = align_header
        cell.border = thin_border

    disk_tiers = [
        ("Healthy / Low", "< 60%", sum(1 for d in inventory_data if d["Usage"] is not None and float(d["Usage"]) < 0.60), font_green, fill_green, "Normal - No action required"),
        ("Moderate", "60% - 64%", sum(1 for d in inventory_data if d["Usage"] is not None and 0.60 <= float(d["Usage"]) < 0.65), font_tbl_data, fill_white, "Healthy operating headroom"),
        ("Elevated Warning", "65% - 74%", sum(1 for d in inventory_data if d["Usage"] is not None and 0.65 <= float(d["Usage"]) < 0.75), font_yellow, fill_yellow, "Monitor and run disk cleanup scripts"),
        ("High Alert", "75% - 84%", sum(1 for d in inventory_data if d["Usage"] is not None and 0.75 <= float(d["Usage"]) < 0.85), font_red, fill_red, "Purge temp files / assess profile bloat"),
        ("Critical (>=85%)", ">= 85%", sum(1 for d in inventory_data if d["Usage"] is not None and float(d["Usage"]) >= 0.85), font_purple, fill_purple, "Immediate intervention / replace SSD")
    ]

    r_d = 9
    for tier, rng, cnt, fnt, fill, action in disk_tiers:
        ws3.row_dimensions[r_d].height = 19
        ws3.cell(r_d, 1, tier).font = font_tbl_bold
        ws3.cell(r_d, 1).border = thin_border
        ws3.cell(r_d, 1).fill = fill_zebra if r_d % 2 == 1 else fill_white
        
        ws3.cell(r_d, 2, rng).font = font_tbl_data
        ws3.cell(r_d, 2).alignment = align_center
        ws3.cell(r_d, 2).border = thin_border
        ws3.cell(r_d, 2).fill = fill_zebra if r_d % 2 == 1 else fill_white
        
        c_cnt = ws3.cell(r_d, 3, cnt)
        c_cnt.font = fnt
        c_cnt.fill = fill
        c_cnt.alignment = align_center
        c_cnt.border = thin_border
        c_cnt.number_format = '#,##0'
        
        c_pct = ws3.cell(r_d, 4, cnt / total_fleet if total_fleet else 0)
        c_pct.font = fnt
        c_pct.fill = fill
        c_pct.alignment = align_center
        c_pct.border = thin_border
        c_pct.number_format = '0.0%'
        
        c_act = ws3.cell(r_d, 5, action)
        c_act.font = fnt
        c_act.fill = fill
        c_act.border = thin_border
        r_d += 1

    # Table 2: RAM Configuration Profile (Placed at Column G)
    ws3.cell(7, 7).value = "RAM Hardware Profile"
    ws3.cell(7, 7).font = font_section

    headers_ram = ["RAM Size", "Workstation Role", "Device Count", "% of Fleet"]
    for c_idx, h in enumerate(headers_ram, 7):
        cell = ws3.cell(8, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_sub_header
        cell.alignment = align_header
        cell.border = thin_border

    ram_tiers = [
        ("8 GB (Legacy)", "Clinical Exam & Office", ram_8gb, font_yellow, fill_yellow),
        ("16 GB (Standard)", "Clinical Exam & Office", ram_16gb, font_green, fill_green),
        ("24 - 32 GB", "Clinical Exam & Office", sum(1 for d in inventory_data if d["RAM"] is not None and 20 <= float(d["RAM"]) < 35), font_tbl_data, fill_white),
        ("40 - 64 GB", "Engineering / Power User", sum(1 for d in inventory_data if d["RAM"] is not None and float(d["RAM"]) >= 35), font_purple, fill_purple)
    ]

    r_ram = 9
    for size, role, cnt, fnt, fill in ram_tiers:
        ws3.cell(r_ram, 7, size).font = font_tbl_bold
        ws3.cell(r_ram, 7).border = thin_border
        ws3.cell(r_ram, 7).fill = fill
        
        ws3.cell(r_ram, 8, role).font = fnt
        ws3.cell(r_ram, 8).border = thin_border
        ws3.cell(r_ram, 8).fill = fill
        
        c_cnt = ws3.cell(r_ram, 9, cnt)
        c_cnt.font = fnt
        c_cnt.fill = fill
        c_cnt.alignment = align_center
        c_cnt.border = thin_border
        c_cnt.number_format = '#,##0'
        
        c_pct = ws3.cell(r_ram, 10, cnt / total_fleet if total_fleet else 0)
        c_pct.font = fnt
        c_pct.fill = fill
        c_pct.alignment = align_center
        c_pct.border = thin_border
        c_pct.number_format = '0.0%'
        r_ram += 1

    # Add RAM Doughnut Chart
    chart_ram = DoughnutChart()
    chart_ram.title = "RAM Hardware Profile"
    chart_ram.style = 10
    chart_ram.width = 10
    chart_ram.height = 5
    ram_data = Reference(ws3, min_col=9, min_row=8, max_row=r_ram - 1)
    ram_cats = Reference(ws3, min_col=7, min_row=9, max_row=r_ram - 1)
    chart_ram.add_data(ram_data, titles_from_data=True)
    chart_ram.set_categories(ram_cats)
    ws3.add_chart(chart_ram, "G14")

    # Table 3: Top Critical Storage Devices Watchlist
    r_top = max(r_d, 21) + 1
    ws3.cell(r_top, 1).value = "Top 12 Fullest Primary Drives (Immediate Intervention Watchlist)"
    ws3.cell(r_top, 1).font = font_section

    headers_top_d = ["Device", "Location", "Device Role", "Device Model", "Last User", "Capacity (GB)", "Free (GB)", "Usage %", "RAM", "Warranty"]
    for c_idx, h in enumerate(headers_top_d, 1):
        cell = ws3.cell(r_top + 1, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_navy_header
        cell.alignment = align_header
        cell.border = thin_border

    valid_usage = [d for d in inventory_data if d["Usage"] is not None]
    valid_usage.sort(key=lambda x: float(x["Usage"]), reverse=True)

    curr_r = r_top + 2
    for d in valid_usage[:12]:
        ws3.row_dimensions[curr_r].height = 19
        pct = float(d["Usage"])
        is_crit = pct >= 0.75
        fnt_alert = font_red if is_crit else font_yellow
        fill_alert = fill_red if is_crit else fill_yellow
        
        row_vals = [
            (d["Device"], align_left, font_tbl_bold, fill_alert),
            (d["Location"], align_left, font_tbl_data, fill_white),
            (d["Role"], align_left, font_tbl_data, fill_white),
            (d["Model"], align_left, font_tbl_data, fill_white),
            (d["User"], align_left, font_tbl_data, fill_white),
            (d["Capacity"], align_right, font_tbl_data, fill_white),
            (d["Free"], align_right, font_tbl_bold, fill_alert),
            (pct, align_center, fnt_alert, fill_alert),
            (f"{float(d['RAM']):.1f} GB" if d["RAM"] else "N/A", align_center, font_tbl_data, fill_white),
            (d["WarrantyStatus"], align_center, font_red if d["WarrantyStatus"] == "Expired" else font_green, fill_red if d["WarrantyStatus"] == "Expired" else fill_green)
        ]
        for c_idx, (val, aln, fnt, fill) in enumerate(row_vals, 1):
            cell = ws3.cell(curr_r, c_idx, val)
            cell.alignment = aln
            cell.font = fnt
            cell.fill = fill
            cell.border = thin_border
            if c_idx in [6, 7]:
                cell.number_format = '#,##0.0'
            elif c_idx == 8:
                cell.number_format = '0.0%'
        curr_r += 1

    # Auto column widths
    ws3_widths = {1: 18, 2: 24, 3: 16, 4: 24, 5: 31, 6: 14, 7: 21, 8: 21, 9: 12, 10: 14}
    for c_idx, w in ws3_widths.items():
        ws3.column_dimensions[get_column_letter(c_idx)].width = w

def _build_site_allocation_tab(wb, inventory_data, total_fleet, clinic_map):
    print("Building Tab 4: Site Allocation...")
    ws4 = create_tab(wb, "Site Allocation")
    ws4.cell(1, 1).value = "CLINICAL SITE ALLOCATION & MOBILITY AUDIT"
    ws4.cell(1, 1).font = font_title
    ws4.cell(2, 1).value = f"Hardware Footprint, Form Factor Breakdown & Mobility Ratios Across All Physical Clinics | {config.CURRENT_MONTH_YEAR}"
    ws4.cell(2, 1).font = font_subtitle

    desktops_total = sum(1 for d in inventory_data if "Laptop" not in d["Role"] and "Dental" not in d["Role"])
    laptops_total = sum(1 for d in inventory_data if "Laptop" in d["Role"])
    dental_total = sum(1 for d in inventory_data if "Dental" in d["Role"])
    mobility_ratio = (laptops_total / total_fleet) * 100 if total_fleet else 0.0

    kpis4 = [
        ("TOTAL LOCATIONS", f"{len(clinic_map):,} Sites", "Including Admin & Staging", False),
        ("FIXED DESKTOPS", f"{desktops_total:,} Units", f"{(desktops_total/total_fleet)*100:.1f}% of Fleet" if total_fleet else "0%", False),
        ("MOBILE LAPTOPS", f"{laptops_total:,} Units", f"{mobility_ratio:.1f}% Mobility Ratio", False),
        ("DENTAL ALL-IN-ONES", f"{dental_total:,} Units", "Specialized Operator PC", False)
    ]
    draw_kpi_cards(ws4, kpis4)

    # Table 1: Detailed Site Allocation Table
    ws4.cell(7, 1).value = "Clinic-by-Clinic Equipment Footprint & Mobility Ratio"
    ws4.cell(7, 1).font = font_section

    headers_site = [
        "Clinic / Facility", "Desktops", "Laptops", "Dental AIO",
        "Total Endpoints", "% Fleet Share", "Mobility %", "Avg RAM (GB)", "Avg Free Disk (GB)"
    ]
    ws4.row_dimensions[8].height = 25
    for c_idx, h in enumerate(headers_site, 1):
        cell = ws4.cell(8, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_navy_header
        cell.alignment = align_header
        cell.border = thin_border

    clinic_sorted_by_size = sorted(clinic_map.items(), key=lambda x: x[1]["Total"], reverse=True)

    site_stats = {}
    for loc in clinic_map.keys():
        loc_devs = [d for d in inventory_data if d["Location"] == loc]
        rams = [float(d["RAM"]) for d in loc_devs if d["RAM"] is not None]
        frees = [float(d["Free"]) for d in loc_devs if d["Free"] is not None]
        site_stats[loc] = {
            "avg_ram": sum(rams)/len(rams) if rams else 0,
            "avg_free": sum(frees)/len(frees) if frees else 0
        }

    r_site = 9
    for loc, cdata in clinic_sorted_by_size:
        ws4.row_dimensions[r_site].height = 19
        is_odd = (r_site % 2 == 1)
        bg = fill_zebra if is_odd else fill_white
        tot = cdata["Total"]
        mob_pct = (cdata["Laptop"] / tot) if tot > 0 else 0
        
        row_vals = [
            (loc, align_left, font_tbl_bold if loc != "newWindowsWorkstation" else font_tbl_data, bg),
            (cdata["Desktop"], align_center, font_tbl_data, bg),
            (cdata["Laptop"], align_center, font_tbl_bold if cdata["Laptop"] > 0 else font_tbl_data, fill_green if mob_pct > 0.40 else bg),
            (cdata["Dental"], align_center, font_tbl_data, bg),
            (tot, align_center, font_tbl_bold, bg),
            (tot / total_fleet if total_fleet else 0, align_center, font_tbl_data, bg),
            (mob_pct, align_center, font_tbl_bold, fill_green if mob_pct > 0.40 else bg),
            (site_stats[loc]["avg_ram"], align_center, font_tbl_data, bg),
            (site_stats[loc]["avg_free"], align_right, font_tbl_data, bg)
        ]
        
        for c_idx, (val, aln, fnt, fill) in enumerate(row_vals, 1):
            cell = ws4.cell(r_site, c_idx, val)
            cell.alignment = aln
            cell.font = fnt
            cell.fill = fill
            cell.border = thin_border
            if c_idx in [2, 3, 4, 5]:
                cell.number_format = '#,##0'
            elif c_idx in [6, 7]:
                cell.number_format = '0.0%'
            elif c_idx == 8:
                cell.number_format = '#,##0.0'
            elif c_idx == 9:
                cell.number_format = '#,##0.0'
        r_site += 1

    # Total Row
    ws4.row_dimensions[r_site].height = 22
    total_site_vals = [
        "TOTAL",
        f"=SUM(B9:B{r_site-1})", f"=SUM(C9:C{r_site-1})", f"=SUM(D9:D{r_site-1})",
        f"=SUM(E9:E{r_site-1})", f"=SUM(F9:F{r_site-1})", f"=IF(E{r_site}>0, C{r_site}/E{r_site}, 0)",
        f"=IFERROR(AVERAGE(H9:H{r_site-1}), 0)", f"=IFERROR(AVERAGE(I9:I{r_site-1}), 0)"
    ]
    for c_idx, v in enumerate(total_site_vals, 1):
        cell = ws4.cell(r_site, c_idx, v)
        cell.font = font_tbl_total
        cell.fill = fill_total
        cell.border = total_border
        cell.alignment = align_left if c_idx == 1 else align_center
        if c_idx in [2, 3, 4, 5]:
            cell.number_format = '#,##0'
        elif c_idx in [6, 7]:
            cell.number_format = '0.0%'
        elif c_idx in [8, 9]:
            cell.number_format = '#,##0.0'

    # Add Stacked Column Chart for Top 10 Sites
    chart_site = BarChart()
    chart_site.type = "col"
    chart_site.grouping = "stacked"
    chart_site.overlap = 100
    chart_site.title = "Device Form Factor Allocation by Clinic"
    chart_site.style = 10
    chart_site.width = 18
    chart_site.height = 9
    chart_site.y_axis.title = "Number of Devices"
    chart_site.x_axis.title = "Clinic Location"

    data_site = Reference(ws4, min_col=2, min_row=8, max_col=4, max_row=r_site - 1)
    cats_site = Reference(ws4, min_col=1, min_row=9, max_row=r_site - 1)
    chart_site.add_data(data_site, titles_from_data=True)
    chart_site.set_categories(cats_site)
    ws4.add_chart(chart_site, "K7")

    # Auto column widths
    ws4_widths = {1: 26, 2: 12, 3: 12, 4: 12, 5: 15, 6: 13, 7: 13, 8: 14, 9: 18, 10: 3}
    for c_idx, w in ws4_widths.items():
        ws4.column_dimensions[get_column_letter(c_idx)].width = w

def _build_stale_assets_tab(wb, inventory_data, total_fleet):
    print("Building Tab 5: Stale Assets...")
    ws5 = create_tab(wb, "Stale Assets")
    ws5.cell(1, 1).value = "INACTIVE ASSET AUDIT & HUNTRESS DROP RISK"
    ws5.cell(1, 1).font = font_title
    ws5.cell(2, 1).value = f"Tracking Endpoints Offline >90 Days (Huntress Agent Removal Window) & >1 Year Decommission Targets | {config.CURRENT_MONTH_YEAR}"
    ws5.cell(2, 1).font = font_subtitle

    stale_devices = []
    for d in inventory_data:
        u = d["Uptime"]
        b = d["Boot"]
        up = d["Update"]
        max_inact = max(u, b, up)
        if max_inact > 60:
            stale_devices.append((max_inact, d))

    stale_devices.sort(key=lambda x: x[0], reverse=True)

    over_365 = sum(1 for m, d in stale_devices if m > 365)
    over_90 = sum(1 for m, d in stale_devices if 90 < m <= 365)
    over_60 = sum(1 for m, d in stale_devices if 60 <= m <= 90)

    kpis5 = [
        ("DECOMMISSION (>365D)", f"{over_365:,} Units", ">1 Year Inactive (Retire AD)", True),
        ("HUNTRESS DROP RISK (91-365D)", f"{over_90:,} Units", "Huntress Deactivation Risk", True),
        ("MONITOR LIST (61-90D)", f"{over_60:,} Units", "Approaching 90-Day Limit", True),
        ("TOTAL HIGH INACTIVITY", f"{len(stale_devices):,} Units", "Action Required", True),
        ("AUDIT STATUS", "CRITICAL", "Immediate physical audit", True)
    ]
    draw_kpi_cards(ws5, kpis5)

    ws5.cell(7, 1).value = f"Delinquent & Inactive Endpoints Watchlist ({len(stale_devices)} Devices with >60 Days Inactivity)"
    ws5.cell(7, 1).font = font_section

    headers_stale = [
        "Device", "Location", "Device Role", "Device Model", "Last User",
        "Uptime (Days)", "Boot (Days)", "Update (Days)", "Warranty", "Recommended Action"
    ]
    ws5.row_dimensions[8].height = 25
    for c_idx, h in enumerate(headers_stale, 1):
        cell = ws5.cell(8, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_navy_header
        cell.alignment = align_header
        cell.border = thin_border

    r_stale = 9
    for max_days, d in stale_devices:
        ws5.row_dimensions[r_stale].height = 19
        is_odd = (r_stale % 2 == 1)
        bg = fill_zebra if is_odd else fill_white
        
        if max_days > 365:
            action = "Decommission & Retire AD Account"
            status_fill = fill_purple
            status_fnt = font_purple
        elif max_days > 180:
            action = "Locate in Storage / Re-deploy"
            status_fill = fill_red
            status_fnt = font_red
        elif max_days > 90:
            action = "Huntress Expired: Check Status / Reconnect"
            status_fill = fill_red
            status_fnt = font_red
        else:
            action = "Power On & Execute Windows Update"
            status_fill = fill_yellow
            status_fnt = font_yellow
            
        row_vals = [
            (d["Device"], align_left, font_tbl_bold, status_fill if max_days > 365 else bg),
            (d["Location"], align_left, font_tbl_data, bg),
            (d["Role"], align_left, font_tbl_data, bg),
            (d["Model"], align_left, font_tbl_data, bg),
            (d["User"], align_left, font_tbl_data, bg),
            (d["Uptime"], align_right, font_tbl_data, status_fill if d["Uptime"] > 365 else (fill_red if d["Uptime"] > 90 else bg)),
            (d["Boot"], align_right, font_tbl_data, status_fill if d["Boot"] > 365 else (fill_red if d["Boot"] > 90 else bg)),
            (d["Update"], align_right, font_tbl_data, status_fill if d["Update"] > 365 else (fill_red if d["Update"] > 90 else bg)),
            (d["WarrantyStatus"], align_center, font_red if d["WarrantyStatus"] == "Expired" else font_green, fill_red if d["WarrantyStatus"] == "Expired" else fill_green),
            (action, align_left, font_tbl_bold, status_fill)
        ]
        
        for c_idx, (val, aln, fnt, fill) in enumerate(row_vals, 1):
            cell = ws5.cell(r_stale, c_idx, val)
            cell.alignment = aln
            cell.font = fnt
            cell.fill = fill
            cell.border = thin_border
            if c_idx in [6, 7, 8]:
                cell.number_format = '#,##0'
        r_stale += 1

    # Mini Table 2: Inactivity Summary by Bracket (Column L)
    ws5.cell(7, 12).value = "Inactivity Severity Distribution"
    ws5.cell(7, 12).font = font_section

    headers_stale_sum = ["Inactivity Bracket", "Device Count", "Action Impact"]
    for c_idx, h in enumerate(headers_stale_sum, 12):
        cell = ws5.cell(8, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_sub_header
        cell.alignment = align_header
        cell.border = thin_border

    stale_summary_rows = [
        ("> 365 Days (Dormant)", over_365, "Decommission & License Reclamation", font_purple, fill_purple),
        ("91 - 365 Days (Critical)", over_90, "Huntress Agent Inactive / Drop", font_red, fill_red),
        ("61 - 90 Days (Warning)", over_60, "Approaching Inactivity Drop", font_yellow, fill_yellow)
    ]

    r_ss = 9
    for bracket, cnt, impact, fnt, fill in stale_summary_rows:
        ws5.cell(r_ss, 12, bracket).font = font_tbl_bold
        ws5.cell(r_ss, 12).fill = fill
        ws5.cell(r_ss, 12).border = thin_border
        
        c_cnt = ws5.cell(r_ss, 13, cnt)
        c_cnt.font = fnt
        c_cnt.fill = fill
        c_cnt.border = thin_border
        c_cnt.alignment = align_center
        c_cnt.number_format = '#,##0'
        
        c_imp = ws5.cell(r_ss, 14, impact)
        c_imp.font = fnt
        c_imp.fill = fill
        c_imp.border = thin_border
        r_ss += 1

    # Total Row
    ws5.cell(r_ss, 12, "TOTAL AT-RISK").font = font_tbl_total
    ws5.cell(r_ss, 12).fill = fill_total
    ws5.cell(r_ss, 12).border = total_border
    ws5.cell(r_ss, 13, f"=SUM(M9:M{r_ss-1})").font = font_tbl_total
    ws5.cell(r_ss, 13).fill = fill_total
    ws5.cell(r_ss, 13).border = total_border
    ws5.cell(r_ss, 13).alignment = align_center
    ws5.cell(r_ss, 13).number_format = '#,##0'
    ws5.cell(r_ss, 14, "Audit Priority 1").font = font_tbl_total
    ws5.cell(r_ss, 14).fill = fill_total
    ws5.cell(r_ss, 14).border = total_border

    # Add Inactivity Doughnut Chart
    chart_stale = DoughnutChart()
    chart_stale.title = "Stale Assets by Severity"
    chart_stale.style = 10
    chart_stale.width = 14
    chart_stale.height = 7
    stale_data = Reference(ws5, min_col=13, min_row=8, max_row=r_ss - 1)
    stale_cats = Reference(ws5, min_col=12, min_row=9, max_row=r_ss - 1)
    chart_stale.add_data(stale_data, titles_from_data=True)
    chart_stale.set_categories(stale_cats)
    ws5.add_chart(chart_stale, "L14")

    # Auto column widths
    ws5_widths = {1: 18, 2: 24, 3: 16, 4: 24, 5: 24, 6: 15, 7: 15, 8: 15, 9: 14, 10: 36, 11: 3, 12: 25, 13: 14, 14: 34}
    for c_idx, w in ws5_widths.items():
        ws5.column_dimensions[get_column_letter(c_idx)].width = w

def _build_ad_cross_ref_tab(wb, inventory_data, total_fleet, sheet_title="AD Cross ref"):
    print(f"Building {sheet_title}: Active Directory Cross-Reference Dashboard...")
    ws_ad = wb["AD Inventory"] if "AD Inventory" in wb.sheetnames else (wb["AD Computers"] if "AD Computers" in wb.sheetnames else None)
    if not ws_ad:
        return

    ad_headers = {ws_ad.cell(1, c).value: c for c in range(1, ws_ad.max_column + 1)}
    c_ad_name = ad_headers.get("Name", 1)
    c_ad_status = ad_headers.get("Status", 2)
    c_ad_clinic = ad_headers.get("Clinic", 3)
    c_ad_dept = ad_headers.get("Department", 4)
    c_ad_type = ad_headers.get("Device Type", 5)
    c_ad_os = ad_headers.get("OperatingSystem") or ad_headers.get("Operating System", 6)
    c_ad_lastlogon = ad_headers.get("LastLogonDate") or ad_headers.get("Last Logon Date", 8)
    c_ad_logondays = ad_headers.get("Last Logon - Days") or ad_headers.get("Last Logon (Days)", 9)
    c_ad_created = ad_headers.get("CreatedDate") or ad_headers.get("Created Date", 10)
    c_ad_createddays = ad_headers.get("Created - Days") or ad_headers.get("Created (Days)", 11)
    c_ad_ou1 = ad_headers.get("OU_Token1") or ad_headers.get("OU Token 1", 14)

    ad_records = []
    for r in range(2, ws_ad.max_row + 1):
        rec_name = ws_ad.cell(r, c_ad_name).value
        if not rec_name:
            continue
        ad_records.append({
            "Name": str(rec_name).strip(),
            "Status": str(ws_ad.cell(r, c_ad_status).value or "").strip(),
            "Clinic": str(ws_ad.cell(r, c_ad_clinic).value or "").strip(),
            "Department": str(ws_ad.cell(r, c_ad_dept).value or "").strip(),
            "DeviceType": str(ws_ad.cell(r, c_ad_type).value or "").strip(),
            "OS": str(ws_ad.cell(r, c_ad_os).value or "").strip(),
            "LastLogonDate": str(ws_ad.cell(r, c_ad_lastlogon).value or "").strip(),
            "LastLogonDays": ws_ad.cell(r, c_ad_logondays).value,
            "CreatedDate": str(ws_ad.cell(r, c_ad_created).value or "").strip(),
            "CreatedDays": ws_ad.cell(r, c_ad_createddays).value if c_ad_createddays else None,
            "OU1": str(ws_ad.cell(r, c_ad_ou1).value or "").strip()
        })

    ad_by_name = {a["Name"].upper(): a for a in ad_records}
    ad_by_norm = {normalize_hostname(a["Name"]): a for a in ad_records if a.get("Name")}
    ninja_names = {str(d["Device"]).strip().upper() for d in inventory_data if d.get("Device")}
    ninja_norms = {normalize_hostname(d["Device"]) for d in inventory_data if d.get("Device")}

    audit_devices = []
    for d in inventory_data:
        dev_name = str(d["Device"]).strip()
        loc = config.canonicalize_clinic_name(d.get("Location") or "Unassigned")
        role = d.get("Role") or "Workstation"
        user = d.get("User") or "N/A"
        up_days = d.get("Uptime") or 0

        match_ad = ad_by_name.get(dev_name.upper()) or ad_by_norm.get(normalize_hostname(dev_name))

        if match_ad:
            ad_status = match_ad["Status"]
            ad_clinic = config.canonicalize_clinic_name(match_ad["Clinic"] or loc)
            ad_logon_date = match_ad["LastLogonDate"]
            ad_days = match_ad["LastLogonDays"]
            try:
                ad_days_num = float(ad_days) if ad_days not in [None, ""] else None
            except (ValueError, TypeError):
                ad_days_num = None

            if ad_status.lower() in ["disabled", "false"]:
                sync_status = "CRITICAL: AD Disabled"
                category = "AD_DISABLED"
                priority = 1
                diag_reason = "Device online in Ninja RMM, but computer account is DISABLED in Active Directory"
                rec_action = "Security check: Re-enable AD account or isolate device"
                fill_st = fill_red
                font_st = font_red
            elif "Unassigned" in ad_clinic or "CN=Computers" in match_ad["OU1"]:
                sync_status = "OU Misplaced (CN=Computers)"
                category = "OU_MISPLACED"
                priority = 3
                diag_reason = "Computer object resides in default unassigned container (CN=Computers)"
                rec_action = f"Move object to {loc} clinic OU in Active Directory"
                fill_st = fill_yellow
                font_st = font_yellow
            elif ad_days_num is not None and ad_days_num > 90:
                sync_status = "AD Logon Lagging (>90d)"
                category = "AD_LAGGING"
                priority = 2
                diag_reason = f"Active in Ninja, but AD secure channel silent for {int(ad_days_num)} days"
                rec_action = "Test-ComputerSecureChannel / Reconnect domain trust"
                fill_st = fill_yellow
                font_st = font_yellow
            else:
                sync_status = "Synchronized & Active"
                category = "SYNCHRONIZED_ACTIVE"
                priority = 4
                diag_reason = f"Compliant: Active in Ninja and Enabled in AD ({ad_clinic})"
                rec_action = "Compliant - Domain object healthy"
                fill_st = fill_green
                font_st = font_green
        else:
            ad_status = "NOT IN AD"
            ad_clinic = "None"
            ad_logon_date = "Never Joined"
            sync_status = "CRITICAL: Missing from AD"
            category = "MISSING_FROM_AD"
            priority = 1
            diag_reason = "Device enrolled in Ninja RMM, but NO computer object found in Active Directory"
            rec_action = "Join computer to domain or verify computer rename"
            fill_st = fill_red
            font_st = font_red

        audit_devices.append({
            "Device": dev_name,
            "Location": loc,
            "Role": role,
            "User": user,
            "Uptime": up_days,
            "ADName": match_ad["Name"] if match_ad else "NOT IN AD",
            "ADStatus": ad_status,
            "ADClinic": ad_clinic,
            "ADLastLogon": ad_logon_date,
            "ADLastLogonDays": int(ad_days_num) if match_ad and ad_days_num is not None else "",
            "SyncStatus": sync_status,
            "Category": category,
            "Priority": priority,
            "DiagnosticReason": diag_reason,
            "RecommendedAction": rec_action,
            "StatusFill": fill_st,
            "StatusFont": font_st
        })

    # Find AD objects not managed in Ninja
    unmanaged_ad = []
    for a in ad_records:
        a_name_u = a["Name"].upper()
        a_name_norm = normalize_hostname(a["Name"])
        if a_name_u not in ninja_names and a_name_norm not in ninja_norms:
            if a["Status"].lower() in ["enabled", "true"]:
                if "Domain Controller" in a["DeviceType"] or "Server" in a["DeviceType"] or "Server" in a["OS"]:
                    role_ident = "Infrastructure / Server (Not in Workstation Inventory)"
                else:
                    role_ident = "Active Directory Workstation - Missing NinjaOne Agent"
                unmanaged_ad.append((a, role_ident))

    cnt_total_ninja = len(inventory_data)
    cnt_synced = sum(1 for d in audit_devices if d["Category"] in ["SYNCHRONIZED_ACTIVE", "OU_MISPLACED", "AD_LAGGING"])
    cnt_ad_disabled = sum(1 for d in audit_devices if d["Category"] == "AD_DISABLED")
    cnt_missing_ad = sum(1 for d in audit_devices if d["Category"] == "MISSING_FROM_AD")
    pct_sync = (cnt_synced / cnt_total_ninja * 100) if cnt_total_ninja > 0 else 0.0

    ws_ad_ref = create_tab(wb, sheet_title)
    ws_ad_ref.cell(1, 1).value = "ACTIVE DIRECTORY CROSS-REFERENCE & RECONCILIATION DASHBOARD"
    ws_ad_ref.cell(1, 1).font = font_title
    ws_ad_ref.cell(2, 1).value = "Cross-Referencing NinjaOne Fleet vs Active Directory Objects | Detecting Rogue, Disabled, & Misplaced Devices"
    ws_ad_ref.cell(2, 1).font = font_subtitle

    kpis_ad = [
        ("NINJA FLEET SIZE", f"{cnt_total_ninja:,} Units", "Active RMM Inventory", False),
        ("AD SYNCHRONIZED", f"{cnt_synced:,} Units", f"{pct_sync:.1f}% Match Rate", False),
        ("CRITICAL: AD DISABLED", f"{cnt_ad_disabled:,} Units", "Active Device / Disabled in AD", True),
        ("MISSING FROM AD", f"{cnt_missing_ad:,} Units", "In Ninja / Unjoined or Missing", True),
        ("AD OBJECTS NOT IN NINJA", f"{len(unmanaged_ad):,} Units", "Enabled in AD / No Ninja Agent", False)
    ]
    draw_kpi_cards(ws_ad_ref, kpis_ad)

    def _ad_sort_key(item):
        u_val = safe_float(item.get("Uptime"), 0.0)
        return (item["Priority"], u_val if item["Priority"] == 1 else -u_val)

    audit_devices.sort(key=_ad_sort_key)

    ws_ad_ref.cell(7, 1).value = f"Fleet Active Directory Synchronization Audit ({len(audit_devices)} Total Devices)"
    ws_ad_ref.cell(7, 1).font = font_section

    headers_ad_tbl = [
        "Device", "Ninja Location", "Device Role", "Last User", "Ninja Uptime (Days)",
        "AD Computer Name", "AD Status", "AD Clinic / OU", "AD Last Logon", "AD Last Logon (Days)",
        "Sync Status", "Diagnostic Notes", "Recommended Action"
    ]
    ws_ad_ref.row_dimensions[8].height = 25
    for c_idx, h in enumerate(headers_ad_tbl, 1):
        cell = ws_ad_ref.cell(8, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_navy_header
        cell.alignment = align_header
        cell.border = thin_border

    r_row = 9
    for item in audit_devices:
        ws_ad_ref.row_dimensions[r_row].height = 19
        is_odd = (r_row % 2 == 1)
        bg = fill_zebra if is_odd else fill_white

        row_vals = [
            (item["Device"], align_left, font_tbl_bold, item["StatusFill"] if item["Priority"] == 1 else bg),
            (item["Location"], align_left, font_tbl_data, bg),
            (item["Role"], align_center, font_tbl_data, bg),
            (item["User"], align_left, font_tbl_data, bg),
            (item["Uptime"], align_center, font_tbl_data, item["StatusFill"] if item["Priority"] in [1, 3] else bg),
            (item["ADName"], align_center, font_tbl_bold if item["ADName"] == "NOT IN AD" else font_tbl_data, item["StatusFill"] if item["Priority"] == 1 else bg),
            (item["ADStatus"], align_center, item["StatusFont"] if item["Priority"] == 1 else font_tbl_data, item["StatusFill"] if item["Priority"] == 1 else bg),
            (item["ADClinic"], align_left, font_tbl_data, bg),
            (item["ADLastLogon"], align_center, font_tbl_data, bg),
            (item["ADLastLogonDays"], align_right, font_tbl_data, bg),
            (item["SyncStatus"], align_center, item["StatusFont"], item["StatusFill"]),
            (item["DiagnosticReason"], align_left, font_tbl_bold if item["Priority"] in [1, 2, 3] else font_tbl_data, item["StatusFill"] if item["Priority"] == 1 else bg),
            (item["RecommendedAction"], align_left, item["StatusFont"] if item["Priority"] == 1 else font_tbl_bold, item["StatusFill"] if item["Priority"] == 1 else bg)
        ]

        for c_idx, (val, aln, fnt, fill) in enumerate(row_vals, 1):
            cell = ws_ad_ref.cell(r_row, c_idx, val)
            cell.alignment = aln
            cell.font = fnt
            cell.fill = fill
            cell.border = thin_border
            if c_idx in [5, 10] and isinstance(val, (int, float)):
                cell.number_format = '#,##0'

        r_row += 1

    # Table 2: Clinic AD Synchronization Summary
    ws_ad_ref.cell(7, 15).value = "Clinic AD Synchronization Summary"
    ws_ad_ref.cell(7, 15).font = font_section

    headers_loc = ["Clinic / Location", "Fleet Total", "AD Synced", "AD Disabled", "Missing in AD", "Sync %"]
    for c_idx, h in enumerate(headers_loc, 15):
        cell = ws_ad_ref.cell(8, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_sub_header
        cell.alignment = align_header
        cell.border = thin_border

    loc_stats = defaultdict(lambda: {"Total": 0, "Synced": 0, "Disabled": 0, "Missing": 0})
    for d in audit_devices:
        loc = d["Location"]
        loc_stats[loc]["Total"] += 1
        if d["Category"] in ["SYNCHRONIZED_ACTIVE", "OU_MISPLACED", "AD_LAGGING"]:
            loc_stats[loc]["Synced"] += 1
        elif d["Category"] == "AD_DISABLED":
            loc_stats[loc]["Disabled"] += 1
        elif d["Category"] == "MISSING_FROM_AD":
            loc_stats[loc]["Missing"] += 1

    sorted_locs = sorted(loc_stats.keys())
    r_loc = 9
    for loc in sorted_locs:
        st = loc_stats[loc]
        ws_ad_ref.row_dimensions[r_loc].height = 18
        bg_l = fill_zebra if r_loc % 2 == 1 else fill_white
        ws_ad_ref.cell(r_loc, 15, loc).font = font_tbl_bold
        ws_ad_ref.cell(r_loc, 15).fill = bg_l
        ws_ad_ref.cell(r_loc, 15).border = thin_border

        ws_ad_ref.cell(r_loc, 16, st["Total"]).font = font_tbl_data
        ws_ad_ref.cell(r_loc, 16).fill = bg_l
        ws_ad_ref.cell(r_loc, 16).alignment = align_center
        ws_ad_ref.cell(r_loc, 16).border = thin_border
        ws_ad_ref.cell(r_loc, 16).number_format = '#,##0'

        ws_ad_ref.cell(r_loc, 17, st["Synced"]).font = font_green if st["Synced"] > 0 else font_tbl_data
        ws_ad_ref.cell(r_loc, 17).fill = fill_green if st["Synced"] == st["Total"] and st["Total"] > 0 else bg_l
        ws_ad_ref.cell(r_loc, 17).alignment = align_center
        ws_ad_ref.cell(r_loc, 17).border = thin_border
        ws_ad_ref.cell(r_loc, 17).number_format = '#,##0'

        ws_ad_ref.cell(r_loc, 18, st["Disabled"]).font = font_red if st["Disabled"] > 0 else font_tbl_data
        ws_ad_ref.cell(r_loc, 18).fill = fill_red if st["Disabled"] > 0 else bg_l
        ws_ad_ref.cell(r_loc, 18).alignment = align_center
        ws_ad_ref.cell(r_loc, 18).border = thin_border
        ws_ad_ref.cell(r_loc, 18).number_format = '#,##0'

        ws_ad_ref.cell(r_loc, 19, st["Missing"]).font = font_red if st["Missing"] > 0 else font_tbl_data
        ws_ad_ref.cell(r_loc, 19).fill = fill_red if st["Missing"] > 0 else bg_l
        ws_ad_ref.cell(r_loc, 19).alignment = align_center
        ws_ad_ref.cell(r_loc, 19).border = thin_border
        ws_ad_ref.cell(r_loc, 19).number_format = '#,##0'

        cov_cell = ws_ad_ref.cell(r_loc, 20, f"=IF(P{r_loc}>0, Q{r_loc}/P{r_loc}, 0)")
        cov_cell.font = font_tbl_bold
        cov_cell.fill = bg_l
        cov_cell.alignment = align_center
        cov_cell.border = thin_border
        cov_cell.number_format = '0.0%'

        r_loc += 1

    # Total row for Table 2
    ws_ad_ref.cell(r_loc, 15, "FLEET TOTAL").font = font_tbl_total
    ws_ad_ref.cell(r_loc, 15).fill = fill_total
    ws_ad_ref.cell(r_loc, 15).border = total_border

    for c_offset, c_let in enumerate(["P", "Q", "R", "S"], 16):
        cell_t = ws_ad_ref.cell(r_loc, c_offset, f"=SUM({c_let}9:{c_let}{r_loc-1})")
        cell_t.font = font_tbl_total
        cell_t.fill = fill_total
        cell_t.border = total_border
        cell_t.alignment = align_center
        cell_t.number_format = '#,##0'

    cov_tot = ws_ad_ref.cell(r_loc, 20, f"=IF(P{r_loc}>0, Q{r_loc}/P{r_loc}, 0)")
    cov_tot.font = font_tbl_total
    cov_tot.fill = fill_total
    cov_tot.border = total_border
    cov_tot.alignment = align_center
    cov_tot.number_format = '0.0%'

    # Doughnut Chart for AD Status
    r_chart_data = r_loc + 2
    ws_ad_ref.cell(r_chart_data, 15, "AD Sync Status").font = font_tbl_header
    ws_ad_ref.cell(r_chart_data, 15).fill = fill_sub_header
    ws_ad_ref.cell(r_chart_data, 16, "Count").font = font_tbl_header
    ws_ad_ref.cell(r_chart_data, 16).fill = fill_sub_header

    chart_cats = [
        ("Synchronized & Active", cnt_synced),
        ("Critical: AD Disabled", cnt_ad_disabled),
        ("Missing from AD", cnt_missing_ad)
    ]
    for idx, (cat_name, cat_cnt) in enumerate(chart_cats, 1):
        ws_ad_ref.cell(r_chart_data + idx, 15, cat_name).font = font_tbl_bold
        ws_ad_ref.cell(r_chart_data + idx, 15).border = thin_border
        c_val = ws_ad_ref.cell(r_chart_data + idx, 16, cat_cnt)
        c_val.font = font_tbl_data
        c_val.alignment = align_center
        c_val.border = thin_border
        c_val.number_format = '#,##0'

    chart_ad = DoughnutChart()
    chart_ad.title = "Active Directory Synchronization Status"
    chart_ad.style = 10
    chart_ad.width = 15
    chart_ad.height = 8.5
    ad_chart_data = Reference(ws_ad_ref, min_col=16, min_row=r_chart_data, max_row=r_chart_data + len(chart_cats))
    ad_chart_cats = Reference(ws_ad_ref, min_col=15, min_row=r_chart_data + 1, max_row=r_chart_data + len(chart_cats))
    chart_ad.add_data(ad_chart_data, titles_from_data=True)
    chart_ad.set_categories(ad_chart_cats)
    ws_ad_ref.add_chart(chart_ad, f"O{r_chart_data + len(chart_cats) + 2}")

    # Table 3: AD Computers Not Enrolled in Ninja
    r_unm = r_row + 3
    ws_ad_ref.cell(r_unm, 1).value = f"Active Directory Computers Not Managed in NinjaOne Fleet ({len(unmanaged_ad)} Computers)"
    ws_ad_ref.cell(r_unm, 1).font = font_section
    ws_ad_ref.cell(r_unm + 1, 1).value = "Domain Controllers, Production Servers & Unenrolled Computers in Active Directory"
    ws_ad_ref.cell(r_unm + 1, 1).font = font_subtitle

    headers_unm = [
        "AD Computer Name", "AD Status", "Clinic", "Department", "Operating System",
        "Device Type", "Last Logon Date", "Last Logon (Days)", "Created Date", "Created (Days)",
        "Identified Role / Infrastructure Category"
    ]
    ws_ad_ref.row_dimensions[r_unm + 2].height = 25
    for c_idx, h in enumerate(headers_unm, 1):
        cell = ws_ad_ref.cell(r_unm + 2, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_accent_header
        cell.alignment = align_header
        cell.border = thin_border

    r_row_u = r_unm + 3
    for a, role_ident in unmanaged_ad:
        ws_ad_ref.row_dimensions[r_row_u].height = 18
        bg_u = fill_zebra if r_row_u % 2 == 1 else fill_white
        row_vals_u = [
            (a["Name"], align_left, font_tbl_bold, bg_u),
            (a["Status"], align_center, font_tbl_data, bg_u),
            (a["Clinic"], align_left, font_tbl_data, bg_u),
            (a["Department"], align_left, font_tbl_data, bg_u),
            (a["OS"], align_left, font_tbl_data, bg_u),
            (a["DeviceType"], align_center, font_tbl_data, bg_u),
            (a["LastLogonDate"], align_center, font_tbl_data, bg_u),
            (a["LastLogonDays"], align_right, font_tbl_data, bg_u),
            (a["CreatedDate"], align_center, font_tbl_data, bg_u),
            (a["CreatedDays"], align_right, font_tbl_data, bg_u),
            (role_ident, align_left, font_tbl_bold, fill_staging if "Server" in role_ident or "Infrastructure" in role_ident else fill_yellow)
        ]
        for c_idx, (val, aln, fnt, fill) in enumerate(row_vals_u, 1):
            cell = ws_ad_ref.cell(r_row_u, c_idx, val)
            cell.alignment = aln
            cell.font = fnt
            cell.fill = fill
            cell.border = thin_border
            if c_idx in [8, 10] and isinstance(val, (int, float)):
                cell.number_format = '#,##0'
        r_row_u += 1

    ws_widths = {
        1: 18, 2: 24, 3: 16, 4: 20, 5: 14, 6: 18, 7: 14, 8: 24, 9: 20, 10: 16,
        11: 24, 12: 66, 13: 57, 14: 3, 15: 28, 16: 12, 17: 12, 18: 14, 19: 14, 20: 14
    }
    for c_idx, w in ws_widths.items():
        ws_ad_ref.column_dimensions[get_column_letter(c_idx)].width = w
        ws_ad_ref.column_dimensions[get_column_letter(c_idx)].width = w

    ws_ad_ref.freeze_panes = "C9"

def _build_huntress_protection_tab(wb, inventory_data, total_fleet, sheet_title="Huntress Cross ref"):
    print(f"Building {sheet_title}: Huntress Endpoint Protection & Cross-Reference Dashboard...")
    ws_hunt = wb["Huntress Inventory"] if "Huntress Inventory" in wb.sheetnames else wb["Huntress Agents"]

    h_headers = {ws_hunt.cell(1, c).value: c for c in range(1, ws_hunt.max_column + 1)}
    c_h_agent = h_headers.get("Agent", 1)
    c_h_org = h_headers.get("Organization", 2)
    c_h_os = h_headers.get("Operating System", 3)
    c_h_aver = h_headers.get("Agent Version", 4)
    c_h_ever = h_headers.get("EDR Version", 5)
    c_h_intip = h_headers.get("Internal IP", 7)
    c_h_reg = h_headers.get("Registered", 8)
    c_h_regdays = h_headers.get("Register - Days") or h_headers.get("Registered - Days", 9)
    c_h_seen = h_headers.get("Last Seen", 10)
    c_h_seendays = h_headers.get("Last Seen - Days") or h_headers.get("Days Inactive", 11)
    c_h_sn = h_headers.get("Serial Number", 12)

    huntress_records = []
    for r in range(2, ws_hunt.max_row + 1):
        ag_name = ws_hunt.cell(r, c_h_agent).value
        if not ag_name:
            continue
        huntress_records.append({
            "Agent": str(ag_name).strip(),
            "Organization": str(ws_hunt.cell(r, c_h_org).value or "") if c_h_org else "",
            "OS": str(ws_hunt.cell(r, c_h_os).value or "") if c_h_os else "",
            "AgentVersion": str(ws_hunt.cell(r, c_h_aver).value or "") if c_h_aver else "",
            "EDRVersion": str(ws_hunt.cell(r, c_h_ever).value or "") if c_h_ever else "",
            "InternalIP": str(ws_hunt.cell(r, c_h_intip).value or "") if c_h_intip else "",
            "Registered": str(ws_hunt.cell(r, c_h_reg).value or "") if c_h_reg else "",
            "RegisteredDays": ws_hunt.cell(r, c_h_regdays).value if c_h_regdays else None,
            "LastSeen": str(ws_hunt.cell(r, c_h_seen).value or "") if c_h_seen else "",
            "DaysInactive": ws_hunt.cell(r, c_h_seendays).value if c_h_seendays else None,
            "Serial": str(ws_hunt.cell(r, c_h_sn).value or "").strip() if c_h_sn else ""
        })

    h_by_name = {a["Agent"].upper(): a for a in huntress_records}
    h_by_norm = {normalize_hostname(a["Agent"]): a for a in huntress_records if a.get("Agent")}
    h_by_sn = {a["Serial"].upper(): a for a in huntress_records if is_valid_serial(a.get("Serial"))}

    ninja_names = {str(d["Device"]).strip().upper() for d in inventory_data if d.get("Device")}
    ninja_norms = {normalize_hostname(d["Device"]) for d in inventory_data if d.get("Device")}
    ninja_sns = {str(d.get("Serial") or "").strip().upper() for d in inventory_data if is_valid_serial(d.get("Serial"))}

    audit_devices = []
    for d in inventory_data:
        dev_name = str(d["Device"]).strip()
        dev_sn = str(d.get("Serial") or "").strip()
        up_days = d.get("Uptime") or 0
        loc = config.canonicalize_clinic_name(d.get("Location") or "Unassigned")
        role = d.get("Role") or "Workstation"
        model = d.get("Model") or "Unknown"
        user = d.get("User") or "N/A"

        # Prioritize Serial Number matching when valid; fall back to Hostname for VMs/generic serials
        match_h = None
        if is_valid_serial(dev_sn) and dev_sn.upper() in h_by_sn:
            match_h = h_by_sn[dev_sn.upper()]
        else:
            # Fall back to hostname if serial number is missing/invalid,
            # or if candidate Huntress agent does NOT have a conflicting valid serial
            candidate = h_by_name.get(dev_name.upper()) or h_by_norm.get(normalize_hostname(dev_name))
            if candidate:
                cand_sn = str(candidate.get("Serial") or "").strip().upper()
                if not is_valid_serial(dev_sn) or not is_valid_serial(cand_sn):
                    match_h = candidate

        if match_h:
            h_agent = match_h["Agent"]
            h_edr = match_h["EDRVersion"] or match_h["AgentVersion"]
            h_seen = match_h["LastSeen"]
            h_inact = match_h["DaysInactive"] if match_h["DaysInactive"] is not None else 0
            try:
                h_inact_num = float(h_inact)
            except (ValueError, TypeError):
                h_inact_num = 0.0

            if h_inact_num <= 30:
                prot_status = "Protected & Active"
                category = "PROTECTED_ACTIVE"
                priority = 5
                diag_reason = f"Compliant: Agent reporting normally (last seen {int(h_inact_num)}d ago)"
                rec_action = "Compliant - Endpoint protected"
                fill_st = fill_green
                font_st = font_green
            elif h_inact_num <= 90:
                prot_status = "Protected (Lagging)"
                category = "PROTECTED_LAGGING"
                priority = 4
                diag_reason = f"Warning: Agent silent for {int(h_inact_num)} days (31-90d window)"
                rec_action = "Verify device connectivity & Huntress service"
                fill_st = fill_yellow
                font_st = font_yellow
            else:
                prot_status = "At Risk of Inactivity Drop"
                category = "HUNTRESS_DROP_RISK"
                priority = 2
                diag_reason = f"Silent for {int(h_inact_num)} days: approaching Huntress removal limit"
                rec_action = "Power on device / verify agent connectivity immediately"
                fill_st = fill_red
                font_st = font_red
        else:
            h_agent = "NOT INSTALLED"
            h_edr = "None"
            h_seen = "Never Enrolled"

            if up_days <= 90:
                prot_status = "CRITICAL: Missing Protection"
                category = "MISSING_ONLINE"
                priority = 1
                if up_days <= 7:
                    diag_reason = "Device is in NinjaOne but NEVER ADDED to Huntress endpoint protection"
                elif up_days <= 30:
                    diag_reason = f"Active device ({up_days}d uptime) with NO Huntress endpoint protection"
                else:
                    diag_reason = f"Active device ({up_days}d uptime <=90d) missing Huntress agent"
                rec_action = "Deploy Huntress Agent immediately via NinjaOne automation"
                fill_st = fill_red
                font_st = font_red
            else:
                prot_status = "Dropped from Inactivity (>90d)"
                category = "DROPPED_INACTIVE"
                priority = 3
                if up_days > 365:
                    diag_reason = f"Past 90 days (offline {up_days}d) - Dropped from inactivity / Dormant"
                    rec_action = "Decommission asset & retire AD account"
                    fill_st = fill_purple
                    font_st = font_purple
                else:
                    diag_reason = f"Past 90 days (offline {up_days}d) - Dropped from inactivity by Huntress"
                    rec_action = "Locate hardware; if returning to service, reinstall Huntress"
                    fill_st = fill_purple
                    font_st = font_purple

        audit_devices.append({
            "Device": dev_name,
            "Location": loc,
            "Role": role,
            "Model": model,
            "User": user,
            "Uptime": up_days,
            "HuntressAgent": h_agent,
            "EDRVersion": h_edr,
            "LastSeen": h_seen,
            "LastSeenDays": int(h_inact_num) if match_h and h_inact_num is not None else "",
            "ProtectionStatus": prot_status,
            "Category": category,
            "Priority": priority,
            "DiagnosticReason": diag_reason,
            "RecommendedAction": rec_action,
            "StatusFill": fill_st,
            "StatusFont": font_st
        })

    unmanaged_huntress = []
    for a in huntress_records:
        ag_u = a["Agent"].upper()
        ag_norm = normalize_hostname(a["Agent"])
        sn_u = a["Serial"].upper() if is_valid_serial(a.get("Serial")) else ""
        # Prioritize Serial Number check; fall back to Hostname if no valid serial exists
        if sn_u:
            is_managed = sn_u in ninja_sns
        else:
            is_managed = (ag_u in ninja_names) or (ag_norm in ninja_norms)

        if not is_managed:
            os_name = a["OS"]
            if "Server" in os_name:
                role_ident = "Windows Server - Infrastructure (Not in Workstation Inventory)"
            elif "macOS" in os_name or "Mac" in os_name:
                role_ident = "macOS Device - Standalone / Unmanaged in Ninja"
            elif "Windows 11" in os_name or "Windows 10" in os_name:
                role_ident = "Workstation - Enrolled in Huntress but Missing from NinjaOne"
            else:
                role_ident = "Standalone Endpoint - Missing from NinjaOne RMM"
            unmanaged_huntress.append((a, role_ident))

    cnt_total_ninja = len(inventory_data)
    cnt_protected_active = sum(1 for d in audit_devices if d["Category"] == "PROTECTED_ACTIVE")
    cnt_protected_lagging = sum(1 for d in audit_devices if d["Category"] == "PROTECTED_LAGGING")
    cnt_total_protected = cnt_protected_active + cnt_protected_lagging
    cnt_missing_online = sum(1 for d in audit_devices if d["Category"] == "MISSING_ONLINE")
    cnt_dropped_inact = sum(1 for d in audit_devices if d["Category"] == "DROPPED_INACTIVE")
    cnt_drop_risk = sum(1 for d in audit_devices if d["Category"] == "HUNTRESS_DROP_RISK")
    pct_cov = (cnt_total_protected / cnt_total_ninja * 100) if cnt_total_ninja > 0 else 0.0

    ws6 = create_tab(wb, sheet_title)
    ws6.cell(1, 1).value = "HUNTRESS ENDPOINT PROTECTION & COMPLIANCE DASHBOARD"
    ws6.cell(1, 1).font = font_title
    ws6.cell(2, 1).value = "Cross-Referencing NinjaOne Fleet vs Huntress EDR Agents | Tracking Missing Agents & Inactivity Drops"
    ws6.cell(2, 1).font = font_subtitle

    kpis6 = [
        ("NINJA FLEET SIZE", f"{cnt_total_ninja:,} Units", "Active RMM Inventory", False),
        ("PROTECTED FLEET", f"{cnt_total_protected:,} Units", f"{pct_cov:.1f}% EDR Coverage", False),
        ("CRITICAL GAP (ONLINE)", f"{cnt_missing_online:,} Units", "Active Device / No Agent", True),
        ("DROPPED (>90D INACTIVE)", f"{cnt_dropped_inact:,} Units", "Inactivity Removal Window", True),
        ("UNMANAGED IN NINJA", f"{len(unmanaged_huntress):,} Units", "Huntress Servers / Standalone", False)
    ]
    draw_kpi_cards(ws6, kpis6)

    def _hp_sort_key(item):
        u_val = safe_float(item.get("Uptime"), 0.0)
        return (item["Priority"], u_val if item["Priority"] == 1 else -u_val)

    audit_devices.sort(key=_hp_sort_key)

    ws6.cell(7, 1).value = f"Endpoint Protection Status & Inactivity Diagnostic Audit ({len(audit_devices)} Total Devices)"
    ws6.cell(7, 1).font = font_section

    headers_hp = [
        "Device", "Location", "Device Role", "Device Model", "Last User",
        "Uptime (Days)", "Huntress Agent Name", "EDR Version", "Huntress Last Seen",
        "Last Seen - Days", "Protection Status", "Root Cause / Diagnostic Notes", "Recommended Action"
    ]
    ws6.row_dimensions[8].height = 25
    for c_idx, h in enumerate(headers_hp, 1):
        cell = ws6.cell(8, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_navy_header
        cell.alignment = align_header
        cell.border = thin_border

    r_hp = 9
    for item in audit_devices:
        ws6.row_dimensions[r_hp].height = 19
        is_odd = (r_hp % 2 == 1)
        bg = fill_zebra if is_odd else fill_white

        row_vals = [
            (item["Device"], align_left, font_tbl_bold, item["StatusFill"] if item["Priority"] == 1 else bg),
            (item["Location"], align_left, font_tbl_data, bg),
            (item["Role"], align_center, font_tbl_data, bg),
            (item["Model"], align_left, font_tbl_data, bg),
            (item["User"], align_left, font_tbl_data, bg),
            (item["Uptime"], align_center, font_tbl_data, item["StatusFill"] if item["Priority"] in [1, 3] else bg),
            (item["HuntressAgent"], align_center, font_tbl_bold if item["HuntressAgent"] == "NOT INSTALLED" else font_tbl_data, item["StatusFill"] if item["Priority"] == 1 else bg),
            (item["EDRVersion"], align_center, font_tbl_data, bg),
            (item["LastSeen"], align_center, font_tbl_data, bg),
            (item["LastSeenDays"], align_right, font_tbl_data, bg),
            (item["ProtectionStatus"], align_center, item["StatusFont"], item["StatusFill"]),
            (item["DiagnosticReason"], align_left, font_tbl_bold if item["Priority"] in [1, 2, 3] else font_tbl_data, item["StatusFill"] if item["Priority"] == 1 else bg),
            (item["RecommendedAction"], align_left, item["StatusFont"] if item["Priority"] == 1 else font_tbl_bold, item["StatusFill"] if item["Priority"] == 1 else bg)
        ]

        for c_idx, (val, aln, fnt, fill) in enumerate(row_vals, 1):
            cell = ws6.cell(r_hp, c_idx, val)
            cell.alignment = aln
            cell.font = fnt
            cell.fill = fill
            cell.border = thin_border
            if c_idx in [6, 10] and isinstance(val, (int, float)):
                cell.number_format = '#,##0'

        r_hp += 1

    # Table 2: Clinic Protection Breakdown
    ws6.cell(7, 15).value = "Clinic Protection & Coverage Summary"
    ws6.cell(7, 15).font = font_section

    headers_loc = ["Clinic / Location", "Fleet Total", "Protected", "Missing (Online)", "Dropped (>90d)", "Coverage %"]
    for c_idx, h in enumerate(headers_loc, 15):
        cell = ws6.cell(8, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_sub_header
        cell.alignment = align_header
        cell.border = thin_border

    loc_stats = defaultdict(lambda: {"Total": 0, "Protected": 0, "MissingOnline": 0, "Dropped": 0})
    for d in audit_devices:
        loc = d["Location"]
        loc_stats[loc]["Total"] += 1
        if d["Category"] in ["PROTECTED_ACTIVE", "PROTECTED_LAGGING"]:
            loc_stats[loc]["Protected"] += 1
        elif d["Category"] == "MISSING_ONLINE":
            loc_stats[loc]["MissingOnline"] += 1
        elif d["Category"] == "DROPPED_INACTIVE":
            loc_stats[loc]["Dropped"] += 1

    sorted_locs = sorted(loc_stats.keys())
    r_loc = 9
    for loc in sorted_locs:
        st = loc_stats[loc]
        ws6.row_dimensions[r_loc].height = 18
        bg_l = fill_zebra if r_loc % 2 == 1 else fill_white
        ws6.cell(r_loc, 15, loc).font = font_tbl_bold
        ws6.cell(r_loc, 15).fill = bg_l
        ws6.cell(r_loc, 15).border = thin_border

        ws6.cell(r_loc, 16, st["Total"]).font = font_tbl_data
        ws6.cell(r_loc, 16).fill = bg_l
        ws6.cell(r_loc, 16).alignment = align_center
        ws6.cell(r_loc, 16).border = thin_border
        ws6.cell(r_loc, 16).number_format = '#,##0'

        ws6.cell(r_loc, 17, st["Protected"]).font = font_green if st["Protected"] > 0 else font_tbl_data
        ws6.cell(r_loc, 17).fill = fill_green if st["Protected"] == st["Total"] and st["Total"] > 0 else bg_l
        ws6.cell(r_loc, 17).alignment = align_center
        ws6.cell(r_loc, 17).border = thin_border
        ws6.cell(r_loc, 17).number_format = '#,##0'

        ws6.cell(r_loc, 18, st["MissingOnline"]).font = font_red if st["MissingOnline"] > 0 else font_tbl_data
        ws6.cell(r_loc, 18).fill = fill_red if st["MissingOnline"] > 0 else bg_l
        ws6.cell(r_loc, 18).alignment = align_center
        ws6.cell(r_loc, 18).border = thin_border
        ws6.cell(r_loc, 18).number_format = '#,##0'

        ws6.cell(r_loc, 19, st["Dropped"]).font = font_purple if st["Dropped"] > 0 else font_tbl_data
        ws6.cell(r_loc, 19).fill = fill_purple if st["Dropped"] > 0 else bg_l
        ws6.cell(r_loc, 19).alignment = align_center
        ws6.cell(r_loc, 19).border = thin_border
        ws6.cell(r_loc, 19).number_format = '#,##0'

        cov_cell = ws6.cell(r_loc, 20, f"=IF(P{r_loc}>0, Q{r_loc}/P{r_loc}, 0)")
        cov_cell.font = font_tbl_bold
        cov_cell.fill = bg_l
        cov_cell.alignment = align_center
        cov_cell.border = thin_border
        cov_cell.number_format = '0.0%'

        r_loc += 1

    # Total row for Table 2
    ws6.cell(r_loc, 15, "FLEET TOTAL").font = font_tbl_total
    ws6.cell(r_loc, 15).fill = fill_total
    ws6.cell(r_loc, 15).border = total_border

    ws6.cell(r_loc, 16, f"=SUM(P9:P{r_loc-1})").font = font_tbl_total
    ws6.cell(r_loc, 16).fill = fill_total
    ws6.cell(r_loc, 16).border = total_border
    ws6.cell(r_loc, 16).alignment = align_center
    ws6.cell(r_loc, 16).number_format = '#,##0'

    ws6.cell(r_loc, 17, f"=SUM(Q9:Q{r_loc-1})").font = font_tbl_total
    ws6.cell(r_loc, 17).fill = fill_total
    ws6.cell(r_loc, 17).border = total_border
    ws6.cell(r_loc, 17).alignment = align_center
    ws6.cell(r_loc, 17).number_format = '#,##0'

    ws6.cell(r_loc, 18, f"=SUM(R9:R{r_loc-1})").font = font_tbl_total
    ws6.cell(r_loc, 18).fill = fill_total
    ws6.cell(r_loc, 18).border = total_border
    ws6.cell(r_loc, 18).alignment = align_center
    ws6.cell(r_loc, 18).number_format = '#,##0'

    ws6.cell(r_loc, 19, f"=SUM(S9:S{r_loc-1})").font = font_tbl_total
    ws6.cell(r_loc, 19).fill = fill_total
    ws6.cell(r_loc, 19).border = total_border
    ws6.cell(r_loc, 19).alignment = align_center
    ws6.cell(r_loc, 19).number_format = '#,##0'

    ws6.cell(r_loc, 20, f"=IF(P{r_loc}>0, Q{r_loc}/P{r_loc}, 0)").font = font_tbl_total
    ws6.cell(r_loc, 20).fill = fill_total
    ws6.cell(r_loc, 20).border = total_border
    ws6.cell(r_loc, 20).alignment = align_center
    ws6.cell(r_loc, 20).number_format = '0.0%'

    # Protection Status Summary & Doughnut Chart
    r_chart_data = r_loc + 2
    ws6.cell(r_chart_data, 15, "Protection Status").font = font_tbl_header
    ws6.cell(r_chart_data, 15).fill = fill_sub_header
    ws6.cell(r_chart_data, 16, "Device Count").font = font_tbl_header
    ws6.cell(r_chart_data, 16).fill = fill_sub_header

    chart_cats = [
        ("Protected & Active", cnt_total_protected),
        ("Critical Missing (Online)", cnt_missing_online),
        ("Dropped from Inactivity (>90d)", cnt_dropped_inact)
    ]
    for idx, (cat_name, cat_cnt) in enumerate(chart_cats, 1):
        ws6.cell(r_chart_data + idx, 15, cat_name).font = font_tbl_bold
        ws6.cell(r_chart_data + idx, 15).border = thin_border
        c_val = ws6.cell(r_chart_data + idx, 16, cat_cnt)
        c_val.font = font_tbl_data
        c_val.alignment = align_center
        c_val.border = thin_border
        c_val.number_format = '#,##0'

    chart_hp = DoughnutChart()
    chart_hp.title = "Huntress Fleet Protection Status"
    chart_hp.style = 10
    chart_hp.width = 15
    chart_hp.height = 8.5
    hp_data = Reference(ws6, min_col=16, min_row=r_chart_data, max_row=r_chart_data + len(chart_cats))
    hp_cats = Reference(ws6, min_col=15, min_row=r_chart_data + 1, max_row=r_chart_data + len(chart_cats))
    chart_hp.add_data(hp_data, titles_from_data=True)
    chart_hp.set_categories(hp_cats)
    ws6.add_chart(chart_hp, f"O{r_chart_data + len(chart_cats) + 2}")

    # Table 3: Huntress Endpoints Not Enrolled in Ninja
    r_unm = r_hp + 3
    ws6.cell(r_unm, 1).value = f"Huntress Endpoints Not Managed in NinjaOne Workstation Fleet ({len(unmanaged_huntress)} Endpoints)"
    ws6.cell(r_unm, 1).font = font_section
    ws6.cell(r_unm + 1, 1).value = "Domain Controllers, Production Servers & Standalone Systems Enrolled in Huntress EDR"
    ws6.cell(r_unm + 1, 1).font = font_subtitle

    headers_unm = [
        "Huntress Agent", "Organization", "Operating System", "Agent Version",
        "EDR Version", "Internal IP", "Registered", "Register - Days",
        "Last Seen", "Last Seen - Days", "Serial Number",
        "Identified Role / Infrastructure Category"
    ]
    ws6.row_dimensions[r_unm + 2].height = 25
    for c_idx, h in enumerate(headers_unm, 1):
        cell = ws6.cell(r_unm + 2, c_idx, h)
        cell.font = font_tbl_header
        cell.fill = fill_accent_header
        cell.alignment = align_header
        cell.border = thin_border

    r_row_u = r_unm + 3
    for a, role_ident in unmanaged_huntress:
        ws6.row_dimensions[r_row_u].height = 18
        bg_u = fill_zebra if r_row_u % 2 == 1 else fill_white

        reg_days_val = a.get("RegisteredDays")
        try:
            reg_days_num = int(float(reg_days_val)) if reg_days_val is not None and str(reg_days_val).strip() != "" else ""
        except (ValueError, TypeError):
            reg_days_num = ""

        seen_days_val = a.get("DaysInactive")
        try:
            seen_days_num = int(float(seen_days_val)) if seen_days_val is not None and str(seen_days_val).strip() != "" else ""
        except (ValueError, TypeError):
            seen_days_num = ""

        row_vals_u = [
            (a["Agent"], align_left, font_tbl_bold, bg_u),
            (a["Organization"], align_left, font_tbl_data, bg_u),
            (a["OS"], align_left, font_tbl_data, bg_u),
            (a["AgentVersion"], align_center, font_tbl_data, bg_u),
            (a["EDRVersion"], align_center, font_tbl_data, bg_u),
            (a["InternalIP"], align_center, font_tbl_data, bg_u),
            (a["Registered"], align_center, font_tbl_data, bg_u),
            (reg_days_num, align_right, font_tbl_data, bg_u),
            (a["LastSeen"], align_center, font_tbl_data, bg_u),
            (seen_days_num, align_right, font_tbl_data, bg_u),
            (a["Serial"], align_center, font_tbl_data, bg_u),
            (role_ident, align_left, font_tbl_bold, fill_staging if "Server" in role_ident else fill_yellow)
        ]
        for c_idx, (val, aln, fnt, fill) in enumerate(row_vals_u, 1):
            cell = ws6.cell(r_row_u, c_idx, val)
            cell.alignment = aln
            cell.font = fnt
            cell.fill = fill
            cell.border = thin_border
            if c_idx in [8, 10] and isinstance(val, (int, float)):
                cell.number_format = '#,##0'
        r_row_u += 1

    ws6_widths = {
        1: 18, 2: 24, 3: 16, 4: 24, 5: 22, 6: 14, 7: 20, 8: 14, 9: 20, 10: 16,
        11: 26, 12: 61, 13: 48, 14: 3, 15: 28, 16: 12, 17: 12, 18: 16, 19: 16, 20: 14
    }
    for c_idx, w in ws6_widths.items():
        ws6.column_dimensions[get_column_letter(c_idx)].width = w

    ws6.freeze_panes = "C9"

def generate_dashboards(wb_or_path, huntress_csv=None, skip_lifecycle=False):
    """
    Main entry point to load a workbook and generate all executive dashboard tabs.
    Accepts either an openpyxl Workbook instance or a file path string.
    """
    is_path = isinstance(wb_or_path, str)
    if is_path:
        print(f"Loading workbook: {wb_or_path}...")
        wb = openpyxl.load_workbook(wb_or_path)
    else:
        wb = wb_or_path

    if skip_lifecycle:
        print("--no-ad flag detected: Lifecycle tab will be skipped (no AD Install Date data).")

    if "Ninja Inventory" in wb.sheetnames:
        master_ws = wb["Ninja Inventory"]
    elif "Inventory" in wb.sheetnames:
        master_ws = wb["Inventory"]
    else:
        master_ws = wb.active

    inventory_data = _load_inventory_data(master_ws)
    total_fleet = len(inventory_data)
    clinic_map, years = _compute_clinic_map(inventory_data)

    _build_patch_reboot_tab(wb, inventory_data, total_fleet)

    if not skip_lifecycle:
        _build_lifecycle_tab(wb, inventory_data, total_fleet, clinic_map, years=years)
    else:
        print("Skipping Lifecycle tab (no AD Install Date data available).")

    _build_performance_capacity_tab(wb, inventory_data, total_fleet)
    _build_site_allocation_tab(wb, inventory_data, total_fleet, clinic_map)
    _build_stale_assets_tab(wb, inventory_data, total_fleet)

    if "AD Inventory" in wb.sheetnames or "AD Computers" in wb.sheetnames:
        _build_ad_cross_ref_tab(wb, inventory_data, total_fleet, sheet_title="AD Cross ref")

    if "Huntress Inventory" in wb.sheetnames or "Huntress Agents" in wb.sheetnames:
        _build_huntress_protection_tab(wb, inventory_data, total_fleet, sheet_title="Huntress Cross ref")

    # Re-order sheets preserving exact target sequence
    sheet_order = [
        "Ninja Inventory",
        "AD Inventory",
        "AD Computers",
        "Huntress Inventory",
        "Huntress Agents",
        "Patch & Reboot Health",
        "Lifecycle",
        "Performance & Capacity",
        "Site Allocation",
        "Stale Assets",
        "AD Dashboard",
        "AD Cross ref",
        "Huntress Dashboard",
        "Huntress Cross ref",
        "Inventory",
        "Dashboard",
        "Active Directory",
        "Huntress Endpoint Protection",
        "Executive Summary"
    ]
    ordered_sheets = [wb[s] for s in sheet_order if s in wb.sheetnames]
    for s in wb.sheetnames:
        if wb[s] not in ordered_sheets:
            ordered_sheets.append(wb[s])
    wb._sheets = ordered_sheets

    if "Ninja Inventory" in wb.sheetnames:
        wb.active = wb["Ninja Inventory"]
    elif "Inventory" in wb.sheetnames:
        wb.active = wb["Inventory"]

    if is_path:
        print(f"Saving updated workbook to {wb_or_path}...")
        wb.save(wb_or_path)
    print("Dashboard tabs successfully added and verified!")
    return wb

def main(argv=None):
    if argv is None:
        argv = sys.argv[1:]

    if len(argv) > 0 and not argv[0].startswith("--"):
        wb_target = argv[0]
    else:
        candidates = glob.glob("Ninja Inventory*.xlsx") or glob.glob(os.path.join("reports", "Ninja Reports", "Ninja Inventory*.xlsx"))
        wb_target = sorted(candidates, key=os.path.getmtime, reverse=True)[0] if candidates else "Ninja Inventory.xlsx"

    cli_skip_lifecycle = "--no-ad" in argv
    cli_huntress = None
    for i, a in enumerate(argv):
        if a == "--huntress" and i + 1 < len(argv):
            cli_huntress = argv[i + 1]

    generate_dashboards(wb_target, huntress_csv=cli_huntress, skip_lifecycle=cli_skip_lifecycle)

if __name__ == "__main__":
    main()
