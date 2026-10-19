"""
build_huntress_report.py
========================
Generates an executive-grade Excel report directly from Huntress Agents CSV export.
Worksheet order:
  Tab 1: Huntress Agents (Master Agent Inventory Table with Auto-Filters, Clean Typography & Styling)
  Tab 2: Executive Summary & Audit Dashboard (KPI Cards, Inactivity Tiers, OS Breakdown & Dormancy Watchlist)
"""

import os
import sys
import csv
import glob
from datetime import datetime
from collections import Counter, defaultdict

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.chart import DoughnutChart, BarChart, Reference
from openpyxl.worksheet.table import Table, TableStyleInfo

import config
from utils import parse_date, create_kpi_card

HUNTRESS_HEADERS = config.HUNTRESS_STANDALONE_HEADERS
HUNTRESS_COL_WIDTHS = config.HUNTRESS_STANDALONE_COL_WIDTHS

def parse_huntress_csv(csv_path):
    today_date = datetime.now().date()
    agents = []
    parsed_seen_dates = []

    with open(csv_path, mode="r", encoding="utf-8-sig", errors="ignore") as f:
        reader = csv.reader(f)
        first_row = next(reader, None)
        if not first_row:
            raise ValueError("Huntress CSV file is empty.")
        h_head = [c.replace("\ufeff", "").strip() for c in first_row]
        h_map = {name: i for i, name in enumerate(h_head)}

        for row in reader:
            if not row or not any(row):
                continue
            def get_h(col):
                idx = h_map.get(col)
                return row[idx].strip() if idx is not None and idx < len(row) else ""

            agent_name = get_h('Agent')
            if not agent_name:
                continue

            last_seen_str = get_h('Last Seen')
            last_seen_date = parse_date(last_seen_str)
            if last_seen_date:
                parsed_seen_dates.append(last_seen_date)

            agents.append({
                "Agent": agent_name,
                "Organization": get_h('Organization'),
                "OperatingSystem": get_h('Operating System'),
                "AgentVersion": get_h('Agent Version'),
                "EDRVersion": get_h('EDR Version'),
                "MacEDREnabled": get_h('MacOS EDR enabled'),
                "InternalIP": get_h('Internal IP'),
                "ExternalIP": get_h('External IP'),
                "Registered": get_h('Registered'),
                "LastSeen": last_seen_str,
                "LastSeenDate": last_seen_date,
                "SerialNumber": get_h('Serial Number'),
                "ADDomain": get_h('AD Domain')
            })

    ref_date = max(today_date, max(parsed_seen_dates)) if parsed_seen_dates else today_date

    # Compute days inactive
    for a in agents:
        if a["LastSeenDate"]:
            a["DaysInactive"] = max(0, (ref_date - a["LastSeenDate"]).days)
        else:
            a["DaysInactive"] = None

    return agents

def build_huntress_inventory_tab(wb, agents, title="Huntress Agents"):
    """
    Builds the Huntress Agents master inventory tab into the provided workbook.
    """
    if title in wb.sheetnames:
        ws_agents = wb[title]
    elif len(wb.sheetnames) == 1 and wb.active.title == "Sheet":
        ws_agents = wb.active
        ws_agents.title = title
    else:
        ws_agents = wb.create_sheet(title=title)
        
    ws_agents.views.sheetView[0].showGridLines = True

    for c_idx, h in enumerate(HUNTRESS_HEADERS, 1):
        cell = ws_agents.cell(1, c_idx, h)
        cell.font = config.FONT_TBL_HEADER
        cell.fill = config.FILL_NAVY_HEADER
        cell.alignment = config.ALIGN_HEADER
        cell.border = config.BORDER_THIN
    ws_agents.row_dimensions[1].height = 28

    for r_idx, a in enumerate(agents, 2):
        bg = config.FILL_ZEBRA_ROW if r_idx % 2 == 0 else config.FILL_WHITE_ROW
        row_vals = [
            a["Agent"], a["Organization"], a["OperatingSystem"], a["AgentVersion"],
            a["EDRVersion"], a["MacEDREnabled"], a["InternalIP"], a["ExternalIP"],
            a["Registered"], a["LastSeen"], a["SerialNumber"], a["ADDomain"], a["DaysInactive"]
        ]
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws_agents.cell(r_idx, c_idx, val)
            cell.font = config.FONT_TBL_DATA
            cell.fill = bg
            cell.border = config.BORDER_THIN
            cell.alignment = config.ALIGN_LEFT if c_idx in [1, 2, 3] else config.ALIGN_CENTER

        # Days Inactive alert styling
        d_val = a["DaysInactive"]
        if d_val is not None:
            c_d = ws_agents.cell(r_idx, 13)
            c_d.number_format = '#,##0'
            if d_val > 90:
                c_d.fill = config.FILL_LAVENDER
                c_d.font = config.FONT_LAVENDER
            elif d_val > 30:
                c_d.fill = config.FILL_ALERT_WARN
                c_d.font = config.FONT_ALERT_WARN

    for c_idx, width in HUNTRESS_COL_WIDTHS.items():
        ws_agents.column_dimensions[get_column_letter(c_idx)].width = width

    if ws_agents.max_row > 1:
        tbl_hunt = Table(displayName="Table_HuntressAgents", ref=f"A1:M{ws_agents.max_row}")
        tbl_hunt.tableStyleInfo = TableStyleInfo(
            name="TableStyleLight8",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        ws_agents.add_table(tbl_hunt)
    ws_agents.freeze_panes = "B2"
    return ws_agents

def build_huntress_dashboard_tab(wb, agents, title="Executive Summary"):
    """
    Builds the Huntress Executive Summary & Audit Dashboard tab into the provided workbook.
    """
    total_agents = len(agents)
    active_30 = sum(1 for a in agents if a["DaysInactive"] is not None and a["DaysInactive"] <= 30)
    inactive_90 = sum(1 for a in agents if a["DaysInactive"] is not None and 30 < a["DaysInactive"] <= 90)
    dormant = sum(1 for a in agents if a["DaysInactive"] is not None and a["DaysInactive"] > 90)

    win_count = sum(1 for a in agents if "Windows" in a["OperatingSystem"])
    mac_count = sum(1 for a in agents if "Mac" in a["OperatingSystem"] or "OS X" in a["OperatingSystem"] or "Darwin" in a["OperatingSystem"])

    if title in wb.sheetnames:
        ws_sum = wb[title]
    else:
        ws_sum = wb.create_sheet(title=title)
        
    ws_sum.views.sheetView[0].showGridLines = True

    # Title & Subtitle
    ws_sum.cell(2, 2, "HUNTRESS ENDPOINT PROTECTION AUDIT").font = config.FONT_TITLE
    curr_my = datetime.now().strftime("%B %Y")
    ws_sum.cell(3, 2, f"Agent Deployment Health, Operating Systems & Dormancy Analysis | {curr_my}").font = config.FONT_SUBTITLE

    # KPI Cards (Row 5 - 6)
    create_kpi_card(ws_sum, 2, 5, 3, 6, "TOTAL AGENTS", f"{total_agents:,}")
    create_kpi_card(ws_sum, 4, 5, 5, 6, "ACTIVE (<=30D)", f"{active_30:,}")
    create_kpi_card(ws_sum, 6, 5, 7, 6, "INACTIVE (31-90D)", f"{inactive_90:,}", is_alert=(inactive_90 > 0))
    create_kpi_card(ws_sum, 8, 5, 9, 6, "DORMANT (>90D)", f"{dormant:,}", is_alert=(dormant > 0))
    create_kpi_card(ws_sum, 10, 5, 11, 6, "WINDOWS AGENTS", f"{win_count:,}")
    create_kpi_card(ws_sum, 12, 5, 13, 6, "MACOS AGENTS", f"{mac_count:,}")

    # Section 1: Inactivity Severity Tiers (Row 9)
    ws_sum.cell(8, 2, "Agent Inactivity & Dormancy Tiers").font = config.FONT_SECTION
    tier_headers = ["Inactivity Window", "Agents", "% of Fleet", "Audit Recommendation"]
    for c_idx, h in enumerate(tier_headers, 2):
        c = ws_sum.cell(9, c_idx, h)
        c.font = config.FONT_TBL_HEADER
        c.fill = config.FILL_NAVY_HEADER
        c.alignment = config.ALIGN_HEADER
        c.border = config.BORDER_THIN
    ws_sum.row_dimensions[9].height = 24

    tiers = [
        ("Active (<= 30 Days)", active_30, config.FONT_ALERT_GREEN, config.FILL_ALERT_GREEN, "Fully Protected & Checking In"),
        ("Inactive (31 - 90 Days)", inactive_90, config.FONT_ALERT_WARN, config.FILL_ALERT_WARN, "Monitor / Check Device Connectivity"),
        ("Critical Dormant (> 90 Days)", dormant, config.FONT_LAVENDER, config.FILL_LAVENDER, "Decommission or Reinstall Agent")
    ]
    cur_r = 10
    for name, cnt, fnt, fill, rec in tiers:
        ws_sum.cell(cur_r, 2, name).font = config.FONT_TBL_BOLD
        ws_sum.cell(cur_r, 2).fill = fill
        ws_sum.cell(cur_r, 2).border = config.BORDER_THIN

        c_cnt = ws_sum.cell(cur_r, 3, cnt)
        c_cnt.font = fnt
        c_cnt.fill = fill
        c_cnt.border = config.BORDER_THIN
        c_cnt.alignment = config.ALIGN_CENTER
        c_cnt.number_format = '#,##0'

        c_pct = ws_sum.cell(cur_r, 4, cnt / total_agents if total_agents else 0)
        c_pct.font = fnt
        c_pct.fill = fill
        c_pct.border = config.BORDER_THIN
        c_pct.alignment = config.ALIGN_CENTER
        c_pct.number_format = '0.0%'

        c_rec = ws_sum.cell(cur_r, 5, rec)
        c_rec.font = config.FONT_TBL_DATA
        c_rec.fill = fill
        c_rec.border = config.BORDER_THIN
        c_rec.alignment = config.ALIGN_LEFT
        cur_r += 1

    # Section 2: Operating System Breakdown (Placed at Column G)
    ws_sum.cell(8, 7, "Operating System Breakdown").font = config.FONT_SECTION
    os_headers = ["Operating System Group", "Count", "Share %"]
    for c_idx, h in enumerate(os_headers, 7):
        c = ws_sum.cell(9, c_idx, h)
        c.font = config.FONT_TBL_HEADER
        c.fill = config.FILL_SUB_HEADER
        c.alignment = config.ALIGN_HEADER
        c.border = config.BORDER_THIN

    os_counts = Counter(
        "Windows Server" if "Server" in a["OperatingSystem"] else
        "Windows Desktop" if "Windows" in a["OperatingSystem"] else
        "macOS" if any(k in a["OperatingSystem"] for k in ["Mac", "OS X", "Darwin"]) else "Other"
        for a in agents
    )

    r_os = 10
    for os_group, cnt in os_counts.most_common():
        ws_sum.cell(r_os, 7, os_group).font = config.FONT_TBL_DATA
        ws_sum.cell(r_os, 7).border = config.BORDER_THIN

        c_cnt = ws_sum.cell(r_os, 8, cnt)
        c_cnt.font = config.FONT_TBL_BOLD
        c_cnt.border = config.BORDER_THIN
        c_cnt.alignment = config.ALIGN_CENTER
        c_cnt.number_format = '#,##0'

        c_pct = ws_sum.cell(r_os, 9, cnt / total_agents if total_agents else 0)
        c_pct.font = config.FONT_TBL_DATA
        c_pct.border = config.BORDER_THIN
        c_pct.alignment = config.ALIGN_CENTER
        c_pct.number_format = '0.0%'
        r_os += 1

    # Add OS Doughnut Chart
    chart_os = DoughnutChart()
    chart_os.title = "Fleet OS Distribution"
    chart_os.width = 13
    chart_os.height = 7
    data_ref = Reference(ws_sum, min_col=8, min_row=9, max_row=r_os - 1)
    cats_ref = Reference(ws_sum, min_col=7, min_row=10, max_row=r_os - 1)
    chart_os.add_data(data_ref, titles_from_data=True)
    chart_os.set_categories(cats_ref)
    ws_sum.add_chart(chart_os, "G15")

    # Section 3: High-Priority Dormant Watchlist (>90 Days)
    dormant_agents = sorted([a for a in agents if a["DaysInactive"] is not None and a["DaysInactive"] > 90], key=lambda x: x["DaysInactive"], reverse=True)
    start_wl = max(cur_r + 2, 23)
    ws_sum.cell(start_wl, 2, f"Critical Dormant Watchlist (>90 Days Inactive) — {len(dormant_agents)} Devices").font = config.FONT_SECTION
    wl_headers = ["Agent Hostname", "Organization", "Operating System", "Last Seen", "Days Inactive", "Serial Number", "AD Domain"]
    for c_idx, h in enumerate(wl_headers, 2):
        c = ws_sum.cell(start_wl + 1, c_idx, h)
        c.font = config.FONT_TBL_HEADER
        c.fill = config.FILL_NAVY_HEADER
        c.alignment = config.ALIGN_HEADER
        c.border = config.BORDER_THIN
    ws_sum.row_dimensions[start_wl + 1].height = 24

    cur_wl_r = start_wl + 2
    for da in dormant_agents[:20]:
        vals = [da["Agent"], da["Organization"], da["OperatingSystem"], da["LastSeen"], da["DaysInactive"], da["SerialNumber"], da["ADDomain"]]
        for c_idx, val in enumerate(vals, 2):
            cell = ws_sum.cell(cur_wl_r, c_idx, val)
            cell.font = config.FONT_LAVENDER if c_idx == 6 else config.FONT_TBL_DATA
            cell.fill = config.FILL_LAVENDER
            cell.border = config.BORDER_THIN
            cell.alignment = config.ALIGN_LEFT if c_idx in [2, 3, 4] else config.ALIGN_CENTER
            if c_idx == 6:
                cell.number_format = '#,##0'
        cur_wl_r += 1

    if not dormant_agents:
        ws_sum.cell(cur_wl_r, 2, "No dormant agents detected (>90 days). All endpoints are compliant!").font = config.FONT_ALERT_GREEN
        ws_sum.cell(cur_wl_r, 2).fill = config.FILL_ALERT_GREEN

    # Column dimensions for Summary Tab
    sum_widths = {1: 4, 2: 26, 3: 16, 4: 16, 5: 36, 6: 4, 7: 24, 8: 14, 9: 14, 10: 16, 11: 16, 12: 18, 13: 18}
    for col_c, w in sum_widths.items():
        ws_sum.column_dimensions[get_column_letter(col_c)].width = w

    return ws_sum

def build_huntress_report(csv_path, output_xlsx=None, inv_title="Inventory", dash_title="Dashboard"):
    if not output_xlsx:
        out_dir = os.path.join(os.getcwd(), "reports", "Huntress Reports")
        os.makedirs(out_dir, exist_ok=True)
        today_str = datetime.now().strftime("%m%d%Y")
        output_xlsx = os.path.join(out_dir, f"Huntress_Report_{today_str}.xlsx")
    else:
        out_dir = os.path.dirname(output_xlsx)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

    agents = parse_huntress_csv(csv_path)
    wb = openpyxl.Workbook()
    ws_agents = build_huntress_inventory_tab(wb, agents, title=inv_title)
    build_huntress_dashboard_tab(wb, agents, title=dash_title)

    wb.active = ws_agents
    try:
        wb.save(output_xlsx)
        print(f"Huntress report successfully generated: {output_xlsx}")
    except PermissionError:
        raise PermissionError(
            f"Cannot save '{os.path.basename(output_xlsx)}' because the file is currently open in Microsoft Excel or another program.\n"
            "Please close the workbook in Excel and re-run the report."
        )
    return output_xlsx

if __name__ == "__main__":
    from detect_files import find_best_ninja_csvs

    csv_in = None
    if len(sys.argv) > 1:
        csv_in = sys.argv[1]
    else:
        detected = find_best_ninja_csvs()
        csv_in = detected.get("huntress_file")

    if not csv_in or not os.path.exists(csv_in):
        print("Error: No Huntress CSV file found.")
        sys.exit(1)

    out_file = sys.argv[2] if len(sys.argv) > 2 else None
    build_huntress_report(csv_in, out_file)
