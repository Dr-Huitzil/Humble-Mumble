"""
config.py
=========
Central configuration for Enterprise Fleet Reporting & Google Tasks Orchestration Studio.
All organization parameters, thresholds, design tokens, typography, borders, fills,
and column definitions are centralized here.
"""

import os
import json
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# -----------------------------------------------------------------------------
# 1. ORGANIZATION & DOMAIN CONFIGURATION
# -----------------------------------------------------------------------------
# Defaults provide an enterprise template. These can be overridden dynamically
# by placing an `organization_config.json` (or `organization_config.local.json`)
# in the project root.
ORGANIZATION_NAME = "Enterprise Fleet Management"
DOMAIN_NAME = "CORPORATE.LOCAL"
DEFAULT_LOCATION = "Headquarters"

# Default generic prefix mapping for computer hostnames (e.g., HQ-D-101 -> Headquarters)
CLINIC_PREFIX_MAP = {
    'HQ': 'Headquarters',
    'CORP': 'Corporate Office',
    'BR1': 'Branch Office 1',
    'BR2': 'Branch Office 2',
    'REM': 'Remote Workforce',
    'DC': 'Datacenter',
    'SRV': 'Servers & Infrastructure',
    'IT': 'IT Department',
    'LAB': 'R&D Lab'
}

# Canonical location mapping to normalize Active Directory OUs and RMM site names
CLINIC_NAME_CANONICAL = {
    "HEADQUARTERS": "Headquarters",
    "HQ": "Headquarters",
    "CORPORATE": "Corporate Office",
    "CORP": "Corporate Office",
    "BRANCH OFFICE 1": "Branch Office 1",
    "BRANCH 1": "Branch Office 1",
    "BRANCH OFFICE 2": "Branch Office 2",
    "BRANCH 2": "Branch Office 2",
    "REMOTE": "Remote Workforce",
    "DATACENTER": "Datacenter Operations",
    "DC": "Datacenter Operations",
    "SERVERS / INFRASTRUCTURE": "Servers / Infrastructure",
    "SERVERS": "Servers / Infrastructure",
    "DEFAULT CONTAINER (UNASSIGNED)": "Default Container (Unassigned)",
    "DISABLED COMPUTERS": "Disabled Computers",
    "UNASSIGNED": "Unassigned"
}

# Auto-load custom organization configuration if present
_cfg_file = None
for _candidate in ["organization_config.local.json", "organization_config.json"]:
    if os.path.exists(_candidate):
        _cfg_file = _candidate
        break

if _cfg_file:
    try:
        with open(_cfg_file, "r", encoding="utf-8") as _f:
            _loaded = json.load(_f)
            ORGANIZATION_NAME = _loaded.get("organization_name", ORGANIZATION_NAME)
            DOMAIN_NAME = _loaded.get("domain_name", DOMAIN_NAME)
            DEFAULT_LOCATION = _loaded.get("default_location_name", DEFAULT_LOCATION)
            if "hostname_prefixes" in _loaded and isinstance(_loaded["hostname_prefixes"], dict):
                CLINIC_PREFIX_MAP.update(_loaded["hostname_prefixes"])
            if "canonical_locations" in _loaded and isinstance(_loaded["canonical_locations"], dict):
                for k, v in _loaded["canonical_locations"].items():
                    CLINIC_NAME_CANONICAL[k.upper()] = v
    except Exception as _e:
        print(f"[Notice] Could not parse {_cfg_file}: {_e}")

def canonicalize_clinic_name(name):
    """
    Normalizes a clinic or location string across Active Directory, NinjaOne,
    and Huntress to ensure clean grouping and single Google Tasks lists per location.
    """
    if not name:
        return "Unassigned"
    s = str(name).strip()
    if not s or s.lower() in ["unassigned", "none", "n/a", "unknown"]:
        return "Unassigned"

    upper_s = s.upper()
    if upper_s in CLINIC_NAME_CANONICAL:
        return CLINIC_NAME_CANONICAL[upper_s]

    # Check without common location suffixes
    for suffix in [" CLINIC", " OFFICE", " LOCATION", " BRANCH", " SITE"]:
        if upper_s.endswith(suffix):
            base = upper_s[:-len(suffix)].strip()
            if base in CLINIC_NAME_CANONICAL:
                return CLINIC_NAME_CANONICAL[base]
            if f"{base} CLINIC" in CLINIC_NAME_CANONICAL:
                return CLINIC_NAME_CANONICAL[f"{base} CLINIC"]

    # If all-caps or generic string, return clean Title Case
    return s.title() if s.isupper() else s

# -----------------------------------------------------------------------------
# 2. DATES & REPORT METADATA
# -----------------------------------------------------------------------------
CURRENT_MONTH_YEAR = datetime.now().strftime("%B %Y")

REPORT_METADATA = {
    "patch_reboot": {
        "title": "ENDPOINT SECURITY, PATCHING & REBOOT",
        "subtitle": f"Tracking OS Build Compliance, Reboot Frequency & Patch Latency | {CURRENT_MONTH_YEAR}"
    },
    "lifecycle": {
        "title": "HARDWARE LIFECYCLE & REFRESH",
        "subtitle": f"Replacement Waves & Lifecycle Forecasting (3-Year Policy) | {CURRENT_MONTH_YEAR}"
    },
    "performance_capacity": {
        "title": "STORAGE CAPACITY & MEMORY",
        "subtitle": f"Surfacing Disk Space Shortages (>=75%) and Sub-16GB RAM Workstations | {CURRENT_MONTH_YEAR}"
    },
    "site_allocation": {
        "title": "SITE ALLOCATION & MOBILITY AUDIT",
        "subtitle": f"Hardware Footprint, Form Factor Breakdown & Mobility Ratios Across All Sites | {CURRENT_MONTH_YEAR}"
    },
    "stale_assets": {
        "title": "INACTIVE ASSET AUDIT",
        "subtitle": f"Tracking Endpoints Offline >90 Days (Agent Removal Window) & >1 Year Decommission Targets | {CURRENT_MONTH_YEAR}"
    },
    "huntress_protection": {
        "title": "HUNTRESS ENDPOINT PROTECTION & COMPLIANCE",
        "subtitle": f"Cross-Referencing Fleet vs Huntress EDR Agents | Identifying Protection Gaps & Inactivity Drops | {CURRENT_MONTH_YEAR}"
    }
}

# -----------------------------------------------------------------------------
# 3. INVENTORY COLUMN HEADERS (26 Columns)
# -----------------------------------------------------------------------------
INVENTORY_HEADERS = [
    "Device", "Location", "Device Role", "Device Model", "Serial Number",
    "Last Login", "Last Uptime_formatted", "Last Uptime - Days",
    "Last Boot Time_formatted", "Last Boot Time - Days", "Policy", "OS",
    "Last Update_formatted", "Last Update - Days", "RAM",
    "Primary Drive Capacity (GB)", "Primary Drive Free (GB)", "Primary Drive Usage %",
    "Processors Name", "Processors Max Clock Speed", "AD Install Date",
    "Refresh Year", "Warranty Start Date_formatted", "Warranty End Date_formatted",
    "Warranty Status", "Warranty Days Remaining"
]

INVENTORY_COL_WIDTHS = {
    1: 18, 2: 24, 3: 16, 4: 24, 5: 14, 6: 22, 7: 20, 8: 14, 9: 20, 10: 14,
    11: 20, 12: 10, 13: 20, 14: 14, 15: 10, 16: 18, 17: 16, 18: 16,
    19: 34, 20: 18, 21: 14, 22: 12, 23: 20, 24: 20, 25: 14, 26: 16
}

# -----------------------------------------------------------------------------
# 4. HUNTRESS AGENT HEADERS & WIDTHS
# -----------------------------------------------------------------------------
HUNTRESS_HEADERS = [
    "Agent", "Organization", "Operating System", "Agent Version", "EDR Version",
    "MacOS EDR enabled", "Internal IP", "External IP", "Registered", "Last Seen",
    "Serial Number", "AD Domain", "Ninja Status", "Days Inactive"
]

HUNTRESS_COL_WIDTHS = {
    1: 18, 2: 14, 3: 28, 4: 16, 5: 14, 6: 18, 7: 18, 8: 18, 9: 22, 10: 22,
    11: 24, 12: 16, 13: 20, 14: 16
}

HUNTRESS_STANDALONE_HEADERS = [
    "Agent", "Organization", "Operating System", "Agent Version", "EDR Version",
    "MacOS EDR enabled", "Internal IP", "External IP", "Registered", "Last Seen",
    "Serial Number", "AD Domain", "Days Inactive"
]

HUNTRESS_STANDALONE_COL_WIDTHS = {
    1: 18, 2: 18, 3: 28, 4: 16, 5: 14, 6: 18, 7: 18, 8: 18, 9: 22, 10: 22,
    11: 24, 12: 18, 13: 16
}

# -----------------------------------------------------------------------------
# 5. ACTIVE DIRECTORY COMPUTER HEADERS & WIDTHS
# -----------------------------------------------------------------------------
AD_HEADERS = [
    "Name", "Status", "Clinic", "Department", "Device Type",
    "Operating System", "OS Version", "Last Logon Date", "Last Logon (Days)",
    "Created Date", "Created (Days)", "Modified Date", "Modified (Days)",
    "OU Token 1", "OU Token 2", "OU Token 3"
]

AD_COL_WIDTHS = {
    1: 20, 2: 12, 3: 24, 4: 16, 5: 14, 6: 28, 7: 16, 8: 20, 9: 16,
    10: 20, 11: 14, 12: 20, 13: 14, 14: 24, 15: 22, 16: 22
}

# -----------------------------------------------------------------------------
# 6. AUDIT & THRESHOLD RULES
# -----------------------------------------------------------------------------
THRESHOLDS = {
    "reboot_compliant_days": 7,
    "reboot_warning_days": 14,
    "reboot_risk_days": 30,
    "reboot_neglect_days": 90,
    "storage_warning_pct": 0.65,
    "storage_alert_pct": 0.75,
    "storage_critical_pct": 0.85,
    "stale_threshold_days": 60,
    "huntress_recent_days": 30,
    "huntress_drop_risk_days": 90,
    "decommission_days": 365,
    "warranty_expiring_soon_days": 90,
    "refresh_cycle_years": 3,
    "auto_detect_max_age_days": 7
}

# -----------------------------------------------------------------------------
# 7. EXECUTIVE DESIGN PALETTE (Hex Codes)
# -----------------------------------------------------------------------------
COLOR_PRIMARY_NAVY = "1F4E79"
COLOR_NAVY         = "1F4E79"
COLOR_SUB_HEADER    = "2F5597"
COLOR_ACCENT_BLUE   = "41719C"
COLOR_ACCENT        = "41719C"
COLOR_KPI_BG        = "F2F5F9"
COLOR_ZEBRA         = "F9FAFB"
COLOR_WHITE         = "FFFFFF"
COLOR_TOTAL_ROW     = "E9EEF4"
COLOR_TOTAL         = "E9EEF4"
COLOR_BORDER        = "E5E7EB"
COLOR_BORDER_ALT    = "D9D9D9"

# Alert colors (Muted, professional pastels)
COLOR_ALERT_YELLOW_FILL   = "FFF2CC"
COLOR_ALERT_YELLOW_FONT   = "B25900"
COLOR_ALERT_RED_FILL      = "FCE4D6"
COLOR_ALERT_RED_FONT      = "C00000"
COLOR_ALERT_DARKRED_FILL  = "F8CBAD"
COLOR_ALERT_DARKRED_FONT  = "78281F"
COLOR_ALERT_GREEN_FILL    = "E2EFDA"
COLOR_ALERT_GREEN_FONT    = "385723"
COLOR_ALERT_LAVENDER_FILL = "E1D5E7"
COLOR_ALERT_LAVENDER_FONT = "5B2C6F"
COLOR_STAGING_FILL        = "DDEBF7"
COLOR_STAGING_FONT        = "1F4E79"
COLOR_PEACH_FILL          = "FCE4D6"
COLOR_PEACH_FONT          = "C65911"
COLOR_SALMON_FILL         = "F8CBAD"
COLOR_SALMON_FONT         = "C00000"
COLOR_LIGHTRED_FILL       = "F4B084"
COLOR_LIGHTRED_FONT       = "8A0000"

# -----------------------------------------------------------------------------
# 8. REUSABLE TYPOGRAPHY & FONTS
# -----------------------------------------------------------------------------
FONT_TITLE       = Font(name="Calibri", size=16, bold=True, color=COLOR_PRIMARY_NAVY)
FONT_SUBTITLE    = Font(name="Calibri", size=10, italic=True, color="595959")
FONT_SECTION     = Font(name="Calibri", size=12, bold=True, color=COLOR_PRIMARY_NAVY)
FONT_CARD_NUM    = Font(name="Calibri", size=18, bold=True, color=COLOR_PRIMARY_NAVY)
FONT_CARD_LABEL  = Font(name="Calibri", size=9, bold=True, color="595959")
FONT_CARD_SUB    = Font(name="Calibri", size=8, italic=True, color="7F7F7F")
FONT_TBL_HEADER  = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_TBL_DATA    = Font(name="Calibri", size=10, color="1F2937")
FONT_TBL_BOLD    = Font(name="Calibri", size=10, bold=True, color="1F2937")
FONT_TBL_TOTAL   = Font(name="Calibri", size=10, bold=True, color="000000")

# Alert / Status Fonts
FONT_ALERT_RED      = Font(name="Calibri", size=10, bold=True, color=COLOR_ALERT_RED_FONT)
FONT_ALERT_WARN     = Font(name="Calibri", size=10, bold=True, color=COLOR_ALERT_YELLOW_FONT)
FONT_ALERT_GREEN    = Font(name="Calibri", size=10, bold=True, color=COLOR_ALERT_GREEN_FONT)
FONT_ALERT_LAVENDER = Font(name="Calibri", size=10, bold=True, color=COLOR_ALERT_LAVENDER_FONT)
FONT_LAVENDER       = FONT_ALERT_LAVENDER
FONT_STAGING        = Font(name="Calibri", size=10, bold=True, color=COLOR_STAGING_FONT)

FONT_WEEK    = Font(name="Calibri", size=10, bold=True, color=COLOR_ALERT_YELLOW_FONT)
FONT_MONTH   = Font(name="Calibri", size=10, bold=True, color=COLOR_PEACH_FONT)
FONT_3MONTHS = Font(name="Calibri", size=10, bold=True, color=COLOR_SALMON_FONT)
FONT_6MONTHS = Font(name="Calibri", size=10, bold=True, color=COLOR_LIGHTRED_FONT)
FONT_ACTIVE  = Font(name="Calibri", size=10, bold=True, color=COLOR_ALERT_GREEN_FONT)

# -----------------------------------------------------------------------------
# 9. REUSABLE PATTERN FILLS
# -----------------------------------------------------------------------------
FILL_NAVY_HEADER   = PatternFill(start_color=COLOR_PRIMARY_NAVY, end_color=COLOR_PRIMARY_NAVY, fill_type="solid")
FILL_SUB_HEADER    = PatternFill(start_color=COLOR_SUB_HEADER, end_color=COLOR_SUB_HEADER, fill_type="solid")
FILL_ACCENT_HEADER = PatternFill(start_color=COLOR_ACCENT_BLUE, end_color=COLOR_ACCENT_BLUE, fill_type="solid")
FILL_KPI_CARD      = PatternFill(start_color=COLOR_KPI_BG, end_color=COLOR_KPI_BG, fill_type="solid")
FILL_ZEBRA         = PatternFill(start_color=COLOR_ZEBRA, end_color=COLOR_ZEBRA, fill_type="solid")
FILL_ZEBRA_ROW     = FILL_ZEBRA
FILL_WHITE         = PatternFill(start_color=COLOR_WHITE, end_color=COLOR_WHITE, fill_type="solid")
FILL_WHITE_ROW     = FILL_WHITE
FILL_TOTAL         = PatternFill(start_color=COLOR_TOTAL_ROW, end_color=COLOR_TOTAL_ROW, fill_type="solid")
FILL_TOTAL_ROW     = FILL_TOTAL

# Alert / Status Fills
FILL_ALERT_RED      = PatternFill(start_color=COLOR_ALERT_RED_FILL, end_color=COLOR_ALERT_RED_FILL, fill_type="solid")
FILL_ALERT_WARN     = PatternFill(start_color=COLOR_ALERT_YELLOW_FILL, end_color=COLOR_ALERT_YELLOW_FILL, fill_type="solid")
FILL_ALERT_GREEN    = PatternFill(start_color=COLOR_ALERT_GREEN_FILL, end_color=COLOR_ALERT_GREEN_FILL, fill_type="solid")
FILL_ALERT_LAVENDER = PatternFill(start_color=COLOR_ALERT_LAVENDER_FILL, end_color=COLOR_ALERT_LAVENDER_FILL, fill_type="solid")
FILL_LAVENDER       = FILL_ALERT_LAVENDER
FILL_STAGING        = PatternFill(start_color=COLOR_STAGING_FILL, end_color=COLOR_STAGING_FILL, fill_type="solid")

FILL_WEEK    = FILL_ALERT_WARN
FILL_MONTH   = PatternFill(start_color=COLOR_PEACH_FILL, end_color=COLOR_PEACH_FILL, fill_type="solid")
FILL_3MONTHS = PatternFill(start_color=COLOR_SALMON_FILL, end_color=COLOR_SALMON_FILL, fill_type="solid")
FILL_6MONTHS = PatternFill(start_color=COLOR_LIGHTRED_FILL, end_color=COLOR_LIGHTRED_FILL, fill_type="solid")
FILL_ACTIVE  = FILL_ALERT_GREEN

# -----------------------------------------------------------------------------
# 10. BORDERS
# -----------------------------------------------------------------------------
BORDER_THIN = Border(
    left=Side(style='thin', color=COLOR_BORDER),
    right=Side(style='thin', color=COLOR_BORDER),
    top=Side(style='thin', color=COLOR_BORDER),
    bottom=Side(style='thin', color=COLOR_BORDER)
)
BORDER_TOTAL = Border(
    left=Side(style='thin', color=COLOR_BORDER),
    right=Side(style='thin', color=COLOR_BORDER),
    top=Side(style='thin', color=COLOR_PRIMARY_NAVY),
    bottom=Side(style='double', color=COLOR_PRIMARY_NAVY)
)
BORDER_KPI = Border(
    left=Side(style='medium', color=COLOR_PRIMARY_NAVY),
    right=Side(style='thin', color=COLOR_BORDER),
    top=Side(style='thin', color=COLOR_BORDER),
    bottom=Side(style='thin', color=COLOR_BORDER)
)

# -----------------------------------------------------------------------------
# 11. ALIGNMENTS
# -----------------------------------------------------------------------------
ALIGN_LEFT   = Alignment(horizontal="left", vertical="center")
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_RIGHT  = Alignment(horizontal="right", vertical="center")
ALIGN_HEADER = Alignment(horizontal="center", vertical="center", wrap_text=True)

# -----------------------------------------------------------------------------
# 12. LOWERCASE ALIASES (Full Backward Compatibility)
# -----------------------------------------------------------------------------
font_title       = FONT_TITLE
font_subtitle    = FONT_SUBTITLE
font_section     = FONT_SECTION
font_card_num    = FONT_CARD_NUM
font_card_label  = FONT_CARD_LABEL
font_card_sub    = FONT_CARD_SUB
font_tbl_header  = FONT_TBL_HEADER
font_tbl_data    = FONT_TBL_DATA
font_tbl_bold    = FONT_TBL_BOLD
font_tbl_total   = FONT_TBL_TOTAL

font_red         = FONT_ALERT_RED
font_yellow      = FONT_ALERT_WARN
font_green       = FONT_ALERT_GREEN
font_purple      = FONT_ALERT_LAVENDER
font_lavender    = FONT_ALERT_LAVENDER
font_staging     = FONT_STAGING
font_week        = FONT_WEEK
font_month       = FONT_MONTH
font_3months     = FONT_3MONTHS
font_6months     = FONT_6MONTHS
font_active      = FONT_ACTIVE

fill_navy_header   = FILL_NAVY_HEADER
fill_sub_header    = FILL_SUB_HEADER
fill_accent_header = FILL_ACCENT_HEADER
fill_kpi_card      = FILL_KPI_CARD
fill_zebra         = FILL_ZEBRA
fill_zebra_row     = FILL_ZEBRA_ROW
fill_white         = FILL_WHITE
fill_white_row     = FILL_WHITE_ROW
fill_total         = FILL_TOTAL
fill_total_row     = FILL_TOTAL_ROW

fill_red         = FILL_ALERT_RED
fill_yellow      = FILL_ALERT_WARN
fill_green       = FILL_ALERT_GREEN
fill_purple      = FILL_ALERT_LAVENDER
fill_lavender    = FILL_ALERT_LAVENDER
fill_staging     = FILL_STAGING
fill_week        = FILL_WEEK
fill_month       = FILL_MONTH
fill_3months     = FILL_3MONTHS
fill_6months     = FILL_6MONTHS
fill_active      = FILL_ACTIVE

thin_border      = BORDER_THIN
total_border     = BORDER_TOTAL
kpi_border       = BORDER_KPI

align_left       = ALIGN_LEFT
align_center     = ALIGN_CENTER
align_right      = ALIGN_RIGHT
align_header     = ALIGN_HEADER
