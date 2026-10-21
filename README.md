# Humble Mumble

### Multi-Source Fleet Reporting & Google Tasks Orchestration Studio
*Unified IT Audit Workbooks, Cross-Reference Hygiene Auditing & Automated Remediation Workflows for **NinjaOne RMM**, **Huntress EDR**, and **Active Directory**.*

---

## Overview

**Humble Mumble** is a modular enterprise IT automation platform that eliminates the operational disconnect between RMM agents, EDR sensors, Active Directory computer objects, and day-to-day task remediation.

Instead of exporting separate spreadsheets that quickly fall out of sync, Humble Mumble reconciles all data sources into **executive-ready, multi-tab Excel workbooks** and automatically translates flagged endpoint hygiene issues into **Google Tasks** organized by office, site, or department.

---

## Key Capabilities

- **Unified Executive Workbooks**: Generates 7 to 12 structured worksheets with executive KPI cards, auto-filters, storage/reboot alert formatting, and distribution charts.
- **Cross-Source Reconciliation Engine**:
  - Surfaces devices active in Active Directory but missing from NinjaOne or Huntress.
  - Flags machines active in RMM whose AD accounts are disabled or orphaned in default containers (`CN=Computers`).
  - Detects "Ghost Assets" (devices dormant for >90 days or >365 days).
- **Direct Google Tasks Orchestration**:
  - Automatically translates flagged endpoints into Google Tasks lists (one list per office/department/site).
  - Creates parent device tasks (`Check Device: <HOSTNAME>`) with priority tags (`[P1]`, `[P2]`, `[P3]`).
  - Nests actionable subtasks (e.g., `Offline > 7 Days`, `Missing Huntress EDR`, `Drive Utilization >= 75%`).
  - Features smart deduplication, re-opening resolved tasks if an issue recurs, and dry-run preview.
- **Dual Desktop GUIs**:
  - **`app_gui.py`**: Full-featured report generation studio with 6 report modes, automatic CSV header detection, file age validation, and pre-flight lock checks.
  - **`google_tasks_gui.py`**: Interactive Google Tasks synchronization console with OAuth 2.0 sign-in and live progress telemetry.
- **Enterprise-Agnostic Architecture**: Fully customizable for any corporation, school district, healthcare provider, or MSP via a simple JSON configuration.

---

## Quick Start

### 1. Prerequisites & Installation
Ensure Python 3.10+ is installed on Windows, macOS, or Linux:

```bash
git clone https://github.com/your-username/humble-mumble.git
cd humble-mumble
pip install -r requirements.txt
```

### 2. Launch the Desktop Studio
Double-click **`Launch_Studio.bat`** (or execute from terminal):
```bash
python app_gui.py
```

*Note: Humble Mumble includes ready-to-run synthetic test files in the `samples/` directory. You can generate a full Extensive Report immediately without connecting to live systems!*

---

## Customization for Any Organization or MSP

Humble Mumble is designed to be completely organization-agnostic. You do not need to modify Python code to adapt it to your company or clients.

### Step 1: Create `organization_config.json`
Copy the included template:
```bash
copy organization_config.example.json organization_config.json
```

### Step 2: Define Your Settings
Configure your company domain, naming conventions, and hostname prefix routing:

```json
{
  "organization_name": "Contoso Enterprise",
  "domain_name": "CONTOSO.LOCAL",
  "default_location_name": "Headquarters",
  "hostname_prefixes": {
    "HQ": "Headquarters",
    "CORP": "Corporate Office",
    "BR1": "Branch Office 1",
    "BR2": "Branch Office 2",
    "REM": "Remote Workforce",
    "DC": "Datacenter Operations",
    "SRV": "Servers / Infrastructure",
    "IT": "IT Department"
  },
  "canonical_locations": {
    "HEADQUARTERS": "Headquarters",
    "HQ": "Headquarters",
    "CORPORATE": "Corporate Office",
    "BRANCH OFFICE 1": "Branch Office 1",
    "BRANCH 1": "Branch Office 1",
    "BRANCH OFFICE 2": "Branch Office 2",
    "BRANCH 2": "Branch Office 2",
    "REMOTE": "Remote Workforce",
    "SERVERS / INFRASTRUCTURE": "Servers / Infrastructure",
    "DEFAULT CONTAINER (UNASSIGNED)": "Default Container (Unassigned)",
    "DISABLED COMPUTERS": "Disabled Computers",
    "UNASSIGNED": "Unassigned"
  }
}
```

*If no `organization_config.json` is supplied, Humble Mumble uses clean built-in presets and automatically applies Title Case formatting to all unmapped Active Directory OUs and RMM locations.*

---

## Report Modes & Tab Composition

| Report Mode | Target Audience | Tab Count | Included Sources |
| :--- | :--- | :-: | :--- |
| **Extensive Report** | IT Directors, CISOs & MSP Leadership | **12 Tabs** | NinjaOne + AD Install Dates + Huntress + AD Raw |
| **Ninja + Active Directory** | Systems Engineers & Domain Admins | **9 Tabs** | NinjaOne + AD Install Dates + AD Raw |
| **Ninja + Huntress** | Security Operations (SOC) & SecOps | **9 Tabs** | NinjaOne + Huntress EDR (No AD) |
| **Ninja Only** | Asset Managers | **6 Tabs** | NinjaOne Fleet Only |
| **Active Directory Standalone** | Active Directory Compliance Audits | **2 Tabs** | Active Directory Raw Export Only |
| **Huntress Standalone** | Endpoint Security Audits | **2 Tabs** | Huntress Agents Export Only |

### Flagship Extensive Report (12-Tab Structure)
1. **Ninja Inventory**: Master hardware table with RAM, CPU speeds, drive utilization, warranty days, and uptime tiers.
2. **AD Inventory**: Complete domain computer inventory with OU token splits, operating system versions, and logon ages.
3. **Huntress Inventory**: Master EDR agent table with client versions, IP addresses, and inactivity metrics.
4. **Patch & Reboot Health**: OS build compliance, reboot age tiers (7d, 14d, 30d, 90d), and patch latency.
5. **Lifecycle**: Multi-year hardware refresh replacement waves with budget forecasting.
6. **Performance & Capacity**: Surfaces low-memory devices (<16GB) and critical drive space constraints (>=75%).
7. **Site Allocation**: Workstation vs. laptop mobility ratios across physical branches and departments.
8. **Stale Assets**: Inactive endpoints offline >90 days and >365 days with recommended disposition actions.
9. **AD Dashboard**: Executive summary of domain hygiene, active vs. disabled assets, and unassigned container warnings.
10. **AD Cross ref**: Cross-fleet audit reconciling NinjaOne devices with Active Directory objects.
11. **Huntress Dashboard**: EDR deployment breakdown, dormant watchlist, and OS distribution doughnut chart.
12. **Huntress Cross ref**: Protection gap analysis matching NinjaOne fleet against active Huntress agents.

---

## Google Tasks Synchronization Setup

Humble Mumble can automatically sync actionable endpoint remediation tasks directly to **Google Tasks**.

### Step 1: Create a Google Cloud Project
1. Navigate to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project (e.g., `Humble-Mumble-Fleet-Tasks`).
3. In **APIs & Services > Library**, search for and enable the **Google Tasks API**.

### Step 2: Configure the OAuth Consent Screen
1. Go to **APIs & Services > OAuth consent screen**.
2. Select **External** (or **Internal** if using Google Workspace).
3. Enter an App Name (e.g., `Humble Mumble Fleet Sync`) and developer contact email.
4. Under **Scopes**, add `https://www.googleapis.com/auth/tasks`.
5. Under **Test users**, add your Google email address.

### Step 3: Create OAuth 2.0 Credentials
1. Go to **APIs & Services > Credentials**.
2. Click **Create Credentials > OAuth client ID**.
3. Select Application Type: **Desktop app**.
4. Name: `Humble Mumble Desktop Client`.
5. Click **Download JSON** and save the file in your project root as:
   ```
   credentials.json
   ```
   *(A sample template is provided at `credentials.json.example`)*

### Step 4: Run Synchronization
Launch the Google Tasks Studio:
```bash
python google_tasks_gui.py
```
1. Click **Sign In to Google** to authorize through your web browser. A local token (`token.json`) is securely cached and protected by `.gitignore`.
2. Click **Test Run / Preview** to simulate task creation without modifying Google Tasks.
3. Click **Push to Google Tasks** to sync live.

---

## Repository Structure & Modular Architecture

Humble Mumble is organized as an enterprise Python package structured by environment under `src/humble_mumble/`:

```
Humble Mumble/
├── pyproject.toml                  # Standard Python packaging (PEP 517/621) with console CLI commands
├── requirements.txt                # Python package dependencies
├── LICENSE                         # MIT Open Source License
├── README.md                       # Comprehensive documentation & quick start
├── organization_config.example.json# Template for custom company prefixes and locations
├── credentials.json.example        # Template for Google Tasks OAuth 2.0 Desktop credentials
│
├── Launch_Studio.bat               # 1-click launcher for the Executive Desktop Studio
├── Launch_Google_Tasks_Sync.bat    # 1-click launcher for Google Tasks Fleet Sync Studio
├── run_pipeline.bat                # 1-click runner for NinjaOne inventory pipeline
├── run_ad_pipeline.bat             # 1-click runner for Active Directory export & report pipeline
├── app_gui.py                      # Root launcher script for the Desktop Studio
├── google_tasks_gui.py             # Root launcher script for Google Tasks sync studio
├── Export_AD_Computers_Raw.ps1     # Standalone Active Directory PowerShell discovery script
│
├── src/
│   └── humble_mumble/              # Core enterprise Python package
│       ├── __init__.py             # Package root & metadata
│       ├── __main__.py             # Supports `python -m humble_mumble`
│       │
│       ├── core/                   # Shared Foundation & Infrastructure
│       │   ├── config.py           # Design tokens, color palette, column definitions & thresholds
│       │   ├── utils.py            # Safe parsing, KPI card builders, date & serial helpers
│       │   └── detect_files.py     # Header sniffing & timestamp auto-discovery for CSV exports
│       │
│       ├── ninjaone/               # NinjaOne RMM Environment
│       │   ├── build_inventory.py  # Core NinjaOne & Extensive 12-tab workbook generator
│       │   └── generate_dashboard_tabs.py # Visual KPIs, donut charts & summary dashboards
│       │
│       ├── ad/                     # Active Directory Domain Environment
│       │   ├── build_ad_report.py  # Standalone Active Directory inventory & hygiene report builder
│       │   └── Export_AD_Computers_Raw.ps1 # RSAT Active Directory query & CSV export script
│       │
│       ├── huntress/               # Huntress EDR Security Environment
│       │   └── build_huntress_report.py # Standalone Huntress EDR inventory & audit report builder
│       │
│       ├── google_tasks/           # Google Tasks Remediation Environment
│       │   ├── task_exporter.py    # Deduplicated flagged devices & Google Tasks CSV exporter
│       │   ├── google_tasks_sync.py# OAuth 2.0 sync engine with exponential backoff & rate pacing
│       │   └── google_tasks_gui.py # Interactive desktop UI for Google Tasks sync & preview
│       │
│       └── gui/                    # Desktop Presentation Layer
│           └── app_gui.py          # Unified multi-mode reporting studio (Tkinter)
│
├── samples/                        # Synthetic mock CSV files for instant offline testing
│   ├── sample_ninja_device_details.csv
│   ├── sample_ninja_custom_fields.csv
│   ├── sample_huntress_agents.csv
│   └── sample_ad_computers_raw.csv
│
└── reports/                        # Destination for generated Excel workbooks & tasks
    ├── AD Reports/
    │   └── raw report/             # Automated destination for Computers_Raw_*.csv exports
    ├── Extensive Reports/
    ├── Huntress Reports/
    ├── Ninja Reports/
    └── Tasks/
```

### CLI Commands & Module Execution

In addition to the batch launchers and desktop GUIs, you can execute commands directly:

```bash
# Install in editable mode (optional, for global CLI commands)
pip install -e .

# Launch Desktop Reporting Studio
humble-mumble
# or: python -m humble_mumble
# or: python app_gui.py

# Launch Google Tasks Studio
humble-tasks
# or: python -m humble_mumble.google_tasks.google_tasks_gui

# Run NinjaOne Inventory Pipeline directly
humble-ninja
# or: python -m humble_mumble.ninjaone.build_inventory

# Run Active Directory Report directly
humble-ad
# or: python -m humble_mumble.ad.build_ad_report

# Run Huntress EDR Report directly
humble-huntress
# or: python -m humble_mumble.huntress.build_huntress_report
```

---

## Security & Data Hygiene

- **No Proprietary Data**: The repository includes zero customer data, zero proprietary credentials, and zero internal corporate network identifiers.
- **Robust `.gitignore`**: All raw client exports (`*.csv`), completed client reports (`*.xlsx`), credentials (`credentials.json`), and OAuth tokens (`token.json`) are strictly excluded from version control.
- **Rate-Limiting & Defensive Retries**: The Google Tasks integration implements exponential backoff on HTTP 429/500/503 errors and introduces pacing delays to prevent API quota starvation.
- **Pre-Flight Lock Prevention**: The studio checks whether destination Excel files are currently locked open by Microsoft Excel before starting threads to prevent write collision crashes.

---

## License

This project is licensed under the [MIT License](LICENSE).
