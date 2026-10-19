"""
app_gui.py
==========
Interactive Desktop UI for NinjaOne, Huntress & Active Directory Reporting Studio.

Features:
  - Multi-Mode Report Customization:
      * Extensive Report (Ninja + AD Install Date + Huntress)
      * Ninja + AD Install Date
      * Ninja + Huntress (No AD Lifecycle)
      * Ninja Only (No AD / No Huntress)
      * Active Directory Report Only (with in-app PowerShell CSV generator)
      * Huntress Report Only (Standalone)
  - Automatic Directory Routing: Reports are routed directly to the designated
    subfolders inside reports/ (Extensive Reports, Ninja Reports, Huntress Reports, AD Reports).
  - Smart File Header Auto-Detection & Real-Time Validation Badges.
  - Non-blocking Threading for execution and live PowerShell script generation.
  - One-Click Excel Launch to open generated reports directly.
"""

import os
import sys
import threading
import subprocess
import time
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from humble_mumble.core.detect_files import classify_csv, find_best_ninja_csvs, find_best_ad_csv
from humble_mumble.core import config
from humble_mumble.ninjaone.build_inventory import build_inventory_sheet
from humble_mumble.ad.build_ad_report import build_report as build_ad_report_sheet
from humble_mumble.huntress.build_huntress_report import build_huntress_report

# Report Mode Definitions
MODES = [
    {
        "id": "extensive",
        "name": "Extensive Report (Ninja + AD + Huntress)",
        "subfolder": "Extensive Reports",
        "file_prefix": "Extensive_Inventory_Report",
        "desc": "Complete enterprise audit combining Ninja hardware, AD 3-Year Lifecycle & Huntress protection."
    },
    {
        "id": "ninja_ad",
        "name": "Ninja + Active Directory",
        "subfolder": "Ninja Reports",
        "file_prefix": "Ninja_Inventory_AD",
        "desc": "Ninja fleet + AD Lifecycle, AD Inventory, AD Dashboard, and AD Cross-Reference audit."
    },
    {
        "id": "ninja_huntress",
        "name": "Ninja + Huntress",
        "subfolder": "Ninja Reports",
        "file_prefix": "Ninja_Inventory_Huntress",
        "desc": "Ninja fleet + Huntress Inventory, Huntress Dashboard, and Huntress Cross-Reference audit."
    },
    {
        "id": "ninja_only",
        "name": "Ninja Only (No AD / No Huntress)",
        "subfolder": "Ninja Reports",
        "file_prefix": "Ninja_Inventory",
        "desc": "Standalone NinjaOne inventory + 4 health dashboards. Lifecycle tab omitted."
    },
    {
        "id": "ad_only",
        "name": "Active Directory Report Only",
        "subfolder": "AD Reports",
        "file_prefix": "AD_Computer_Inventory_Report",
        "desc": "Standalone 2-tab Active Directory computer inventory & executive audit dashboard."
    },
    {
        "id": "huntress_only",
        "name": "Huntress Report Only",
        "subfolder": "Huntress Reports",
        "file_prefix": "Huntress_Report",
        "desc": "Standalone 2-tab Huntress agent inventory & executive protection summary dashboard."
    }
]

class ReportingStudioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Executive Reporting Studio — NinjaOne, Huntress & Active Directory")
        self.root.geometry("820x820")
        self.root.minsize(760, 720)
        self.root.configure(bg="#F4F6F9")

        # Palette
        self.c_navy = "#1F4E79"
        self.c_slate = "#2F5597"
        self.c_bg = "#F4F6F9"
        self.c_card = "#FFFFFF"
        self.c_border = "#D9E2EC"
        self.c_green = "#2E7D32"
        self.c_red = "#C62828"
        self.c_warn = "#B25900"
        self.c_gray = "#6B7280"
        self.c_text = "#1F2937"

        # State Variables
        self.selected_mode_idx = tk.IntVar(value=0)
        self.mode_desc_var = tk.StringVar()

        self.details_path = tk.StringVar()
        self.custom_path = tk.StringVar()
        self.huntress_path = tk.StringVar()
        self.ad_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self.details_status = tk.StringVar(value="Waiting for file...")
        self.custom_status = tk.StringVar(value="Waiting for file...")
        self.huntress_status = tk.StringVar(value="Waiting for file...")
        self.ad_status = tk.StringVar(value="Waiting for file...")

        self.is_custom_output = False
        self.is_running_ad_export = False

        self.setup_ui()
        self.on_mode_change()
        self.auto_detect_files(notify=False)

    def get_current_mode(self):
        idx = self.selected_mode_idx.get()
        if 0 <= idx < len(MODES):
            return MODES[idx]
        return MODES[0]

    def setup_ui(self):
        # -------------------------------------------------------------
        # 1. Header Banner
        # -------------------------------------------------------------
        header_frame = tk.Frame(self.root, bg=self.c_navy, height=76)
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)

        title_lbl = tk.Label(
            header_frame,
            text="Executive Reporting Studio",
            font=("Segoe UI", 16, "bold"),
            fg="#FFFFFF",
            bg=self.c_navy
        )
        title_lbl.pack(anchor="w", padx=20, pady=(12, 2))

        sub_lbl = tk.Label(
            header_frame,
            text="NinjaOne RMM  •  Huntress EDR  •  Active Directory Inventory & Compliance Pipelines",
            font=("Segoe UI", 9),
            fg="#D9E2EC",
            bg=self.c_navy
        )
        sub_lbl.pack(anchor="w", padx=20)

        # Main Scrollable / Outer Container
        main_frame = tk.Frame(self.root, bg=self.c_bg, padx=18, pady=10)
        main_frame.pack(fill="both", expand=True)

        # -------------------------------------------------------------
        # 2. Card 0: Report Mode Selector
        # -------------------------------------------------------------
        card_mode = tk.LabelFrame(
            main_frame,
            text=" 1. Report Mode Selection ",
            font=("Segoe UI", 10, "bold"),
            fg=self.c_navy,
            bg=self.c_card,
            padx=14,
            pady=8,
            relief="solid",
            bd=1
        )
        card_mode.pack(fill="x", pady=(0, 8))

        mode_names = [m["name"] for m in MODES]
        self.combo_mode = ttk.Combobox(
            card_mode,
            values=mode_names,
            state="readonly",
            font=("Segoe UI", 9, "bold"),
            width=55
        )
        self.combo_mode.current(0)
        self.combo_mode.pack(anchor="w", pady=(2, 4))
        self.combo_mode.bind("<<ComboboxSelected>>", lambda e: self.on_mode_change())

        self.lbl_mode_desc = tk.Label(
            card_mode,
            textvariable=self.mode_desc_var,
            font=("Segoe UI", 8, "italic"),
            fg=self.c_slate,
            bg=self.c_card,
            wraplength=700,
            justify="left"
        )
        self.lbl_mode_desc.pack(anchor="w", pady=(0, 2))

        # -------------------------------------------------------------
        # 3. Card 1: CSV Data Sources
        # -------------------------------------------------------------
        card1 = tk.LabelFrame(
            main_frame,
            text=" 2. CSV Data Sources & Scripts ",
            font=("Segoe UI", 10, "bold"),
            fg=self.c_navy,
            bg=self.c_card,
            padx=14,
            pady=8,
            relief="solid",
            bd=1
        )
        card1.pack(fill="x", pady=(0, 8))

        # Configure columns: Col 0 expands, Col 1 fixed for buttons
        card1.columnconfigure(0, weight=1)

        # Row 1: Device Details CSV (Full Ninja Export)
        self.lbl_d1 = tk.Label(card1, text="Device Details CSV (Full Ninja Export):", font=("Segoe UI", 9, "bold"), bg=self.c_card, fg=self.c_text)
        self.lbl_d1.grid(row=0, column=0, sticky="w", pady=(2, 1))

        self.entry_d1 = ttk.Entry(card1, textvariable=self.details_path)
        self.entry_d1.grid(row=1, column=0, sticky="ew", padx=(0, 8), pady=(0, 2))
        self.btn_d1 = ttk.Button(card1, text="Browse...", command=self.browse_details)
        self.btn_d1.grid(row=1, column=1, pady=(0, 2))

        self.badge_d1 = tk.Label(card1, textvariable=self.details_status, font=("Segoe UI", 8), bg=self.c_card, fg=self.c_gray)
        self.badge_d1.grid(row=2, column=0, sticky="w", pady=(0, 4))

        # Row 2: Custom Fields CSV (AD Install Date)
        self.lbl_d0 = tk.Label(card1, text="Custom Fields CSV (AD Install Date):", font=("Segoe UI", 9, "bold"), bg=self.c_card, fg=self.c_text)
        self.lbl_d0.grid(row=3, column=0, sticky="w", pady=(2, 1))

        self.entry_d0 = ttk.Entry(card1, textvariable=self.custom_path)
        self.entry_d0.grid(row=4, column=0, sticky="ew", padx=(0, 8), pady=(0, 2))
        self.btn_d0 = ttk.Button(card1, text="Browse...", command=self.browse_custom)
        self.btn_d0.grid(row=4, column=1, pady=(0, 2))

        self.badge_d0 = tk.Label(card1, textvariable=self.custom_status, font=("Segoe UI", 8), bg=self.c_card, fg=self.c_gray)
        self.badge_d0.grid(row=5, column=0, sticky="w", pady=(0, 4))

        # Row 3: Huntress Agents CSV
        self.lbl_dh = tk.Label(card1, text="Huntress Agents CSV:", font=("Segoe UI", 9, "bold"), bg=self.c_card, fg=self.c_text)
        self.lbl_dh.grid(row=6, column=0, sticky="w", pady=(2, 1))

        self.entry_dh = ttk.Entry(card1, textvariable=self.huntress_path)
        self.entry_dh.grid(row=7, column=0, sticky="ew", padx=(0, 8), pady=(0, 2))

        btn_h_frame = tk.Frame(card1, bg=self.c_card)
        btn_h_frame.grid(row=7, column=1, pady=(0, 2), sticky="w")
        self.btn_dh = ttk.Button(btn_h_frame, text="Browse...", command=self.browse_huntress)
        self.btn_dh.pack(side="left", padx=(0, 4))
        self.btn_ch = ttk.Button(btn_h_frame, text="Clear", command=self.clear_huntress)
        self.btn_ch.pack(side="left")

        self.badge_dh = tk.Label(card1, textvariable=self.huntress_status, font=("Segoe UI", 8), bg=self.c_card, fg=self.c_gray)
        self.badge_dh.grid(row=8, column=0, sticky="w", pady=(0, 4))

        # Row 4: Active Directory Raw Export CSV + PowerShell Generator Button
        self.lbl_ad = tk.Label(card1, text="Active Directory Raw Export (Computers_Raw_*.csv):", font=("Segoe UI", 9, "bold"), bg=self.c_card, fg=self.c_text)
        self.lbl_ad.grid(row=9, column=0, sticky="w", pady=(2, 1))

        self.entry_ad = ttk.Entry(card1, textvariable=self.ad_path)
        self.entry_ad.grid(row=10, column=0, sticky="ew", padx=(0, 8), pady=(0, 2))

        btn_ad_frame = tk.Frame(card1, bg=self.c_card)
        btn_ad_frame.grid(row=10, column=1, pady=(0, 2), sticky="w")
        self.btn_gen_ad_csv = tk.Button(
            btn_ad_frame,
            text="Generate CSV",
            font=("Segoe UI", 8, "bold"),
            bg="#D9E2EC",
            fg=self.c_navy,
            relief="groove",
            padx=6,
            pady=1,
            cursor="hand2",
            command=self.start_ad_csv_export_thread
        )
        self.btn_gen_ad_csv.pack(side="left", padx=(0, 4))
        self.btn_ad_browse = ttk.Button(btn_ad_frame, text="Browse...", command=self.browse_ad)
        self.btn_ad_browse.pack(side="left")

        self.badge_ad = tk.Label(card1, textvariable=self.ad_status, font=("Segoe UI", 8), bg=self.c_card, fg=self.c_gray)
        self.badge_ad.grid(row=11, column=0, sticky="w", pady=(0, 4))

        # Row 5: Auto-detect Button inside Card 1
        self.btn_detect = ttk.Button(
            card1,
            text="Auto-Detect All Recent Exports (< 7 Days)",
            command=lambda: self.auto_detect_files(notify=True)
        )
        self.btn_detect.grid(row=12, column=0, columnspan=2, sticky="w", pady=(4, 2))

        # -------------------------------------------------------------
        # 4. Card 2: Output Destination
        # -------------------------------------------------------------
        card2 = tk.LabelFrame(
            main_frame,
            text=" 3. Excel Output Destination ",
            font=("Segoe UI", 10, "bold"),
            fg=self.c_navy,
            bg=self.c_card,
            padx=14,
            pady=8,
            relief="solid",
            bd=1
        )
        card2.pack(fill="x", pady=(0, 8))

        entry_out = ttk.Entry(card2, textvariable=self.output_path)
        entry_out.pack(side="left", fill="x", expand=True, padx=(0, 8))
        btn_out = ttk.Button(card2, text="Change...", command=self.browse_output)
        btn_out.pack(side="right")

        # -------------------------------------------------------------
        # 5. Action & Progress Section
        # -------------------------------------------------------------
        action_frame = tk.Frame(main_frame, bg=self.c_bg)
        action_frame.pack(fill="x", pady=(2, 4))

        self.btn_run = tk.Button(
            action_frame,
            text="Generate Selected Report",
            font=("Segoe UI", 11, "bold"),
            bg=self.c_navy,
            fg="#FFFFFF",
            activebackground=self.c_slate,
            activeforeground="#FFFFFF",
            relief="flat",
            padx=20,
            pady=8,
            cursor="hand2",
            command=self.start_pipeline_thread
        )
        self.btn_run.pack(fill="x")

        self.progress_bar = ttk.Progressbar(main_frame, mode="indeterminate")
        self.progress_bar.pack(fill="x", pady=(4, 4))

        # -------------------------------------------------------------
        # 6. Live Activity Console
        # -------------------------------------------------------------
        card3 = tk.LabelFrame(
            main_frame,
            text=" Live Activity Log ",
            font=("Segoe UI", 9, "bold"),
            fg="#595959",
            bg=self.c_card,
            padx=8,
            pady=6,
            relief="solid",
            bd=1
        )
        card3.pack(fill="both", expand=True, pady=(0, 6))

        self.log_text = tk.Text(card3, font=("Consolas", 9), bg="#FAFAFA", fg="#1F2937", wrap="word", height=7, relief="flat")
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(card3, orient="vertical", command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)

        # -------------------------------------------------------------
        # 7. Bottom Open File Actions
        # -------------------------------------------------------------
        bottom_frame = tk.Frame(main_frame, bg=self.c_bg)
        bottom_frame.pack(fill="x")

        self.btn_open_excel = ttk.Button(
            bottom_frame,
            text="Open Generated Excel Report",
            state="disabled",
            command=self.open_excel_file
        )
        self.btn_open_excel.pack(side="left", padx=(0, 10))

        self.btn_open_tasks = ttk.Button(
            bottom_frame,
            text="View Tasks CSV",
            command=self.open_tasks_folder
        )
        self.btn_open_tasks.pack(side="left", padx=(0, 10))

        self.btn_sync_google = ttk.Button(
            bottom_frame,
            text="Sync to Google Tasks",
            command=self.open_google_tasks_sync_gui
        )
        self.btn_sync_google.pack(side="left", padx=(0, 10))

        self.btn_open_folder = ttk.Button(
            bottom_frame,
            text="Open Folder Location",
            command=self.open_output_folder
        )
        self.btn_open_folder.pack(side="left")

    # -----------------------------------------------------------------
    # Mode Handling & Dynamic UI Updates
    # -----------------------------------------------------------------
    def on_mode_change(self):
        idx = self.combo_mode.current()
        self.selected_mode_idx.set(idx)
        mode = self.get_current_mode()
        m_id = mode["id"]

        self.mode_desc_var.set(f"Folder: reports/{mode['subfolder']}/  |  {mode['desc']}")

        # Update Default Output Path (unless manually modified to a custom path)
        if not self.is_custom_output:
            today_str = datetime.now().strftime("%m%d%Y")
            if m_id == "ad_only":
                today_str = datetime.now().strftime("%Y%m%d")
            out_dir = os.path.join(os.getcwd(), "reports", mode["subfolder"])
            os.makedirs(out_dir, exist_ok=True)
            new_out = os.path.abspath(os.path.join(out_dir, f"{mode['file_prefix']}_{today_str}.xlsx"))
            self.output_path.set(new_out)

        # Update Run Button Text
        if m_id == "extensive":
            self.btn_run.config(text="Generate Extensive Report (Ninja + AD + Huntress)")
        elif m_id == "ninja_ad":
            self.btn_run.config(text="Generate Ninja & AD Inventory Report")
        elif m_id == "ninja_huntress":
            self.btn_run.config(text="Generate Ninja & Huntress Report (No AD Lifecycle)")
        elif m_id == "ninja_only":
            self.btn_run.config(text="Generate Ninja Inventory Report (No AD Lifecycle)")
        elif m_id == "ad_only":
            self.btn_run.config(text="Generate Active Directory Excel Report")
        elif m_id == "huntress_only":
            self.btn_run.config(text="Generate Standalone Huntress Report")

        # Enable/Disable Rows based on selected mode
        self.update_widget_states_for_mode(m_id)
        self.update_badges()

    def update_widget_states_for_mode(self, m_id):
        # Helper to set state
        def set_row_state(lbl, entry, btn, badge, status_var, is_needed, not_needed_msg="Not used in this mode"):
            if is_needed:
                lbl.config(fg=self.c_text)
                entry.config(state="normal")
                btn.config(state="normal")
            else:
                lbl.config(fg=self.c_gray)
                entry.config(state="disabled")
                btn.config(state="disabled")
                status_var.set(not_needed_msg)
                badge.config(fg=self.c_gray)

        # Ninja Device Details
        ninja_needed = m_id in ["extensive", "ninja_ad", "ninja_huntress", "ninja_only"]
        set_row_state(self.lbl_d1, self.entry_d1, self.btn_d1, self.badge_d1, self.details_status, ninja_needed)

        # AD Custom Fields
        custom_needed = m_id in ["extensive", "ninja_ad"]
        not_msg_custom = "Skipped — Lifecycle tab will be omitted" if m_id in ["ninja_huntress", "ninja_only"] else "Not used in this mode"
        set_row_state(self.lbl_d0, self.entry_d0, self.btn_d0, self.badge_d0, self.custom_status, custom_needed, not_msg_custom)

        # Huntress Agents
        huntress_needed = m_id in ["extensive", "ninja_huntress", "huntress_only"]
        if huntress_needed:
            self.lbl_dh.config(fg=self.c_text)
            self.entry_dh.config(state="normal")
            self.btn_dh.config(state="normal")
            self.btn_ch.config(state="normal")
        else:
            self.lbl_dh.config(fg=self.c_gray)
            self.entry_dh.config(state="disabled")
            self.btn_dh.config(state="disabled")
            self.btn_ch.config(state="disabled")
            self.huntress_status.set("Not included in this report mode")
            self.badge_dh.config(fg=self.c_gray)

        # AD Raw Export
        ad_raw_needed = m_id in ["extensive", "ninja_ad", "ad_only"]
        if ad_raw_needed:
            self.lbl_ad.config(fg=self.c_text)
            self.entry_ad.config(state="normal")
            self.btn_gen_ad_csv.config(state="normal")
            self.btn_ad_browse.config(state="normal")
        else:
            self.lbl_ad.config(fg=self.c_gray)
            self.entry_ad.config(state="disabled")
            self.btn_gen_ad_csv.config(state="disabled")
            self.btn_ad_browse.config(state="disabled")
            self.ad_status.set("Not used in this mode (only for Extensive, Ninja+AD, or AD Report)")
            self.badge_ad.config(fg=self.c_gray)

    # -----------------------------------------------------------------
    # Logging & Validation
    # -----------------------------------------------------------------
    def log(self, message):
        self.log_text.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see("end")

    def validate_file(self, path, expected_type):
        if not path or not path.strip():
            return False, "File not specified"
        if not os.path.exists(path):
            return False, "File does not exist"
        ctype = classify_csv(path)
        if ctype == expected_type or (expected_type == "AD_COMPUTERS_RAW" and "Computers_Raw" in os.path.basename(path)):
            try:
                mtime = os.path.getmtime(path)
                age_days = round((time.time() - mtime) / 86400.0, 1)
                if age_days > 7.0:
                    return "warn", f"[Warning] Verified ({age_days}d old): {os.path.basename(path)}"
            except Exception:
                pass
            return True, f"[Verified] {os.path.basename(path)}"
        elif ctype == "VISIBLE_COLUMNS_REDUNDANT":
            return False, "[Warning] Contains only visible columns (Redundant)"
        else:
            return False, f"[Warning] Detected as '{ctype}', expected '{expected_type}'"

    def update_badges(self):
        m_id = self.get_current_mode()["id"]

        def apply_badge(v, msg, status_var, badge_widget):
            status_var.set(msg)
            if v is True:
                badge_widget.config(fg=self.c_green)
            elif v == "warn":
                badge_widget.config(fg=self.c_warn)
            else:
                badge_widget.config(fg=self.c_red)

        # Device Details
        if m_id in ["extensive", "ninja_ad", "ninja_huntress", "ninja_only"]:
            v1, msg1 = self.validate_file(self.details_path.get(), "DEVICE_DETAILS")
            apply_badge(v1, msg1, self.details_status, self.badge_d1)

        # Custom Fields
        if m_id in ["extensive", "ninja_ad"]:
            v2, msg2 = self.validate_file(self.custom_path.get(), "CUSTOM_FIELDS")
            apply_badge(v2, msg2, self.custom_status, self.badge_d0)

        # Huntress
        if m_id in ["extensive", "ninja_huntress", "huntress_only"]:
            vh, msgh = self.validate_file(self.huntress_path.get(), "HUNTRESS_AGENTS")
            apply_badge(vh, msgh, self.huntress_status, self.badge_dh)

        # AD Raw
        if m_id in ["extensive", "ninja_ad", "ad_only"]:
            if not self.ad_path.get().strip():
                self.ad_status.set("Click 'Generate CSV' or Browse an existing Computers_Raw export")
                self.badge_ad.config(fg=self.c_warn)
            else:
                va, msga = self.validate_file(self.ad_path.get(), "AD_COMPUTERS_RAW")
                apply_badge(va, msga, self.ad_status, self.badge_ad)

    # -----------------------------------------------------------------
    # File Pickers
    # -----------------------------------------------------------------
    def browse_details(self):
        f = filedialog.askopenfilename(title="Select NinjaOne Device Details CSV", filetypes=[("CSV Files", "*.csv")])
        if f:
            self.details_path.set(os.path.abspath(f))
            self.update_badges()

    def browse_custom(self):
        f = filedialog.askopenfilename(title="Select NinjaOne Custom Fields CSV", filetypes=[("CSV Files", "*.csv")])
        if f:
            self.custom_path.set(os.path.abspath(f))
            self.update_badges()

    def browse_huntress(self):
        f = filedialog.askopenfilename(title="Select Huntress Agents CSV", filetypes=[("CSV Files", "*.csv")])
        if f:
            self.huntress_path.set(os.path.abspath(f))
            self.update_badges()

    def clear_huntress(self):
        self.huntress_path.set("")
        self.update_badges()
        self.log("Cleared Huntress CSV input.")

    def browse_ad(self):
        f = filedialog.askopenfilename(title="Select Active Directory Raw Computers CSV", filetypes=[("CSV Files", "Computers_Raw_*.csv;*.csv")])
        if f:
            self.ad_path.set(os.path.abspath(f))
            self.update_badges()

    def browse_output(self):
        f = filedialog.asksaveasfilename(
            title="Select Output Excel Workbook Path",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if f:
            self.output_path.set(os.path.abspath(f))
            self.is_custom_output = True

    # -----------------------------------------------------------------
    # Auto-Detection Logic
    # -----------------------------------------------------------------
    def auto_detect_files(self, notify=True):
        self.log("Scanning workspace and Downloads folder for recent exports (< 7 Days)...")
        res = find_best_ninja_csvs(max_age_days=7)
        ad_file, stale_ad = find_best_ad_csv(max_age_days=7)
        found_fresh = False
        stale_auto_filled = []

        # Device Details
        if res.get("details_file"):
            self.details_path.set(os.path.abspath(res["details_file"]))
            self.log(f"Auto-selected Device Details: {os.path.basename(res['details_file'])}")
            found_fresh = True
        elif res.get("stale_details"):
            st_f, st_d = res["stale_details"]
            self.details_path.set(os.path.abspath(st_f))
            self.log(f"[Auto-selected] older Device Details CSV: {os.path.basename(st_f)} ({st_d}d old)")
            stale_auto_filled.append(f"• Device Details: {os.path.basename(st_f)} ({st_d}d old)")

        # Custom Fields
        if res.get("custom_file"):
            self.custom_path.set(os.path.abspath(res["custom_file"]))
            self.log(f"Auto-selected Custom Fields: {os.path.basename(res['custom_file'])}")
            found_fresh = True
        elif res.get("stale_custom"):
            st_f, st_d = res["stale_custom"]
            self.custom_path.set(os.path.abspath(st_f))
            self.log(f"[Auto-selected] older Custom Fields CSV: {os.path.basename(st_f)} ({st_d}d old)")
            stale_auto_filled.append(f"• Custom Fields: {os.path.basename(st_f)} ({st_d}d old)")

        # Huntress Agents
        if res.get("huntress_file"):
            self.huntress_path.set(os.path.abspath(res["huntress_file"]))
            self.log(f"Auto-selected Huntress Agents: {os.path.basename(res['huntress_file'])}")
            found_fresh = True
        elif res.get("stale_huntress"):
            st_f, st_d = res["stale_huntress"]
            self.huntress_path.set(os.path.abspath(st_f))
            self.log(f"[Auto-selected] older Huntress CSV: {os.path.basename(st_f)} ({st_d}d old)")
            stale_auto_filled.append(f"• Huntress Agents: {os.path.basename(st_f)} ({st_d}d old)")

        # AD Raw Export
        if ad_file:
            self.ad_path.set(os.path.abspath(ad_file))
            self.log(f"Auto-selected AD Raw Computers: {os.path.basename(ad_file)}")
            found_fresh = True
        elif stale_ad:
            st_f, st_d = stale_ad
            self.ad_path.set(os.path.abspath(st_f))
            self.log(f"[Auto-selected] older AD Raw CSV: {os.path.basename(st_f)} ({st_d}d old)")
            stale_auto_filled.append(f"• AD Raw Export: {os.path.basename(st_f)} ({st_d}d old)")

        if res.get("redundant_found"):
            for red in res["redundant_found"]:
                self.log(f"Ignored redundant file: {os.path.basename(red)}")

        self.update_badges()

        if notify:
            msg_items = []
            if res.get("details_file"):
                msg_items.append(f"• Device Details: {os.path.basename(res['details_file'])}")
            if res.get("custom_file"):
                msg_items.append(f"• Custom Fields: {os.path.basename(res['custom_file'])}")
            if res.get("huntress_file"):
                msg_items.append(f"• Huntress Agents: {os.path.basename(res['huntress_file'])}")
            if ad_file:
                msg_items.append(f"• AD Raw Export: {os.path.basename(ad_file)}")

            if msg_items:
                stale_note = ""
                if stale_auto_filled:
                    stale_note = "\n\nAlso auto-filled existing older exports (> 7 days):\n" + "\n".join(stale_auto_filled)
                messagebox.showinfo("Smart Detection", "Auto-detected recent exports (< 7 days):\n\n" + "\n".join(msg_items) + stale_note)
            elif stale_auto_filled:
                messagebox.showwarning("Smart Detection", "No fresh exports (< 7 days) found. Auto-loaded existing older files:\n\n" + "\n".join(stale_auto_filled))
            else:
                messagebox.showwarning("Smart Detection", "No export files detected automatically. Please browse manually.")

    # -----------------------------------------------------------------
    # AD PowerShell Script Generator
    # -----------------------------------------------------------------
    def start_ad_csv_export_thread(self):
        if self.is_running_ad_export:
            return

        ps_script = os.path.join(os.getcwd(), "Export_AD_Computers_Raw.ps1")
        if not os.path.exists(ps_script):
            repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
            candidates = [
                os.path.join(repo_root, "Export_AD_Computers_Raw.ps1"),
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "ad", "Export_AD_Computers_Raw.ps1")
            ]
            for c in candidates:
                if os.path.exists(c):
                    ps_script = os.path.abspath(c)
                    break
        if not os.path.exists(ps_script):
            messagebox.showerror("Script Not Found", f"PowerShell script not found:\n{ps_script}")
            return

        self.is_running_ad_export = True
        self.btn_gen_ad_csv.config(state="disabled", text="Exporting...")
        self.btn_run.config(state="disabled")
        self.progress_bar.start(10)

        self.log("--- Executing Active Directory PowerShell Query ---")
        self.log("Running Export_AD_Computers_Raw.ps1 via PowerShell...")

        thread = threading.Thread(target=self.run_ad_export_worker, args=(ps_script,))
        thread.daemon = True
        thread.start()

    def run_ad_export_worker(self, ps_script):
        try:
            cmd = [
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-File", ps_script
            ]
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                cwd=os.getcwd()
            )

            stdout, stderr = process.communicate()

            if stdout:
                for line in stdout.strip().splitlines():
                    if line.strip():
                        self.log(f"[PS] {line.strip()}")

            if process.returncode != 0:
                err_text = stderr.strip() if stderr else f"Process exited with code {process.returncode}"
                self.log(f"[PS ERROR] {err_text}")
                self.root.after(0, lambda: self.on_ad_export_failed(err_text))
            else:
                self.root.after(0, self.on_ad_export_success)
        except Exception as e:
            self.log(f"[PS EXCEPTION] {str(e)}")
            self.root.after(0, lambda: self.on_ad_export_failed(str(e)))

    def on_ad_export_success(self):
        self.is_running_ad_export = False
        self.progress_bar.stop()
        self.btn_gen_ad_csv.config(state="normal", text="Generate CSV")
        self.btn_run.config(state="normal")

        # Detect the newly created file
        ad_file, _ = find_best_ad_csv(max_age_days=1)
        if ad_file and os.path.exists(ad_file):
            self.ad_path.set(os.path.abspath(ad_file))
            self.log(f"SUCCESS: Generated and selected {os.path.basename(ad_file)}")
            self.update_badges()
            messagebox.showinfo("AD Export Complete", f"Active Directory raw CSV export generated successfully!\n\nFile: {os.path.basename(ad_file)}\n\nYou can now generate the executive Excel report.")
        else:
            self.update_badges()
            messagebox.showwarning("Notice", "PowerShell script completed, but the newest CSV could not be automatically located. Please browse manually.")

    def on_ad_export_failed(self, err_msg):
        self.is_running_ad_export = False
        self.progress_bar.stop()
        self.btn_gen_ad_csv.config(state="normal", text="Generate CSV")
        self.btn_run.config(state="normal")
        self.update_badges()
        messagebox.showerror("PowerShell Error", f"Failed to execute Active Directory export:\n\n{err_msg}\n\nEnsure this machine is domain-joined or connected via VPN with ActiveDirectory module installed.")

    # -----------------------------------------------------------------
    # Pipeline Execution
    # -----------------------------------------------------------------
    def start_pipeline_thread(self):
        mode = self.get_current_mode()
        m_id = mode["id"]

        d1 = self.details_path.get().strip()
        d0 = self.custom_path.get().strip()
        dh = self.huntress_path.get().strip()
        ad = self.ad_path.get().strip()
        out = self.output_path.get().strip()

        # Enforce Restrictions
        if m_id in ["extensive", "ninja_ad", "ninja_huntress", "ninja_only"]:
            if not d1 or not os.path.exists(d1):
                messagebox.showerror("Missing File", "Please select a valid Ninja Device Details CSV export.")
                return

        if m_id in ["extensive", "ninja_ad"]:
            if not d0 or not os.path.exists(d0):
                messagebox.showerror("Missing File", "Please select a valid Custom Fields CSV export (AD Install Date).")
                return

        if m_id in ["extensive", "ninja_huntress", "huntress_only"]:
            if not dh or not os.path.exists(dh):
                messagebox.showerror("Missing File", "Please select a valid Huntress Agents CSV export.")
                return

        if m_id in ["extensive", "ninja_ad", "ad_only"]:
            if not ad or not os.path.exists(ad):
                messagebox.showerror("Missing File", "No Active Directory raw CSV selected.\n\nPlease click 'Generate CSV' to export directly from AD, or browse an existing Computers_Raw_*.csv file.")
                return

        # Ensure destination folder exists
        out_folder = os.path.dirname(out)
        if out_folder:
            os.makedirs(out_folder, exist_ok=True)

        # Pre-flight check: Test if output Excel file is currently open/locked
        if os.path.exists(out):
            try:
                with open(out, "r+"):
                    pass
            except PermissionError:
                messagebox.showerror(
                    "File Locked by Excel",
                    f"Cannot write to destination Excel file:\n\n{out}\n\n"
                    "The file appears to be open in Microsoft Excel or another program.\n\n"
                    "Please save and close the workbook in Excel, then click Generate again."
                )
                self.log(f"PRE-FLIGHT ERROR: Destination file is locked: {out}")
                return
            except Exception:
                pass

        self.btn_run.config(state="disabled")
        self.btn_open_excel.config(state="disabled")
        self.progress_bar.start(10)
        self.log(f"--- Starting Pipeline Execution [{mode['name']}] ---")

        thread = threading.Thread(target=self.run_pipeline_worker, args=(m_id, d1, d0, dh, ad, out))
        thread.daemon = True
        thread.start()

    def run_pipeline_worker(self, m_id, d1, d0, dh, ad, out):
        try:
            if m_id in ["extensive", "ninja_ad", "ninja_huntress", "ninja_only"]:
                # Parameter mapping based on selected mode
                custom_file = d0 if m_id in ["extensive", "ninja_ad"] else None
                huntress_file = dh if m_id in ["extensive", "ninja_huntress"] else None
                ad_raw_file = ad if m_id in ["extensive", "ninja_ad"] else None

                self.log(f"Reading and cleaning: {os.path.basename(d1)}...")
                if custom_file:
                    self.log(f"Merging AD Install Dates from: {os.path.basename(custom_file)}...")
                else:
                    self.log("Skipping AD Install Dates (Lifecycle tab will be omitted).")

                if huntress_file:
                    self.log(f"Merging Huntress endpoint report: {os.path.basename(huntress_file)}...")
                else:
                    self.log("Huntress report not included for this mode.")

                if ad_raw_file:
                    self.log(f"Merging Active Directory raw export: {os.path.basename(ad_raw_file)}...")

                build_inventory_sheet(d1, custom_file, out, csv_huntress_path=huntress_file, csv_ad_path=ad_raw_file)
                self.log("Master 'Ninja Inventory' tab successfully created and formatted.")
                if ad_raw_file:
                    self.log("'AD Inventory', 'AD Dashboard', and 'AD Cross ref' tabs included.")
                if huntress_file:
                    self.log("'Huntress Inventory', 'Huntress Dashboard', and 'Huntress Cross ref' tabs included.")
                self.log("Executive Dashboard tabs generated with charts & KPIs.")
                if m_id == "extensive":
                    today_str = datetime.now().strftime("%m%d%Y")
                    out_dir = os.path.dirname(os.path.abspath(out))
                    parent_dir = os.path.dirname(out_dir) if os.path.basename(out_dir).lower() == "extensive reports" else out_dir
                    tasks_csv = os.path.join(parent_dir, "Tasks", f"flaggedDevices_{today_str}.csv")
                    if not os.path.exists(tasks_csv):
                        tasks_csv = os.path.join("reports", "Tasks", f"flaggedDevices_{today_str}.csv")
                    if os.path.exists(tasks_csv):
                        self.log(f"Google Tasks CSV created: {tasks_csv}")
                    else:
                        self.log("Notice: Google Tasks CSV (flagged devices) was not generated.")

            elif m_id == "ad_only":
                self.log(f"Processing Active Directory raw export: {os.path.basename(ad)}...")
                build_ad_report_sheet(ad, output_xlsx=out)
                self.log("AD Computers master inventory tab formatted with auto-filters & status highlights.")
                self.log("Executive Summary & Audit Dashboard generated with clinic allocations & hygiene targets.")

            elif m_id == "huntress_only":
                self.log(f"Processing Huntress agents export: {os.path.basename(dh)}...")
                build_huntress_report(dh, output_xlsx=out)
                self.log("Huntress Agents inventory tab created with auto-filters & inactivity tiers.")
                self.log("Executive Summary dashboard generated with OS distribution chart & dormancy watchlist.")

            self.log(f"SUCCESS: Report saved to {out}")
            self.root.after(0, self.on_pipeline_success)

        except PermissionError as pe:
            err_msg = (
                f"Permission denied while saving Excel report:\n\n{out}\n\n"
                "The file appears to be open in Microsoft Excel or another program.\n\n"
                "Please close Microsoft Excel and try again."
            )
            self.log(f"PERMISSION ERROR: {pe}")
            self.root.after(0, lambda: messagebox.showerror("File Locked by Excel", err_msg))
            self.root.after(0, self.on_pipeline_error)
        except Exception as e:
            self.log(f"ERROR: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Pipeline Error", f"An error occurred during report generation:\n\n{str(e)}"))
            self.root.after(0, self.on_pipeline_error)

    def on_pipeline_success(self):
        self.progress_bar.stop()
        self.btn_run.config(state="normal")
        self.btn_open_excel.config(state="normal")
        messagebox.showinfo("Complete", "Executive Report successfully generated!\n\nClick 'Open Generated Excel Report' to view the results.")

    def on_pipeline_error(self):
        self.progress_bar.stop()
        self.btn_run.config(state="normal")

    def open_excel_file(self):
        out = self.output_path.get()
        if os.path.exists(out):
            try:
                os.startfile(out)
            except Exception as e:
                messagebox.showerror("Error", f"Could not launch Excel: {e}")
        else:
            messagebox.showwarning("Not Found", "Output file not found. Please generate it first.")

    def open_tasks_folder(self):
        tasks_dir = os.path.abspath(os.path.join("reports", "Tasks"))
        if not os.path.exists(tasks_dir):
            os.makedirs(tasks_dir, exist_ok=True)
        try:
            os.startfile(tasks_dir)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open tasks folder: {e}")

    def open_google_tasks_sync_gui(self):
        try:
            try:
                subprocess.Popen([sys.executable, "-m", "humble_mumble.google_tasks.google_tasks_gui"])
            except Exception:
                sync_script = os.path.join(os.getcwd(), "google_tasks_gui.py")
                if not os.path.exists(sync_script):
                    sync_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "google_tasks", "google_tasks_gui.py")
                subprocess.Popen([sys.executable, sync_script])
        except Exception as e:
            messagebox.showerror("Error", f"Could not launch Google Tasks Studio:\n{str(e)}")

    def open_output_folder(self):
        out = self.output_path.get()
        folder = os.path.dirname(out) if os.path.exists(out) else os.getcwd()
        try:
            os.startfile(folder)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder: {e}")

def main():
    root = tk.Tk()
    app = ReportingStudioApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
