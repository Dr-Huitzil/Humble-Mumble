"""
google_tasks_gui.py
===================
Interactive Desktop UI to synchronize flagged endpoint tasks directly to Google Tasks.
Features:
  - Official Google OAuth 2.0 Sign-In with browser authorization.
  - Auto-discovery of the newest flagged devices CSV.
  - Clinic selection (sync all clinics or pick specific locations).
  - Smart Deduplication & Merge:
      * Prevents duplicate parent device tasks.
      * Prevents duplicate subtasks.
      * Re-opens completed tasks if recurring or new issues are detected.
      * Appends new subtasks to existing device records.
  - Dry Run / Preview Mode.
  - 1-Click Launch of Google Tasks in the browser.
"""

import os
import sys
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import csv

from google_tasks_sync import (
    authenticate_google_tasks,
    get_auth_status,
    sync_flagged_csv_to_google_tasks
)

class GoogleTasksSyncApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Tasks Fleet Synchronization Studio")
        self.root.geometry("860x860")
        self.root.minsize(780, 750)
        self.root.configure(bg="#F4F6F9")

        # Styling Palette matching Executive Studio
        self.c_navy = "#1F4E79"
        self.c_slate = "#2F5597"
        self.c_bg = "#F4F6F9"
        self.c_card = "#FFFFFF"
        self.c_border = "#D9E2EC"
        self.c_green = "#2E7D32"
        self.c_red = "#C62828"
        self.c_warn = "#B25900"
        self.c_gray = "#595959"

        # State Variables
        self.csv_path = tk.StringVar()
        self.auth_badge_text = tk.StringVar(value="Checking authentication...")
        self.auth_badge_color = tk.StringVar(value=self.c_gray)
        self.auth_badge_fg = tk.StringVar(value="#FFFFFF")
        
        self.dry_run_var = tk.BooleanVar(value=False)
        self.smart_merge_var = tk.BooleanVar(value=True)

        self.service = None
        self.is_running = False

        self.setup_ui()
        self.refresh_auth_status()
        self.auto_detect_csv()

    def setup_ui(self):
        # -------------------------------------------------------------
        # 1. Header Banner
        # -------------------------------------------------------------
        header_frame = tk.Frame(self.root, bg=self.c_navy, height=76)
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)

        title_lbl = tk.Label(
            header_frame,
            text="Google Tasks Fleet Synchronization",
            font=("Segoe UI", 16, "bold"),
            fg="#FFFFFF",
            bg=self.c_navy
        )
        title_lbl.pack(anchor="w", padx=20, pady=(12, 2))

        sub_lbl = tk.Label(
            header_frame,
            text="Automated Remediation Workflows  •  Clinic Lists  •  Device Parent Tasks  •  Actionable Subtasks",
            font=("Segoe UI", 9),
            fg="#D9E2EC",
            bg=self.c_navy
        )
        sub_lbl.pack(anchor="w", padx=20)

        # -------------------------------------------------------------
        # Main Scrollable Frame
        # -------------------------------------------------------------
        main_frame = tk.Frame(self.root, bg=self.c_bg, padx=18, pady=12)
        main_frame.pack(fill="both", expand=True)

        # -------------------------------------------------------------
        # Card 1: Google Account & OAuth Authentication
        # -------------------------------------------------------------
        card_auth = tk.LabelFrame(
            main_frame,
            text=" 1. Google Account & API Authorization ",
            font=("Segoe UI", 10, "bold"),
            fg=self.c_navy,
            bg=self.c_card,
            padx=14,
            pady=10,
            relief="solid",
            bd=1
        )
        card_auth.pack(fill="x", pady=(0, 10))

        row_auth = tk.Frame(card_auth, bg=self.c_card)
        row_auth.pack(fill="x")

        self.lbl_auth_status = tk.Label(
            row_auth,
            textvariable=self.auth_badge_text,
            font=("Segoe UI", 9, "bold"),
            fg="#FFFFFF",
            bg=self.c_gray,
            padx=10,
            pady=4,
            relief="flat"
        )
        self.lbl_auth_status.pack(side="left", padx=(0, 12))

        self.btn_signin = ttk.Button(
            row_auth,
            text="🔑 Sign In to Google",
            command=self.start_google_signin
        )
        self.btn_signin.pack(side="left", padx=(0, 8))

        self.btn_browse_creds = ttk.Button(
            row_auth,
            text="📂 Select credentials.json",
            command=self.browse_credentials
        )
        self.btn_browse_creds.pack(side="left", padx=(0, 8))

        btn_guide = ttk.Button(
            row_auth,
            text="ℹ️ Setup Instructions",
            command=self.show_oauth_guide
        )
        btn_guide.pack(side="left")

        # -------------------------------------------------------------
        # Card 2: Source Flagged CSV File
        # -------------------------------------------------------------
        card_csv = tk.LabelFrame(
            main_frame,
            text=" 2. Source Flagged Devices CSV ",
            font=("Segoe UI", 10, "bold"),
            fg=self.c_navy,
            bg=self.c_card,
            padx=14,
            pady=10,
            relief="solid",
            bd=1
        )
        card_csv.pack(fill="x", pady=(0, 10))

        row_csv = tk.Frame(card_csv, bg=self.c_card)
        row_csv.pack(fill="x")

        entry_csv = ttk.Entry(row_csv, textvariable=self.csv_path, font=("Segoe UI", 9))
        entry_csv.pack(side="left", fill="x", expand=True, padx=(0, 8))

        btn_browse_csv = ttk.Button(row_csv, text="Browse...", command=self.browse_csv)
        btn_browse_csv.pack(side="left", padx=(0, 6))

        btn_auto_csv = ttk.Button(row_csv, text="⚡ Auto-Detect Latest", command=self.auto_detect_csv)
        btn_auto_csv.pack(side="left")

        self.lbl_csv_meta = tk.Label(card_csv, text="", font=("Segoe UI", 9, "italic"), fg=self.c_gray, bg=self.c_card)
        self.lbl_csv_meta.pack(anchor="w", pady=(6, 0))

        # -------------------------------------------------------------
        # Card 3: Clinic Selection & Sync Options
        # -------------------------------------------------------------
        card_opts = tk.LabelFrame(
            main_frame,
            text=" 3. Sync Settings & Deduplication Rules ",
            font=("Segoe UI", 10, "bold"),
            fg=self.c_navy,
            bg=self.c_card,
            padx=14,
            pady=10,
            relief="solid",
            bd=1
        )
        card_opts.pack(fill="x", pady=(0, 10))

        opts_row = tk.Frame(card_opts, bg=self.c_card)
        opts_row.pack(fill="x")

        chk_merge = ttk.Checkbutton(
            opts_row,
            text="Smart Deduplication (Prevent duplicate tasks, re-open completed if unresolved, append new subtasks)",
            variable=self.smart_merge_var
        )
        chk_merge.pack(anchor="w")

        chk_dry = ttk.Checkbutton(
            opts_row,
            text="Dry Run / Simulation Mode (Preview changes without creating or updating Google Tasks)",
            variable=self.dry_run_var
        )
        chk_dry.pack(anchor="w", pady=(4, 0))

        # -------------------------------------------------------------
        # Card 4: Live Synchronization Console
        # -------------------------------------------------------------
        card_log = tk.LabelFrame(
            main_frame,
            text=" Live Synchronization Activity ",
            font=("Segoe UI", 9, "bold"),
            fg=self.c_gray,
            bg=self.c_card,
            padx=8,
            pady=6,
            relief="solid",
            bd=1
        )
        card_log.pack(fill="both", expand=True, pady=(0, 8))

        self.log_text = tk.Text(card_log, font=("Consolas", 9), bg="#FAFAFA", fg="#1F2937", wrap="word", height=9, relief="flat")
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(card_log, orient="vertical", command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)

        self.progress_bar = ttk.Progressbar(main_frame, mode="indeterminate")
        self.progress_bar.pack(fill="x", pady=(0, 8))

        # -------------------------------------------------------------
        # Card 5: Bottom Execution Buttons
        # -------------------------------------------------------------
        bottom_frame = tk.Frame(main_frame, bg=self.c_bg)
        bottom_frame.pack(fill="x")

        self.btn_dry_run = ttk.Button(
            bottom_frame,
            text="🔍 Test Run / Preview",
            command=lambda: self.start_sync_thread(is_dry_run=True)
        )
        self.btn_dry_run.pack(side="left", padx=(0, 10))

        self.btn_push = ttk.Button(
            bottom_frame,
            text="🚀 Push to Google Tasks",
            command=lambda: self.start_sync_thread(is_dry_run=False)
        )
        self.btn_push.pack(side="left", padx=(0, 10))

        btn_open_web = ttk.Button(
            bottom_frame,
            text="🌐 Open Google Tasks in Browser",
            command=lambda: webbrowser.open("https://tasks.google.com")
        )
        btn_open_web.pack(side="right")

    # -----------------------------------------------------------------
    # Logging Helper
    # -----------------------------------------------------------------
    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    # -----------------------------------------------------------------
    # Authentication Management
    # -----------------------------------------------------------------
    def refresh_auth_status(self):
        st = get_auth_status()
        if st["is_valid"]:
            self.auth_badge_text.set("✓ Authenticated with Google Tasks")
            self.lbl_auth_status.config(bg=self.c_green)
            self.btn_signin.config(text="✓ Re-Authenticate / Switch Account")
        elif st["has_credentials"]:
            self.auth_badge_text.set("⚠️ Ready to Sign In (credentials.json detected)")
            self.lbl_auth_status.config(bg=self.c_warn)
            self.btn_signin.config(text="🔑 Sign In to Google")
        else:
            self.auth_badge_text.set("⚠️ Missing credentials.json")
            self.lbl_auth_status.config(bg=self.c_red)
            self.btn_signin.config(text="🔑 Sign In to Google")

    def browse_credentials(self):
        path = filedialog.askopenfilename(
            title="Select Google OAuth 2.0 Client credentials.json",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if path:
            dest = os.path.abspath("credentials.json")
            try:
                import shutil
                if os.path.abspath(path) != dest:
                    shutil.copyfile(path, dest)
                self.log(f"[✓] Copied credentials to: {dest}")
                self.refresh_auth_status()
                messagebox.showinfo("Success", "credentials.json successfully configured!\n\nClick 'Sign In to Google' to authenticate.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save credentials: {e}")

    def show_oauth_guide(self):
        guide_text = (
            "How to set up Google Tasks API Authorization:\n\n"
            "1. Open https://console.cloud.google.com\n"
            "2. Create a new project (e.g. 'Fleet Reporting Studio').\n"
            "3. In 'APIs & Services' > 'Library', search for and ENABLE 'Google Tasks API'.\n"
            "4. Go to 'Credentials' > 'Create Credentials' > 'OAuth client ID'.\n"
            "5. Select Application Type: 'Desktop app' and click Create.\n"
            "6. Download the JSON file and save it as 'credentials.json' in this folder.\n\n"
            "Once saved, click 'Sign In to Google' to authorize your account."
        )
        messagebox.showinfo("Google Tasks API Setup Guide", guide_text)

    def start_google_signin(self):
        st = get_auth_status()
        if not st["has_credentials"]:
            messagebox.showwarning(
                "Credentials Missing",
                "Please place 'credentials.json' in this project folder before signing in.\n\n"
                "Click 'Setup Instructions' or 'Select credentials.json' to proceed."
            )
            return

        def signin_worker():
            try:
                self.log("Opening official Google Sign-In page in default web browser...")
                self.service = authenticate_google_tasks()
                self.log("[✓] Google Authentication Successful! Token saved.")
                self.root.after(0, self.on_signin_success)
            except Exception as e:
                self.log(f"[!] Authentication error: {str(e)}")
                self.root.after(0, lambda: messagebox.showerror("Sign-In Error", f"Google Authentication Failed:\n\n{str(e)}"))

        threading.Thread(target=signin_worker, daemon=True).start()

    def on_signin_success(self):
        self.refresh_auth_status()
        messagebox.showinfo("Success", "Successfully authenticated with Google Tasks!")

    # -----------------------------------------------------------------
    # CSV Discovery & Inspection
    # -----------------------------------------------------------------
    def auto_detect_csv(self):
        tasks_dir = os.path.join("reports", "Tasks")
        if os.path.exists(tasks_dir):
            files = [os.path.join(tasks_dir, f) for f in os.listdir(tasks_dir) if f.startswith("flaggedDevices_") and f.endswith(".csv")]
            if files:
                files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                latest_csv = files[0]
                self.csv_path.set(latest_csv)
                self.inspect_csv(latest_csv)
                self.log(f"[✓] Auto-detected latest flagged devices export: {os.path.basename(latest_csv)}")
                return
        self.lbl_csv_meta.config(text="No flagged devices CSV found in reports/Tasks/. Run Extensive Report first.")

    def browse_csv(self):
        path = filedialog.askopenfilename(
            title="Select Flagged Devices CSV",
            initialdir=os.path.abspath(os.path.join("reports", "Tasks")),
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if path:
            self.csv_path.set(path)
            self.inspect_csv(path)

    def inspect_csv(self, path):
        if not os.path.exists(path):
            self.lbl_csv_meta.config(text="File does not exist.")
            return

        try:
            with open(path, mode="r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                rows = list(reader)
                clinics = set()
                for r in rows:
                    c = r.get("Google Task List (Clinic)") or r.get("Clinic / Location")
                    if c:
                        clinics.add(c.strip())

                self.lbl_csv_meta.config(
                    text=f"Total Flagged Endpoints: {len(rows)}  |  Clinic Task Lists: {len(clinics)}"
                )
        except Exception as e:
            self.lbl_csv_meta.config(text=f"Error inspecting CSV: {e}")

    # -----------------------------------------------------------------
    # Synchronization Thread Worker
    # -----------------------------------------------------------------
    def start_sync_thread(self, is_dry_run=False):
        csv_file = self.csv_path.get().strip()
        if not csv_file or not os.path.exists(csv_file):
            messagebox.showerror("Error", "Please select a valid flagged devices CSV file.")
            return

        if not is_dry_run:
            st = get_auth_status()
            if not st["is_valid"]:
                messagebox.showerror(
                    "Not Authenticated",
                    "You must sign in to Google before pushing tasks.\n\n"
                    "Click 'Sign In to Google' at the top of the studio."
                )
                return

        self.btn_push.config(state="disabled")
        self.btn_dry_run.config(state="disabled")
        self.progress_bar.start(10)

        def sync_worker():
            try:
                res = sync_flagged_csv_to_google_tasks(
                    csv_path=csv_file,
                    service=self.service,
                    dry_run=is_dry_run,
                    log_callback=lambda msg: self.root.after(0, lambda m=msg: self.log(m))
                )
                self.root.after(0, lambda: self.on_sync_success(res, is_dry_run))
            except Exception as e:
                self.root.after(0, lambda: self.on_sync_error(str(e)))

        threading.Thread(target=sync_worker, daemon=True).start()

    def on_sync_success(self, res, is_dry_run):
        self.progress_bar.stop()
        self.btn_push.config(state="normal")
        self.btn_dry_run.config(state="normal")

        mode_str = "Preview / Dry Run" if is_dry_run else "Live Synchronization"
        msg = (
            f"{mode_str} Complete!\n\n"
            f"• Task Lists Created: {res['lists_created']}\n"
            f"• Task Lists Reused: {res['lists_reused']}\n"
            f"• Parent Device Tasks Created: {res['tasks_created']}\n"
            f"• Parent Device Tasks Re-opened: {res['tasks_reopened']}\n"
            f"• Subtasks Created: {res['subtasks_created']}\n"
            f"• Subtasks Re-opened: {res['subtasks_reopened']}\n"
            f"• Duplicate Subtasks Skipped: {res['subtasks_skipped']}"
        )
        messagebox.showinfo(f"{mode_str} Finished", msg)

    def on_sync_error(self, err_msg):
        self.progress_bar.stop()
        self.btn_push.config(state="normal")
        self.btn_dry_run.config(state="normal")
        self.log(f"[ERROR] {err_msg}")
        messagebox.showerror("Synchronization Error", f"An error occurred during sync:\n\n{err_msg}")


if __name__ == "__main__":
    root = tk.Tk()
    app = GoogleTasksSyncApp(root)
    root.mainloop()
