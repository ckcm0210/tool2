# -*- coding: utf-8 -*-
"""
Created on Wed Jun 25 09:09:45 2025

@author: kccheng
"""

import tkinter as tk
from tkinter import ttk, messagebox
from worksheet_pane import WorksheetPane

class ExcelFormulaComparator:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Formula Comparator")
        self.root.geometry("860x1100")
        self.root.app = self

        self.left_pane = None
        self.right_pane = None

        self.setup_ui()

    def setup_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        button_frame = ttk.Frame(self.root, padding="10")
        button_frame.grid(row=0, column=0, sticky="ew")
        button_frame.columnconfigure(7, weight=1)

        large_button_font = ("Arial", 11, "bold")
        style = ttk.Style()
        style.configure("Large.TButton", font=large_button_font, padding=[10, 5])

        self.btn_scan1_full = ttk.Button(
            button_frame, text="Scan Worksheet1 (Full)",
            command=self.scan_left_full, style="Large.TButton")
        self.btn_scan1_full.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.btn_scan1_quick = ttk.Button(
            button_frame, text="Scan Worksheet1 (Quick)",
            command=self.scan_left_quick, style="Large.TButton")
        self.btn_scan1_quick.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        self.btn_scan2_full = ttk.Button(
            button_frame, text="Scan Worksheet2 (Full)",
            command=self.scan_right_full, style="Large.TButton")
        self.btn_scan2_full.grid(row=0, column=2, padx=5, pady=5, sticky="w")

        self.btn_scan2_quick = ttk.Button(
            button_frame, text="Scan Worksheet2 (Quick)",
            command=self.scan_right_quick, style="Large.TButton")
        self.btn_scan2_quick.grid(row=0, column=3, padx=5, pady=5, sticky="w")

        separator = ttk.Separator(button_frame, orient=tk.VERTICAL)
        separator.grid(row=0, column=4, sticky='ns', padx=10, pady=5)

        self.btn_sync_1_to_2 = ttk.Button(
            button_frame, text="Sync 1 -> 2",
            command=self.sync_1_to_2, style="Large.TButton")
        self.btn_sync_1_to_2.grid(row=0, column=5, padx=5, pady=5, sticky="w")

        self.btn_sync_2_to_1 = ttk.Button(
            button_frame, text="Sync 2 -> 1",
            command=self.sync_2_to_1, style="Large.TButton")
        self.btn_sync_2_to_1.grid(row=0, column=6, padx=5, pady=5, sticky="w")

        self.content_frame = ttk.Frame(self.root, padding="5")
        self.content_frame.grid(row=1, column=0, sticky="nsew")
        self.content_frame.columnconfigure(0, weight=1)
        self.content_frame.columnconfigure(1, weight=1)
        self.content_frame.rowconfigure(0, weight=1)

    def _get_left_pane(self):
        if not self.left_pane:
            left_frame = ttk.Frame(self.content_frame, padding=5)
            left_frame.grid(row=0, column=0, sticky="nsew")
            self.left_pane = WorksheetPane(left_frame, self.root, "Worksheet1")
            self.left_pane.setup_ui()
        return self.left_pane

    def _get_right_pane(self):
        if not self.right_pane:
            self.root.geometry("1800x1100") 
            right_frame = ttk.Frame(self.content_frame, padding=5)
            right_frame.grid(row=0, column=1, sticky="nsew")
            self.right_pane = WorksheetPane(right_frame, self.root, "Worksheet2")
            self.right_pane.setup_ui()
        return self.right_pane

    def scan_left_full(self):
        pane = self._get_left_pane()
        pane.refresh_data(self.btn_scan1_full, scan_mode='full')

    def scan_left_quick(self):
        pane = self._get_left_pane()
        pane.refresh_data(self.btn_scan1_quick, scan_mode='quick')

    def scan_right_full(self):
        pane = self._get_right_pane()
        pane.refresh_data(self.btn_scan2_full, scan_mode='full')

    def scan_right_quick(self):
        pane = self._get_right_pane()
        pane.refresh_data(self.btn_scan2_quick, scan_mode='quick')

    def sync_formulas(self, source_pane, target_pane, source_name, target_name):
        root = self.root
        original_topmost = root.attributes("-topmost")
        root.attributes("-topmost", True)

        target_pane.progress_container_frame.grid(row=2, column=0, columnspan=6, sticky=tk.EW, pady=5, padx=5)
        target_pane.progress_label.config(text="Starting sync...")
        target_pane.progress_bar.config(value=0)
        root.update_idletasks()

        if not source_pane or not target_pane or not source_pane.all_formulas:
            messagebox.showerror("Error", f"Source ({source_name}) not scanned or no data.")
            target_pane.progress_container_frame.grid_forget()
            root.attributes("-topmost", original_topmost)
            return

        if not target_pane.worksheet:
            messagebox.showerror("Error", f"Target ({target_name}) not connected to a valid Excel worksheet.")
            target_pane.progress_container_frame.grid_forget()
            root.attributes("-topmost", original_topmost)
            return

        source_formulas = {
            data[1]: data[2] for data in source_pane.all_formulas
        }

        confirmation = messagebox.askyesno(
            "Confirm Sync",
            f"You are about to sync {len(source_formulas)} formulas from {source_name} to {target_name}.\n\n"
            f"This will overwrite corresponding cells in '{target_pane.worksheet.Name}'.\n"
            f"This operation CANNOT be undone.\n\nAre you sure you want to proceed?"
        )

        if not confirmation:
            messagebox.showinfo("Cancelled", "Sync operation cancelled.")
            target_pane.progress_container_frame.grid_forget()
            root.attributes("-topmost", original_topmost)
            return

        target_pane.activate_excel_window()
        target_ws = target_pane.worksheet
        target_ws.Activate()

        total_formulas = len(source_formulas)
        updated_count = 0
        error_count = 0
        errors = []

        target_pane.progress_bar.config(maximum=total_formulas, value=0)
        target_pane.progress_label.config(text="Starting sync...")
        self.root.update_idletasks()

        for i, (address, formula) in enumerate(source_formulas.items()):
            try:
                target_ws.Range(address).Formula = formula
                updated_count += 1
            except Exception as e:
                error_count += 1
                errors.append(f"  - {address}: {e}")

            target_pane.progress_bar.config(value=i + 1)
            target_pane.progress_label.config(text=f"Syncing: {i + 1}/{total_formulas} formulas")
            self.root.update_idletasks()

        summary_message = f"Sync Complete.\n\nSuccessfully updated: {updated_count}\nFailed to update: {error_count}"
        if errors:
            summary_message += "\n\nErrors:\n" + "\n".join(errors[:5])
        messagebox.showinfo("Sync Result", summary_message + "\n\nPlease re-scan the target worksheet to view changes.")

        target_pane.progress_bar.config(value=0)
        target_pane.progress_label.config(text="")
        target_pane.progress_container_frame.grid_forget()
        self.root.update_idletasks()
        root.attributes("-topmost", original_topmost)

    def sync_1_to_2(self):
        self.sync_formulas(self.left_pane, self.right_pane, "Worksheet1", "Worksheet2")
        self._get_right_pane().refresh_data(btn=None, scan_mode='quick')

    def sync_2_to_1(self):
        self.sync_formulas(self.right_pane, self.left_pane, "Worksheet2", "Worksheet1")
        self._get_left_pane().refresh_data(btn=None, scan_mode='quick')