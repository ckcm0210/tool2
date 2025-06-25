# -*- coding: utf-8 -*-
"""
Created on Wed Jun 25 09:11:09 2025

@author: kccheng
"""

import os
import time
import win32com.client
from tkinter import messagebox
import psutil
import win32gui
import win32process
import win32con
from excel_utils import classify_formula_type

def refresh_data(self, btn, scan_mode='full'):
    if not self.ui_initialized:
        self.setup_ui()
    if btn is not None:
        btn.config(state='disabled')
    self.progress_bar['value'] = 0
    self.progress_label.config(text="Connecting to active Excel...")
    self.root.update_idletasks()
    excel_pids = set(
        p.pid for p in psutil.process_iter(['name'])
        if p.info['name'] and p.info['name'].lower() == "excel.exe"
    )
    pids_with_window = set()
    def callback(hwnd, extra):
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid in excel_pids and win32gui.IsWindowVisible(hwnd):
                pids_with_window.add(pid)
        except Exception:
            pass
        return True
    win32gui.EnumWindows(callback, None)
    processes_with_ui = [
        p for p in psutil.process_iter(['pid', 'name'])
        if p.info['pid'] in pids_with_window
    ]
    processes_without_ui = [
        p for p in psutil.process_iter(['pid', 'name'])
        if p.info['name'] and p.info['name'].lower() == "excel.exe"
        and p.info['pid'] not in pids_with_window
    ]
    if len(processes_without_ui) > 0:
        answer = messagebox.askyesno(
            "Excel 殘留進程偵測",
            "系統偵測到有 {0} 個冇 UI（殘留）嘅 EXCEL.EXE 進程。\n\n"
            "這些通常係因為 Excel 曾經 crash、強制結束，或者外部自動化失敗殘留。\n"
            "殘留進程會阻礙自動連接 Excel 或導致工具異常。\n\n"
            "你是否要立即結束全部殘留進程？（此操作唔會影響你有 UI 嘅 Excel 視窗）"
            .format(len(processes_without_ui))
        )
        if answer:
            for p in processes_without_ui:
                try:
                    p.kill()
                except Exception:
                    pass
            messagebox.showinfo("處理完成", "所有殘留進程已經關閉。")
    try:
        try:
            self.xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception as e:
            try:
                self.xl = win32com.client.Dispatch("Excel.Application")
                self.xl.Visible = True
                messagebox.showinfo("Info", "No existing Excel instance detected. A new Excel instance has been started automatically. Please open your file manually and press Scan again.")
                if btn is not None:
                    btn.config(state='normal')
                self.progress_bar['value'] = 0
                self.progress_label.config(text="Please open your Excel file and try again.")
                return
            except Exception as e2:
                messagebox.showerror("Connection Error", f"Could not find an existing Excel instance or start a new one.\nPlease check for any leftover EXCEL.EXE processes and verify your permission.\n\nError: {e2}")
                if btn is not None:
                    btn.config(state='normal')
                self.progress_bar['value'] = 0
                self.progress_label.config(text="Connection Failed.")
                return
        try:
            self.workbook = self.xl.ActiveWorkbook
            self.worksheet = self.xl.ActiveSheet
        except Exception as e:
            messagebox.showerror("Connection Error", "Excel is open, but there is no active workbook or worksheet.\nPlease open a file and worksheet in Excel and try again.")
            if btn is not None:
                btn.config(state='normal')
            self.progress_bar['value'] = 0
            self.progress_label.config(text="No active workbook.")
            return
        print(f"INFO: Connected to Excel: {self.workbook.Name} - {self.worksheet.Name}")
        self.last_workbook_path = self.workbook.FullName
        self.last_worksheet_name = self.worksheet.Name
        self.progress_bar['value'] = 10
        self.progress_label.config(text="Reading workbook information...")
        self.root.update_idletasks()
        file_path = self.workbook.FullName
        display_path = os.path.dirname(file_path)
        max_path_display_length = 60
        if len(display_path) > max_path_display_length:
            truncated_path = "..." + display_path[-(max_path_display_length-3):]
        else:
            truncated_path = display_path
        self.file_label.config(text=os.path.basename(file_path), foreground="black")
        self.path_label.config(text=truncated_path, foreground="black")
        self.sheet_label.config(text=self.worksheet.Name, foreground="black")
        current_scan_range = self.worksheet.UsedRange
        current_scan_range_str = current_scan_range.Address.replace('$', '')
        self.range_label.config(text=f"Scanning: UsedRange ({current_scan_range_str})", foreground="black")
        self.all_formulas.clear()
        formula_cells_found = 0
        self.progress_bar['value'] = 30
        self.progress_label.config(text=f"Searching for formulas in Excel in range {current_scan_range_str} (this may take a moment)...")
        self.root.update_idletasks()
        start_time = time.time()
        try:
            xlCellTypeFormulas = -4123
            formula_range = current_scan_range.SpecialCells(xlCellTypeFormulas)
            areas_to_process = []
            if formula_range.Areas.Count > 1:
                for area in formula_range.Areas:
                    areas_to_process.append(area)
            else:
                areas_to_process.append(formula_range)
            total_cells_to_process = sum(area.Cells.Count for area in areas_to_process)
            current_cell_count = 0
            for area in areas_to_process:
                for cell in area.Cells:
                    current_cell_count += 1
                    formula_cells_found += 1
                    formula = ""
                    formula_type = "unknown"
                    cell_value = None
                    display_val = "Error"
                    cell_text = "Error"
                    cell_address = ""
                    try:
                        formula = cell.Formula
                        formula_type = classify_formula_type(formula)
                        cell_value = cell.Value
                        display_val = str(cell_value)[:50] if cell_value is not None else "No Value"
                        if scan_mode == 'quick':
                            cell_text = "N/A (Quick Scan)"
                        else:
                            cell_text = str(cell.Text).strip()
                        cell_address = cell.Address.replace('$', '')
                        self.all_formulas.append((formula_type, cell_address, formula, display_val, cell_text))
                    except Exception as cell_processing_e:
                        print(f"WARNING: Error processing cell at {cell_address if cell_address else 'unknown'}: {cell_processing_e}")
                        self.all_formulas.append((formula_type, cell_address if cell_address else "ERROR_ADDR", str(formula), str(display_val), f"ERROR: {cell_processing_e}"))
                    if current_cell_count % 100 == 0 or current_cell_count == total_cells_to_process:
                        progress = 30 + (current_cell_count / total_cells_to_process) * 60
                        self.progress_bar['value'] = min(int(progress), 90)
                        self.progress_label.config(text=f"Found {formula_cells_found} formulas. Processing {current_cell_count}/{total_cells_to_process} cells...")
                        self.root.update_idletasks()
        except Exception as e:
            no_formula_error = (
                "(-2146827284, 'OLE error.', None, None)" in str(e)
                or "0x800A03EC" in str(e)
                or '找不到所要找的儲存格' in str(e)
                or 'Unable to get the' in str(e)
            )
            if no_formula_error:
                print(f"INFO: No formulas found in the current worksheet within range {current_scan_range_str}.")
                self.progress_label.config(text=f"No formulas found in this worksheet's selected range ({current_scan_range_str}).")
                self.progress_bar['value'] = 100
                self.root.update_idletasks()
                if self.formula_list_label:
                    self.formula_list_label.config(text="Formula List (No Formula Found)")
                self.apply_filter()
                if btn is not None:
                    btn.config(state='normal')
                return
            else:
                messagebox.showerror("Scan Error", f"An error occurred while scanning formulas: {e}")
                print(f"ERROR: Error scanning for formulas: {e}")
            if btn is not None:
                btn.config(state='normal')
            return
        end_time = time.time()
        time_taken = end_time - start_time
        self.progress_bar['value'] = 90
        self.progress_label.config(text=f"Found {len(self.all_formulas)} formulas. Loading... (Scan took {time_taken:.2f} seconds)")
        self.root.update_idletasks()
        self.apply_filter()
        self.progress_bar['value'] = 100
        self.progress_label.config(text=f"Completed: Found {len(self.all_formulas)} formulas. (Total scan time: {time_taken:.2f} seconds)")
        if btn is not None:
            btn.config(state='normal')
        if self.formula_list_label:
            if len(self.all_formulas) == 0:
                self.formula_list_label.config(text="Formula List (No Formula Found)")
            else:
                self.formula_list_label.config(text="Formula List:")
    except Exception as e:
        import traceback
        err_detail = traceback.format_exc()
        messagebox.showerror(
            "Connection Error",
            "Could not connect to Excel. Please ensure:\n"
            "1. Excel is open\n"
            "2. There are no leftover EXCEL.EXE processes (check with Task Manager)\n"
            "3. Permissions are consistent (do not run one as admin and the other not)\n\n"
            f"Error: {e}\n\nTraceback:\n{err_detail}"
        )
        if btn is not None:
            btn.config(state='normal')
        self.progress_bar['value'] = 0
        self.progress_label.config(text="Connection Failed.")
        return

def reconnect_to_excel(self):
    if not self.last_workbook_path or not self.last_worksheet_name:
        from tkinter import messagebox
        messagebox.showerror("Cannot Reconnect", "No saved workbook path or worksheet name.\nPlease scan a worksheet first.")
        return
    try:
        self.xl = win32com.client.GetActiveObject("Excel.Application")
        self.xl.Visible = True
    except Exception:
        from tkinter import messagebox
        messagebox.showerror("Connection Error", "Excel is not running. Please open Excel and try again.")
        return
    found_workbook = None
    for wb in self.xl.Workbooks:
        if wb.FullName == self.last_workbook_path:
            found_workbook = wb
            break
    if found_workbook:
        self.workbook = found_workbook
    else:
        try:
            if os.path.exists(self.last_workbook_path):
                self.workbook = self.xl.Workbooks.Open(self.last_workbook_path)
            else:
                from tkinter import messagebox
                messagebox.showerror("File Not Found", f"Saved path file does not exist:\n{self.last_workbook_path}")
                return
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror("Open Error", f"Unable to open workbook '{os.path.basename(self.last_workbook_path)}'.\nError: {e}")
            return
    try:
        self.worksheet = self.workbook.Worksheets(self.last_worksheet_name)
        self.workbook.Activate()
        self.worksheet.Activate()
        self.activate_excel_window()
        file_path = self.workbook.FullName
        display_path = os.path.dirname(file_path)
        max_path_display_length = 60
        if len(display_path) > max_path_display_length:
            truncated_path = "..." + display_path[-(max_path_display_length - 3):]
        else:
            truncated_path = display_path
        self.file_label.config(text=os.path.basename(file_path), foreground="black")
        self.path_label.config(text=truncated_path, foreground="black")
        self.sheet_label.config(text=self.worksheet.Name, foreground="black")
        current_scan_range_str = self.worksheet.UsedRange.Address.replace('$', '')
        self.range_label.config(text=f"UsedRange ({current_scan_range_str})", foreground="black")
        from tkinter import messagebox
        messagebox.showinfo("Connection Successful", f"Successfully reconnected to:\n{self.workbook.Name} - {self.worksheet.Name}")
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror("Connection Error", f"Unable to activate worksheet '{self.last_worksheet_name}'.\nError: {e}")
        self.file_label.config(text="Not Connected", foreground="red")
        self.path_label.config(text="Not Connected", foreground="red")
        self.sheet_label.config(text="Not Connected", foreground="red")
        self.range_label.config(text="Not Connected", foreground="red")

def activate_excel_window(self):
    if not self.xl:
        return
    try:
        self.xl.Visible = True
        excel_hwnd = self.xl.Hwnd
        import win32gui, win32con
        if win32gui.IsIconic(excel_hwnd):
            win32gui.ShowWindow(excel_hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(excel_hwnd)
        self.xl.ActiveWindow.Activate()
    except Exception as e:
        from tkinter import messagebox
        messagebox.showwarning("Activate Excel Window", f"Could not activate Excel window. Please switch to Excel manually. Error: {e}")