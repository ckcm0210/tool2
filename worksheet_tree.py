# -*- coding: utf-8 -*-
"""
Created on Wed Jun 25 09:11:26 2025

@author: kccheng
"""

import re
import os
import tkinter as tk
from tkinter import messagebox
import win32com.client
from openpyxl.utils import column_index_from_string

def apply_filter(self, event=None):
    self.result_tree.delete(*self.result_tree.get_children())
    self.cell_addresses.clear()
    address_filter_str = self.filter_entries['address'].get().strip()
    parsed_address_filters = []
    if address_filter_str and address_filter_str != self.placeholder_text:
        address_tokens = [token.strip() for token in address_filter_str.split(',') if token.strip()]
        if address_tokens:
            try:
                for token in address_tokens:
                    parsed_address_filters.append(self.parse_excel_address(token))
            except Exception as e:
                from tkinter import messagebox
                messagebox.showerror("Invalid Excel Address", str(e))
                return
    other_filters = {
        'type': (self.show_formula.get(), self.show_local_link.get(), self.show_external_link.get()),
        'formula': self.filter_entries['formula'].get().lower(),
        'result': self.filter_entries['result'].get().lower(),
        'display_value': self.filter_entries['display_value'].get().lower()
    }
    filtered_formulas = []
    for formula_data in self.all_formulas:
        if len(formula_data) < 5: continue
        formula_type, address, formula_content, result_val, display_val = formula_data
        type_map = {'formula': other_filters['type'][0], 'local link': other_filters['type'][1], 'external link': other_filters['type'][2]}
        if not type_map.get(formula_type, True): continue
        if other_filters['formula'] and other_filters['formula'] not in str(formula_content).lower(): continue
        if other_filters['result'] and other_filters['result'] not in str(result_val).lower(): continue
        if other_filters['display_value'] and other_filters['display_value'] not in str(display_val).lower(): continue
        if parsed_address_filters:
            addr_upper = address.replace("$", "").upper()
            current_cell_match = re.match(r"([A-Z]+)([0-9]+)", addr_upper)
            if not current_cell_match: continue
            cell_col_str, cell_row_str = current_cell_match.groups()
            cell_col_idx = column_index_from_string(cell_col_str)
            cell_row_idx = int(cell_row_str)
            is_match = False
            for f_type, f_val in parsed_address_filters:
                if f_type == 'cell' and addr_upper == f_val:
                    is_match = True; break
                elif f_type == 'row_range':
                    start_r, end_r = map(int, f_val.split(':'))
                    if start_r <= cell_row_idx <= end_r:
                        is_match = True; break
                elif f_type == 'col_range':
                    start_c, end_c = f_val.split(':')
                    if column_index_from_string(start_c) <= cell_col_idx <= column_index_from_string(end_c):
                        is_match = True; break
                elif f_type == 'range':
                    start_cell, end_cell = f_val.split(':')
                    sc_str, sr_str = re.match(r"([A-Z]+)([0-9]+)", start_cell).groups()
                    ec_str, er_str = re.match(r"([A-Z]+)([0-9]+)", end_cell).groups()
                    if (column_index_from_string(sc_str) <= cell_col_idx <= column_index_from_string(ec_str) and
                        int(sr_str) <= cell_row_idx <= int(er_str)):
                        is_match = True; break
            if not is_match: continue
        filtered_formulas.append(formula_data)
    if self.current_sort_column:
        col_index = self.tree_columns.index(self.current_sort_column)
        sort_dir = self.sort_directions[self.current_sort_column]
        filtered_formulas.sort(key=lambda x: str(x[col_index]), reverse=(sort_dir == -1))
    count = len(filtered_formulas)
    self.formula_list_label.config(text=f"Formula List ({count} records):")
    for i, data in enumerate(filtered_formulas):
        tag = "evenrow" if i % 2 == 0 else "oddrow"
        item_id = self.result_tree.insert("", "end", values=data, tags=(tag,))
        address_index = self.tree_columns.index("address")
        if address_index < len(data):
            self.cell_addresses[item_id] = data[address_index]

def sort_column(self, col_id):
    self.current_sort_column = col_id
    self.sort_directions[col_id] *= -1
    self.apply_filter()
    for column in self.tree_columns:
        original_text = self.result_tree.heading(column, "text").split(' ')[0]
        self.result_tree.heading(column, text=original_text, image='')
    current_direction = " \u2191" if self.sort_directions[col_id] == 1 else " \u2193"
    current_text = self.result_tree.heading(col_id, "text").split(' ')[0]
    self.result_tree.heading(col_id, text=current_text + current_direction)

def go_to_reference(self, workbook_path, sheet_name, cell_address):
    try:
        try:
            self.xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            try:
                self.xl = win32com.client.Dispatch("Excel.Application")
                self.xl.Visible = True
            except Exception as e:
                messagebox.showerror("Excel Error", f"Could not start or connect to Excel.\nError: {e}")
                return

        target_workbook = None
        normalized_workbook_path = os.path.normpath(workbook_path) if workbook_path else None

        if normalized_workbook_path:
            for wb in self.xl.Workbooks:
                if os.path.normpath(wb.FullName) == normalized_workbook_path:
                    target_workbook = wb
                    break
            if not target_workbook:
                if os.path.exists(normalized_workbook_path):
                    target_workbook = self.xl.Workbooks.Open(normalized_workbook_path)
                else:
                    messagebox.showerror("File Not Found", f"The workbook path was not found:\n{normalized_workbook_path}")
                    return
        else:
            target_workbook = self.workbook

        if not target_workbook:
            messagebox.showerror("Error", "Could not access the target workbook.")
            return

        target_worksheet = None
        try:
            target_worksheet = target_workbook.Worksheets(sheet_name)
        except Exception:
            messagebox.showerror("Worksheet Not Found", f"Could not find worksheet '{sheet_name}' in workbook '{os.path.basename(target_workbook.FullName)}'.")
            return

        self.activate_excel_window()
        target_workbook.Activate()
        target_worksheet.Activate()
        target_worksheet.Range(cell_address).Select()

    except Exception as e:
        messagebox.showerror("Navigation Error", f"Could not navigate to cell '{cell_address}'.\nError: {e}")


def on_select(self, event):
    selected_item = self.result_tree.selection()
    if not selected_item:
        self.detail_text.delete(1.0, 'end')
        return
    item_id = selected_item[0]
    values = self.result_tree.item(item_id, "values")
    if len(values) < 5:
        self.detail_text.delete(1.0, 'end')
        self.detail_text.insert(1.0, "Selected item has incomplete data.")
        return
    formula_type, cell_address, formula, result, display_value = values
    self.detail_text.delete(1.0, 'end')
    self.detail_text.insert('end', "Type: ", "label")
    self.detail_text.insert('end', f"{formula_type} / ", "value")
    self.detail_text.insert('end', "Cell Address: ", "label")
    self.detail_text.insert('end', f"{cell_address}\n", "value")
    self.detail_text.insert('end', "Calculated Result: ", "label")
    self.detail_text.insert('end', f"{result} / ", "result_value")
    self.detail_text.insert('end', "Displayed Value: ", "label")
    self.detail_text.insert('end', f"{display_value}\n\n", "value")
    self.detail_text.insert('end', "Formula Content:\n", "label")
    self.detail_text.insert('end', f"{formula}\n\n", "formula_content")
    if self.xl and self.worksheet:
        referenced_values = self.get_referenced_cell_values(
            formula,
            self.worksheet,
            self.workbook.FullName,
            self._read_external_cell_value,
            self._find_matching_sheet
        )
        if referenced_values:
            self.detail_text.insert('end', "Referenced Cell Values (Non-Range):\n", "label")
            for ref_addr, ref_val in referenced_values.items():
                display_text = ref_addr
                if '|' in ref_addr:
                    _, display_text = ref_addr.split('|', 1)

                self.detail_text.insert('end', f"  {display_text}: {ref_val}  ", "referenced_value")

                workbook_path = None
                sheet_name = None
                cell_address_to_go = None
                
                try:
                    if '|' in ref_addr:
                        full_path, display_ref = ref_addr.split('|', 1)
                        workbook_path = full_path
                        
                        if ']' in display_ref and '!' in display_ref:
                            sheet_and_cell = display_ref.split(']', 1)[1]
                            parts = sheet_and_cell.rsplit('!', 1)
                            sheet_name = parts[0].strip("'")
                            cell_address_to_go = parts[1]
                    else:
                        workbook_path = self.workbook.FullName
                        if '!' in ref_addr:
                            parts = ref_addr.rsplit('!', 1)
                            sheet_name = parts[0]
                            cell_address_to_go = parts[1]

                    if workbook_path and sheet_name and cell_address_to_go:
                        def build_handler(wp, sn, ca):
                            def handler():
                                self.go_to_reference(wp, sn, ca)
                            return handler
                        
                        btn = tk.Button(self.detail_text, text="Go to Reference", font=("Arial", 7), cursor="hand2", command=build_handler(workbook_path, sheet_name, cell_address_to_go))
                        self.detail_text.window_create('end', window=btn)

                except Exception as e:
                    print(f"INFO: Could not create navigation button for '{ref_addr}': {e}")

                self.detail_text.insert('end', "\n")
        else:
            self.detail_text.insert('end', "Referenced Cell Values (Non-Range):\n", "label")
            self.detail_text.insert('end', "  No individual cell references found or accessible.\n", "info_text")
    else:
        self.detail_text.insert('end', "Excel connection not active to retrieve referenced values.\n", "info_text")
        
def on_double_click(self, event):
    selected_item = self.result_tree.selection()
    if not selected_item:
        return
    item_id = selected_item[0]
    cell_address = self.cell_addresses.get(item_id)
    if cell_address:
        try:
            try:
                self.xl = win32com.client.GetActiveObject("Excel.Application")
            except Exception:
                self.xl = None
            if not self.xl:
                try:
                    self.xl = win32com.client.Dispatch("Excel.Application")
                    self.xl.Visible = True
                    if self.last_workbook_path and os.path.exists(self.last_workbook_path):
                        self.workbook = self.xl.Workbooks.Open(self.last_workbook_path)
                    else:
                        messagebox.showwarning("File Not Found", "The last scanned Excel file path is not valid or found. Please open Excel manually.")
                        return
                except Exception as e:
                    messagebox.showerror("Excel Launch Error", f"Could not launch Excel or open the workbook.\nError: {e}")
                    return
            if self.workbook and self.workbook.FullName != self.last_workbook_path:
                response = messagebox.askyesno("Workbook Mismatch", 
                            f"The current active workbook is '{self.workbook.Name}', but the record is from '{os.path.basename(self.last_workbook_path)}'.\nDo you want to open the correct workbook?")
                if response:
                    try:
                        self.workbook.Close(SaveChanges=False) 
                    except Exception as close_e:
                        pass
                    try:
                        self.workbook = self.xl.Workbooks.Open(self.last_workbook_path)
                    except Exception as open_e:
                        messagebox.showerror("Workbook Open Error", f"Could not open workbook '{os.path.basename(self.last_workbook_path)}'.\nError: {open_e}")
                        return
                else:
                    if not self.workbook:
                        messagebox.showinfo("Operation Cancelled", "Operation cancelled. Please select the correct workbook manually in Excel.")
                    return
            if self.last_worksheet_name and self.workbook:
                try:
                    self.worksheet = self.workbook.Worksheets(self.last_worksheet_name)
                except Exception:
                    self.worksheet = self.workbook.ActiveSheet
                    messagebox.showwarning("Worksheet Not Found", f"Worksheet '{self.last_worksheet_name}' not found in '{self.workbook.Name}'. Activating current sheet.")
            elif self.workbook:
                self.worksheet = self.workbook.ActiveSheet
            else:
                messagebox.showerror("Error", "No active workbook to select cell in.")
                return
            self.workbook.Activate()
            self.worksheet.Activate()
            self.worksheet.Range(cell_address).Select()
            self.activate_excel_window()
        except Exception as e:
            messagebox.showerror("Excel Selection Error", f"Could not select cell {cell_address} in Excel. Please ensure the workbook and worksheet are still valid.\nError: {e}")