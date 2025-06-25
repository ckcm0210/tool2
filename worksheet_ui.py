# -*- coding: utf-8 -*-
"""
Created on Wed Jun 25 09:11:40 2025

@author: kccheng
"""

import tkinter as tk
from tkinter import font
from tkinter import ttk

def setup_ui(self):
    self.parent.columnconfigure(0, weight=1)
    self.parent.rowconfigure(5, weight=1)
    self.parent.rowconfigure(7, weight=2)
    default_content_font = ("Consolas", 10)
    main_label_font = ("Arial", 12, "bold")
    filter_label_font = ("Arial", 9, "bold")
    style = ttk.Style()
    style.configure("Treeview", font=default_content_font)
    style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
    style.configure("evenrow", background="#F0F0F0")
    style.configure("oddrow", background="#FFFFFF")
    info_frame = ttk.Frame(self.parent)
    info_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
    info_frame.columnconfigure(1, weight=1)
    ttk.Label(info_frame, text="File Path:", font=main_label_font).grid(row=0, column=0, sticky=tk.W)
    self.path_label = ttk.Label(info_frame, text="Not Connected", foreground="red", wraplength=400)
    self.path_label.grid(row=0, column=1, sticky=tk.W)
    ttk.Label(info_frame, text="File Name:", font=main_label_font).grid(row=1, column=0, sticky=tk.W)
    self.file_label = ttk.Label(info_frame, text="Not Connected", foreground="red")
    self.file_label.grid(row=1, column=1, sticky=tk.W)
    ttk.Label(info_frame, text="Worksheet:", font=main_label_font).grid(row=2, column=0, sticky=tk.W)
    self.sheet_label = ttk.Label(info_frame, text="Not Connected", foreground="red")
    self.sheet_label.grid(row=2, column=1, sticky=tk.W)
    ttk.Label(info_frame, text="Data Range:", font=main_label_font).grid(row=3, column=0, sticky=tk.W)
    self.range_label = ttk.Label(info_frame, text="Not Connected", foreground="red")
    self.range_label.grid(row=3, column=1, sticky=tk.W)
    self.progress_frame = ttk.Frame(self.parent)
    self.progress_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
    self.progress_frame.columnconfigure(0, weight=1)
    self.progress_label = ttk.Label(self.progress_frame, text="")
    self.progress_label.pack(fill=tk.X)
    self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate')
    self.progress_bar.pack(fill=tk.X, pady=(2, 0))
    filter_main_frame = ttk.LabelFrame(self.parent, text="Filters", borderwidth=2, relief=tk.GROOVE, padding=10)
    filter_main_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
    filter_main_frame.columnconfigure(0, weight=1)
    filter_checkbox_frame = ttk.Frame(filter_main_frame)
    filter_checkbox_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))
    ttk.Label(filter_checkbox_frame, text="Type:", font=filter_label_font).pack(side=tk.LEFT, padx=(0, 5))
    ttk.Checkbutton(filter_checkbox_frame, text="Formula", variable=self.show_formula, command=self.apply_filter).pack(side=tk.LEFT, padx=5)
    ttk.Checkbutton(filter_checkbox_frame, text="Local Link", variable=self.show_local_link, command=self.apply_filter).pack(side=tk.LEFT, padx=5)
    ttk.Checkbutton(filter_checkbox_frame, text="External Link", variable=self.show_external_link, command=self.apply_filter).pack(side=tk.LEFT, padx=5)
    openpyxl_check = ttk.Checkbutton(filter_checkbox_frame, text="Enable Non-GUI File Reading for Cell Results", variable=self.use_openpyxl, command=lambda: self.on_select(event=None))
    openpyxl_check.pack(side=tk.LEFT, padx=15)
    filter_entry_frame = ttk.Frame(filter_main_frame)
    filter_entry_frame.pack(side=tk.TOP, fill=tk.X)
    filter_entry_frame.columnconfigure(1, weight=1)
    filter_entry_frame.columnconfigure(2, weight=0)
    self.tree_columns = ("type", "address", "formula", "result", "display_value")
    self.columns_with_entries = ("address", "formula", "result", "display_value")
    self.filter_entries = {}
    column_display_names = {"address": "Address", "formula": "Formula", "result": "Result", "display_value": "Display Value"}
    row_idx = 0
    for col_id in self.columns_with_entries:
        ttk.Label(filter_entry_frame, text=f"{column_display_names[col_id]}:", font=filter_label_font).grid(row=row_idx, column=0, sticky=tk.W, padx=(5,0), pady=2)
        entry = ttk.Entry(filter_entry_frame, font=("Consolas", 10))
        entry.grid(row=row_idx, column=1, sticky=(tk.W, tk.E), padx=(0,5), pady=2)
        entry.bind("<Return>", self.apply_filter)
        self.filter_entries[col_id] = entry
        btn_text = "⏎"
        btn = ttk.Button(filter_entry_frame, text=btn_text, command=self.apply_filter, width=3)
        btn.grid(row=row_idx, column=2, padx=(0,5), pady=2)
        if col_id == 'address':
            self.default_fg_color = entry.cget("foreground")
            base_font = font.Font(font=entry.cget("font"))
            self.default_font = base_font
            self.placeholder_font = font.Font(family=base_font.cget("family"), size=base_font.cget("size"), slant="italic")
            entry.bind("<FocusIn>", self._on_focus_in)
            entry.bind("<FocusOut>", self._on_focus_out)
        row_idx += 1
    self._set_placeholder()
    summary_frame = ttk.Frame(self.parent)
    summary_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=5, padx=10)
    self.summarize_button = ttk.Button(summary_frame, text="Summarize External Links", command=self.summarize_external_links)
    self.summarize_button.pack(side=tk.LEFT, padx=(0, 5))
    self.export_button = ttk.Button(summary_frame, text="Export and Open List", command=self.export_formulas_to_excel)
    self.export_button.pack(side=tk.LEFT, padx=5)
    self.import_button = ttk.Button(summary_frame, text="Import and Update Formulas", command=self.import_and_update_formulas)
    self.import_button.pack(side=tk.LEFT, padx=5)
    self.reconnect_button = ttk.Button(summary_frame, text="Reconnect", command=self.reconnect_to_excel)
    self.reconnect_button.pack(side=tk.LEFT, padx=5)
    self.formula_list_label = ttk.Label(self.parent, text="Formula List:", font=main_label_font)
    self.formula_list_label.grid(row=4, column=0, sticky=tk.W, pady=(10, 0))
    tree_frame = ttk.Frame(self.parent)
    tree_frame.grid(row=5, column=0, sticky="nsew")
    tree_frame.columnconfigure(0, weight=1)
    tree_frame.rowconfigure(0, weight=1)
    self.result_tree = ttk.Treeview(tree_frame, columns=self.tree_columns, show="headings", height=15)
    headings = {"type": "Type", "address": "Address", "formula": "Formula Content", "result": "Result", "display_value": "Display Value"}
    widths = {"type": 70, "address": 70, "formula": 400, "result": 120, "display_value": 120}
    for col_id, text in headings.items():
        self.result_tree.heading(col_id, text=text, command=lambda c=col_id: self.sort_column(c))
        self.result_tree.column(col_id, width=widths[col_id], minwidth=60)
    self.result_tree.grid(row=0, column=0, sticky="nsew")
    scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.result_tree.yview)
    scrollbar.grid(row=0, column=1, sticky="ns")
    self.result_tree.configure(yscrollcommand=scrollbar.set)
    self.result_tree.bind("<Double-Button-1>", self.on_double_click)
    self.result_tree.bind("<<TreeviewSelect>>", self.on_select)
    ttk.Label(self.parent, text="Details:", font=main_label_font).grid(row=6, column=0, sticky=tk.W, pady=(10, 0))
    detail_frame = ttk.Frame(self.parent)
    detail_frame.grid(row=7, column=0, sticky="nsew")
    detail_frame.columnconfigure(0, weight=1)
    detail_frame.rowconfigure(0, weight=1)
    self.detail_text = tk.Text(detail_frame, height=20, wrap=tk.WORD, font=default_content_font)
    self.detail_text.grid(row=0, column=0, sticky="nsew")
    detail_scrollbar = ttk.Scrollbar(detail_frame, command=self.detail_text.yview)
    detail_scrollbar.grid(row=0, column=1, sticky="ns")
    self.detail_text.configure(yscrollcommand=detail_scrollbar.set)
    self.detail_text.tag_configure("label", font=("Consolas", 10, "bold"), foreground="navy")
    self.detail_text.tag_configure("value", font=("Consolas", 10), foreground="black")
    self.detail_text.tag_configure("formula_content", font=("Consolas", 10, "italic"), foreground="darkgreen")
    self.detail_text.tag_configure("result_value", font=("Consolas", 10), foreground="darkblue")
    self.detail_text.tag_configure("referenced_value", font=("Consolas", 10), foreground="purple")
    self.detail_text.tag_configure("info_text", font=("Consolas", 10, "italic"), foreground="grey")
    self.ui_initialized = True

def _set_placeholder(self):
    entry = self.filter_entries.get('address')
    if entry:
        entry.delete(0, tk.END)
        entry.insert(0, self.placeholder_text)
        entry.config(foreground=self.placeholder_color, font=self.placeholder_font)
        entry.icursor(0)

def _on_focus_in(self, event):
    entry = event.widget
    if entry.get() == self.placeholder_text:
        entry.delete(0, tk.END)
        entry.config(foreground=self.default_fg_color, font=self.default_font)

def _on_mouse_click(self, event):
    entry = event.widget
    if entry.get() == self.placeholder_text:
        entry.delete(0, tk.END)
        entry.config(foreground=self.default_fg_color, font=self.default_font)
        return "break"

def _on_focus_out(self, event):
    entry = event.widget
    if not entry.get():
        entry.insert(0, self.placeholder_text)
        entry.config(foreground=self.placeholder_color, font=self.placeholder_font)