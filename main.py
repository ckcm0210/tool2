# -*- coding: utf-8 -*-
"""
Created on Wed Jun 25 09:10:08 2025

@author: kccheng
"""

import tkinter as tk
from formula_comparator import ExcelFormulaComparator

def main():
    root = tk.Tk()
    app = ExcelFormulaComparator(root)
    root.mainloop()

if __name__ == "__main__":
    main()