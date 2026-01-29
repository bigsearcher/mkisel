#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Minimal tkinter test for PyInstaller onedir.
If this shows a window on the target PC, the problem is in TallyConverter code/deps.
If it does not, the problem is tkinter/Tcl on that PC (driver, AV, policy).
"""
import tkinter as tk

root = tk.Tk()
root.title("Tkinter test - OK")
root.geometry("300x120")
tk.Label(root, text="If you see this window,\ntkinter works on this PC.", font=("Segoe UI", 10)).pack(pady=20)
root.mainloop()
