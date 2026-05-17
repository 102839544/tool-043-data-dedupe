#!/usr/bin/env python3
"""
数据去重工具 - CSV/Excel去重
"""
import sys, os, tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

class App:
    def __init__(self, root):
        self.root = root
        root.title("数据去重工具 v1.0")
        root.geometry("550x400")
        self.file = None
        self.build_ui()
    
    def build_ui(self):
        f = tk.Frame(self.root, bg="#f57c00", height=50)
        f.pack(fill="x")
        tk.Label(f, text="🧹 数据去重工具", font=("Arial",14,"bold"),
                 fg="white", bg="#f57c00").pack(pady=12)
        
        main = tk.Frame(self.root, padx=15, pady=10)
        main.pack(fill="both", expand=True)
        
        tk.Button(main, text="选择CSV/Excel文件", command=self.select_file,
                  bg="#f57c00", fg="white", padx=15).pack(pady=10)
        
        self.file_label = tk.Label(main, text="未选择文件", fg="gray")
        self.file_label.pack()
        
        of = tk.Frame(main)
        of.pack(pady=15)
        tk.Label(of, text="去重列（留空则整行去重）：").pack()
        self.col_entry = tk.Entry(of, width=30)
        self.col_entry.pack(pady=5)
        
        tk.Button(main, text="去重并保存", command=self.dedupe,
                  bg="#4caf50", fg="white", font=("Arial",10,"bold"),
                  padx=20).pack(pady=15)
        
        self.status = tk.Label(main, text="", fg="gray")
        self.status.pack()
    
    def select_file(self):
        f = filedialog.askopenfilename(title="选择文件",
             filetypes=[("数据文件","*.csv *.xlsx *.xls")])
        if f:
            self.file = f
            self.file_label.config(text=Path(f).name)
    
    def dedupe(self):
        if not self.file:
            messagebox.showwarning("提示", "请先选择文件")
            return
        if not HAS_PANDAS:
            messagebox.showerror("缺少依赖", "请运行：pip install pandas openpyxl")
            return
        
        try:
            if self.file.endswith(".csv"):
                df = pd.read_csv(self.file)
            else:
                df = pd.read_excel(self.file)
            
            orig_len = len(df)
            col = self.col_entry.get().strip()
            
            if col:
                df = df.drop_duplicates(subset=[col], keep="first")
            else:
                df = df.drop_duplicates(keep="first")
            
            out = filedialog.asksaveasfilename(title="保存",
                 defaultextension=".xlsx" if self.file.endswith((".xlsx",".xls")) else ".csv",
                 filetypes=[("Excel","*.xlsx"),("CSV","*.csv")])
            if not out: return
            
            if out.endswith(".csv"):
                df.to_csv(out, index=False, encoding="utf-8-sig")
            else:
                df.to_excel(out, index=False)
            
            self.status.config(text=f"✅ 去重完成：{orig_len} → {len(df)} 行")
            messagebox.showinfo("完成", f"去重成功！\n原始：{orig_len} 行\n去重后：{len(df)} 行")
        except Exception as e:
            messagebox.showerror("错误", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
