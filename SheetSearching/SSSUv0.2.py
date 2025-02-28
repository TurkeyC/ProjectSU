import pandas as pd
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk
import sys
import os

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# def search_excel(file_path, search_term, case_sensitive=False):
#     try:
#         # 读取 Excel 文件，不将任何值视为缺失值
#         xls = pd.ExcelFile(file_path)
#         results = {}
#
#         # 遍历所有 sheet
#         for sheet_name in xls.sheet_names:
#             df = pd.read_excel(xls, sheet_name=sheet_name, keep_default_na=False, na_values=[])
#
#             # 根据大小写敏感设置进行搜索
#             if case_sensitive:
#                 result = df[df.apply(lambda row: row.astype(str).str.contains(search_term, regex=False).any(), axis=1)]
#             else:
#                 result = df[df.apply(lambda row: row.astype(str).str.contains(search_term, case=False, regex=False).any(), axis=1)]
#
#             if not result.empty:
#                 results[sheet_name] = result
#
#         return results
#     except Exception as e:
#         return {"error": str(e)}

def search_excel(file_path, search_term, case_sensitive=False):
    try:
        # 读取 Excel 文件，不将任何值视为缺失值，并且不使用第一行作为标题
        xls = pd.ExcelFile(file_path)
        results = {}

        # 遍历所有 sheet
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, keep_default_na=False,
                              na_values=[], header=None)

            # 根据大小写敏感设置进行搜索
            if case_sensitive:
                result = df[df.apply(lambda row: row.astype(str).str.contains(search_term, regex=False).any(), axis=1)]
            else:
                result = df[df.apply(lambda row: row.astype(str).str.contains(search_term, case=False, regex=False).any(), axis=1)]

            if not result.empty:
                results[sheet_name] = result

        return results
    except Exception as e:
        return {"error": str(e)}

class ExcelSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel多表搜索工具")
        self.root.geometry("800x600")

        # 文件路径选择
        file_frame = tk.Frame(root)
        file_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(file_frame, text="Excel文件:").pack(side=tk.LEFT)
        self.file_path = tk.StringVar()
        tk.Entry(file_frame, textvariable=self.file_path, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(file_frame, text="浏览...", command=self.browse_file).pack(side=tk.LEFT)

        # 搜索框
        search_frame = tk.Frame(root)
        search_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(search_frame, text="搜索内容:").pack(side=tk.LEFT)
        self.search_term = tk.StringVar()
        tk.Entry(search_frame, textvariable=self.search_term, width=50).pack(side=tk.LEFT, padx=5)

        # 选项
        options_frame = tk.Frame(root)
        options_frame.pack(fill=tk.X, padx=20, pady=5)

        self.case_sensitive = tk.BooleanVar()
        tk.Checkbutton(options_frame, text="区分大小写", variable=self.case_sensitive).pack(side=tk.LEFT)

        # 搜索按钮
        tk.Button(options_frame, text="搜索", command=self.search).pack(side=tk.LEFT, padx=10)

        # 结果显示区域
        result_frame = tk.Frame(root)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 结果标签页
        self.result_notebook = ttk.Notebook(result_frame)
        self.result_notebook.pack(fill=tk.BOTH, expand=True)

        # 状态栏
        self.status = tk.StringVar()
        tk.Label(root, textvariable=self.status, bd=1, relief=tk.SUNKEN, anchor=tk.W).pack(side=tk.BOTTOM, fill=tk.X)

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.file_path.set(filename)

    def search(self):
        # 清除之前的结果
        for tab in self.result_notebook.winfo_children():
            tab.destroy()

        file_path = self.file_path.get()
        search_term = self.search_term.get()

        if not file_path or not search_term:
            self.status.set("请提供Excel文件路径和搜索内容")
            return

        self.status.set("正在搜索...")
        self.root.update()

        results = search_excel(file_path, search_term, self.case_sensitive.get())

        if "error" in results:
            self.status.set(f"错误: {results['error']}")
            return

        if not results:
            self.status.set("未找到匹配内容")
            return

        # 显示结果
        total_rows = 0
        for sheet_name, df in results.items():
            tab = ttk.Frame(self.result_notebook)
            self.result_notebook.add(tab, text=f"{sheet_name} ({len(df)})")

            text_area = scrolledtext.ScrolledText(tab, wrap=tk.WORD)
            text_area.pack(fill=tk.BOTH, expand=True)
            text_area.insert(tk.END, df.to_string())
            text_area.config(state=tk.DISABLED)

            total_rows += len(df)

        self.status.set(f"在 {len(results)} 个表中找到 {total_rows} 行匹配内容")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSearchApp(root)
    root.mainloop()