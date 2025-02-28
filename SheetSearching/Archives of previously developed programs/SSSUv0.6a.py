import pandas as pd
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk, messagebox
import sys
import os
import glob
from pathlib import Path
import importlib.util
import tempfile
import shutil
import time

def repair_excel_with_com(file_path):
    """
    Repair Excel file using Excel's COM interface.
    This simulates opening the file in Excel and saving it as a new file.
    """
    try:
        import win32com.client
    except ImportError:
        raise ImportError("需要安装pywin32库才能进行Excel修复。请运行: pip install pywin32")

    tmp_path = None
    excel_app = None

    try:
        # Create a temporary file path for the repaired file
        tmp_fd, tmp_path = tempfile.mkstemp(suffix='.xlsx')
        os.close(tmp_fd)

        # Create Excel application
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.DisplayAlerts = False  # Don't show alerts
        excel_app.Visible = False        # Don't show Excel

        # Open the problematic file
        workbook = excel_app.Workbooks.Open(os.path.abspath(file_path))

        # Save as new file (this is the repair step)
        workbook.SaveAs(os.path.abspath(tmp_path))
        workbook.Close()

        # Read the repaired file with pandas
        xl = pd.ExcelFile(tmp_path)
        sheet_names = xl.sheet_names

        results = {}
        for sheet_name in sheet_names:
            df = pd.read_excel(
                tmp_path,
                sheet_name=sheet_name,
                header=None,
                dtype=str,
                na_filter=False
            )
            results[sheet_name] = df

        return results

    except Exception as e:
        raise Exception(f"Excel COM修复失败: {str(e)}")

    finally:
        # Clean up resources
        if excel_app:
            excel_app.Quit()

        # Clean up the temporary file
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except:
                pass

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath("../MyBuild")

    return os.path.join(base_path, relative_path)

def is_package_installed(package_name):
    """Check if a package is installed"""
    return importlib.util.find_spec(package_name) is not None

def get_available_engines():
    """Return a list of available Excel engines based on installed packages"""
    engines = []

    # Always include default engine
    engines.append({'engine': 'openpyxl', 'options': {}})
    engines.append({'engine': 'openpyxl', 'options': {'read_only': True, 'data_only': True}})

    # Check for other engines
    if is_package_installed('xlrd'):
        engines.append({'engine': 'xlrd', 'options': {}})

    if is_package_installed('pyxlsb'):
        engines.append({'engine': 'pyxlsb', 'options': {}})

    if is_package_installed('odf'):
        engines.append({'engine': 'odf', 'options': {}})

    return engines

def read_problematic_excel(file_path):
    """
    Special method for handling problematic Excel files by creating a temporary copy.
    This simulates the "copy-paste to new file" approach that fixed the original issue.
    """
    tmp_path = None
    try:
        # Create a temporary file
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
            tmp_path = tmp_file.name

        # Copy the original file to the temporary location
        shutil.copy2(file_path, tmp_path)

        # Try to read with pandas from the temporary file
        # First try to get sheet names
        xl = pd.ExcelFile(tmp_path)
        sheet_names = xl.sheet_names

        results = {}
        for sheet_name in sheet_names:
            df = pd.read_excel(
                tmp_path,
                sheet_name=sheet_name,
                header=None,
                dtype=str,
                na_filter=False
            )
            results[sheet_name] = df

        # Clean up the temporary file
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)

        return results

    except Exception as e:
        # Clean up if still exists
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except:
                pass
        raise e

def search_excel_files(file_paths, search_term, case_sensitive=False):
    all_results = {}
    available_engines = get_available_engines()

    for file_path in file_paths:
        file_name = Path(file_path).name
        file_results = {}
        sheet_errors = []
        last_error = ""

        # Step 1: Try all standard engines
        for engine_config in available_engines:
            engine = engine_config['engine']
            options = engine_config['options']

            try:
                # 尝试先获取表名
                xl = pd.ExcelFile(file_path, engine=engine, **options)
                sheet_names = xl.sheet_names

                for sheet_name in sheet_names:
                    try:
                        # 使用严格的文本模式读取数据
                        df = pd.read_excel(
                            file_path,
                            sheet_name=sheet_name,
                            engine=engine,
                            header=None,
                            dtype=str,
                            na_filter=False,
                            keep_default_na=False,
                            **options
                        )

                        # 优化搜索逻辑
                        mask = df.apply(
                            lambda row: row.astype(str).str.contains(
                                search_term,
                                case=case_sensitive,
                                regex=False,
                                na=False
                            ).any(),
                            axis=1
                        )
                        result = df[mask]

                        if not result.empty:
                            file_results[sheet_name] = result

                    except Exception as e:
                        sheet_errors.append(f"[{sheet_name}] {str(e)}")
                        continue

                if file_results:
                    all_results[file_name] = file_results
                    break  # 找到结果就停止尝试其他引擎

            except Exception as e:
                last_error = str(e)
                continue

        # Step 2: If all standard approaches failed, try the temp file approach
        if not file_results:
            try:
                # Try the temporary file approach
                repaired_sheets = read_problematic_excel(file_path)

                # Search in the repaired data
                for sheet_name, df in repaired_sheets.items():
                    mask = df.apply(
                        lambda row: row.astype(str).str.contains(
                            search_term,
                            case=case_sensitive,
                            regex=False,
                            na=False
                        ).any(),
                        axis=1
                    )
                    result = df[mask]

                    if not result.empty:
                        file_results[sheet_name] = result

                if file_results:
                    all_results[file_name] = file_results

            except Exception as e:
                last_error = f"{last_error}; 常规修复尝试失败: {str(e)}"

        # Step 3: If all previous approaches failed, try using Excel COM automation
        if not file_results:
            try:
                # Try Excel COM automation repair
                excel_repaired_sheets = repair_excel_with_com(file_path)

                # Search in the Excel-repaired data
                for sheet_name, df in excel_repaired_sheets.items():
                    mask = df.apply(
                        lambda row: row.astype(str).str.contains(
                            search_term,
                            case=case_sensitive,
                            regex=False,
                            na=False
                        ).any(),
                        axis=1
                    )
                    result = df[mask]

                    if not result.empty:
                        file_results[sheet_name] = result

                if file_results:
                    all_results[file_name] = file_results

            except Exception as e:
                last_error = f"{last_error}; Excel COM修复尝试失败: {str(e)}"

        # Record errors if all attempts failed
        if not file_results:
            error_msg = last_error if last_error else "；".join(sheet_errors) if sheet_errors else "未知错误"
            all_results[file_name] = {"error": f"所有解析方式失败: {error_msg}"}

    return all_results

class ExcelSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel多表搜索工具")
        self.root.geometry("800x600")

        # 文件路径选择
        file_frame = tk.Frame(root)
        file_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(file_frame, text="Excel文件:").pack(side=tk.LEFT)
        self.file_paths = []
        self.files_listbox = tk.Listbox(file_frame, width=50, height=3)
        self.files_listbox.pack(side=tk.LEFT, padx=5, fill=tk.BOTH, expand=True)

        file_buttons_frame = tk.Frame(file_frame)
        file_buttons_frame.pack(side=tk.LEFT)

        tk.Button(file_buttons_frame, text="浏览文件...", command=self.browse_files).pack(fill=tk.X, pady=2)
        tk.Button(file_buttons_frame, text="浏览文件夹...", command=self.browse_folder).pack(fill=tk.X, pady=2)
        tk.Button(file_buttons_frame, text="清除列表", command=self.clear_files).pack(fill=tk.X, pady=2)

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

        # 结果标签页 - 双层
        self.result_notebook = ttk.Notebook(result_frame)
        self.result_notebook.pack(fill=tk.BOTH, expand=True)

        # 状态栏
        self.status = tk.StringVar()
        tk.Label(root, textvariable=self.status, bd=1, relief=tk.SUNKEN, anchor=tk.W).pack(side=tk.BOTTOM, fill=tk.X)

    def browse_files(self):
        filenames = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filenames:
            for filename in filenames:
                if filename not in self.file_paths:
                    self.file_paths.append(filename)
                    self.files_listbox.insert(tk.END, Path(filename).name)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            excel_files = glob.glob(os.path.join(folder, "*.xlsx")) + glob.glob(os.path.join(folder, "*.xls"))
            for file in excel_files:
                if file not in self.file_paths:
                    self.file_paths.append(file)
                    self.files_listbox.insert(tk.END, Path(file).name)

    def clear_files(self):
        self.file_paths = []
        self.files_listbox.delete(0, tk.END)

    def search(self):
        # 清除之前的结果
        for tab in self.result_notebook.winfo_children():
            tab.destroy()

        file_paths = self.file_paths
        search_term = self.search_term.get()

        if not file_paths or not search_term:
            self.status.set("请选择Excel文件并提供搜索内容")
            return

        self.status.set("正在搜索...")
        self.root.update()

        results = search_excel_files(file_paths, search_term, self.case_sensitive.get())

        # 检查是否所有文件都遇到错误
        if not results:
            self.status.set("未找到匹配内容")
            return

        # 显示结果
        total_files = len(results)
        total_sheets = 0
        total_rows = 0

        for file_name, file_results in results.items():
            # 为每个文件创建一个标签页
            file_tab = ttk.Frame(self.result_notebook)

            # 跳过错误文件的行计数
            if "error" not in file_results:
                file_sheet_count = len(file_results)
                file_row_count = sum(len(df) for df in file_results.values())
                self.result_notebook.add(file_tab, text=f"{file_name} ({file_sheet_count}表/{file_row_count}行)")

                # 为文件内的每个sheet创建子标签页
                sheet_notebook = ttk.Notebook(file_tab)
                sheet_notebook.pack(fill=tk.BOTH, expand=True)

                for sheet_name, df in file_results.items():
                    sheet_tab = ttk.Frame(sheet_notebook)
                    sheet_notebook.add(sheet_tab, text=f"{sheet_name} ({len(df)})")

                    text_area = scrolledtext.ScrolledText(sheet_tab, wrap=tk.WORD)
                    text_area.pack(fill=tk.BOTH, expand=True)
                    text_area.insert(tk.END, df.to_string())
                    text_area.config(state=tk.DISABLED)

                    total_rows += len(df)

                total_sheets += file_sheet_count
            else:
                self.result_notebook.add(file_tab, text=file_name)
                text_area = scrolledtext.ScrolledText(file_tab, wrap=tk.WORD)
                text_area.pack(fill=tk.BOTH, expand=True)
                text_area.insert(tk.END, f"处理文件时出错: {file_results['error']}")
                text_area.config(state=tk.DISABLED)

        self.status.set(f"在 {total_files} 个文件的 {total_sheets} 个表中找到 {total_rows} 行匹配内容")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSearchApp(root)
    # 添加异常捕获
    try:
        root.mainloop()
    except Exception as e:
        messagebox.showerror("严重错误", f"程序崩溃: {str(e)}")