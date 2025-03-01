import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image
import sys
import os
import glob
from pathlib import Path
import importlib.util
import tempfile
import shutil
import time

# Set appearance mode and default color theme
ctk.set_appearance_mode("System")  # Modes: "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue", "green", "dark-blue"

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath("MyBuild")

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

class ModernExcelSearchApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Basic configuration
        self.title("Excel多表搜索工具 v0.8")
        self.geometry("1200x800")
        self.minsize(1000, 700)

        # Load icons (you can replace with your own icons)
        try:
            self.icon_path = {
                "search": ctk.CTkImage(light_image=Image.new("RGB", (24, 24), color="#007acc"),
                                       dark_image=Image.new("RGB", (24, 24), color="#007acc"),
                                       size=(24, 24)),
                "settings": ctk.CTkImage(light_image=Image.new("RGB", (24, 24), color="#555555"),
                                         dark_image=Image.new("RGB", (24, 24), color="#aaaaaa"),
                                         size=(24, 24))
            }
        except Exception:
            # Fallback if image loading fails
            self.icon_path = {}

        # Initialize variables
        self.file_paths = []
        self.search_term = ctk.StringVar()
        self.case_sensitive = ctk.BooleanVar(value=False)
        self.status = ctk.StringVar(value="就绪")

        # Build the interface
        self._create_sidebar()
        self._create_main_content()
        self._create_bottom_bar()

    def _create_sidebar(self):
        """Left sidebar navigation"""
        self.sidebar = ctk.CTkFrame(self, width=240, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)  # Prevent shrinking

        # App title
        title_label = ctk.CTkLabel(
            self.sidebar,
            text="Excel搜索工具",
            font=ctk.CTkFont(size=18, weight="bold"),
            anchor="w"
        )
        title_label.pack(fill="x", padx=20, pady=(20, 10))

        # Separator
        separator = ctk.CTkFrame(self.sidebar, height=1)
        separator.pack(fill="x", padx=20, pady=10)

        # Navigation buttons
        self.search_btn = ctk.CTkButton(
            self.sidebar,
            text="文件搜索",
            image=self.icon_path.get("search", None),
            anchor="w",
            height=40,
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            corner_radius=5
        )
        self.search_btn.pack(fill="x", padx=20, pady=5)

        # Theme switch at bottom of sidebar
        self.theme_switch = ctk.CTkSwitch(
            self.sidebar,
            text="深色模式",
            command=self._toggle_theme,
            progress_color="#2FA572"
        )
        self.theme_switch.pack(side="bottom", pady=20, padx=20, anchor="w")

        # Set the switch based on current appearance mode
        if ctk.get_appearance_mode() == "Dark":
            self.theme_switch.select()

    def _create_main_content(self):
        """Main content area"""
        self.main_frame = ctk.CTkFrame(self, corner_radius=0)
        self.main_frame.pack(side="right", fill="both", expand=True)

        # Create a tabview for organization
        self.tab_view = ctk.CTkTabview(self.main_frame)
        self.tab_view.pack(fill="both", expand=True, padx=10, pady=10)

        # Search tab
        self.search_tab = self.tab_view.add("搜索")
        self._init_search_interface()

    def _init_search_interface(self):
        """Build the search interface"""
        # File selection section
        file_frame = ctk.CTkFrame(self.search_tab)
        file_frame.pack(fill="x", padx=10, pady=10, expand=False)

        file_label = ctk.CTkLabel(
            file_frame,
            text="文件选择",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        file_label.pack(anchor="w", padx=10, pady=5)

        # Files listbox (with a frame for proper styling)
        files_frame = ctk.CTkFrame(file_frame)
        files_frame.pack(fill="x", padx=10, pady=5, expand=False)

        self.files_listbox = tk.Listbox(
            files_frame,
            height=5,
            borderwidth=1,
            relief="solid",
            background="#ffffff" if ctk.get_appearance_mode() == "Light" else "#2b2b2b",
            foreground="#000000" if ctk.get_appearance_mode() == "Light" else "#ffffff",
            selectbackground="#007acc",
            highlightthickness=0
        )
        self.files_listbox.pack(side="left", padx=5, fill="both", expand=True)

        # Scrollbar for listbox
        listbox_scrollbar = ctk.CTkScrollbar(
            files_frame,
            command=self.files_listbox.yview
        )
        self.files_listbox.config(yscrollcommand=listbox_scrollbar.set)
        listbox_scrollbar.pack(side="right", fill="y")

        # File action buttons
        buttons_frame = ctk.CTkFrame(file_frame)
        buttons_frame.pack(fill="x", padx=10, pady=10)

        browse_file_btn = ctk.CTkButton(
            buttons_frame,
            text="浏览文件",
            command=self.browse_files,
            fg_color="#007acc",
            hover_color="#005fa3"
        )
        browse_file_btn.pack(side="left", padx=5)

        browse_folder_btn = ctk.CTkButton(
            buttons_frame,
            text="浏览文件夹",
            command=self.browse_folder,
            fg_color="#007acc",
            hover_color="#005fa3"
        )
        browse_folder_btn.pack(side="left", padx=5)

        clear_btn = ctk.CTkButton(
            buttons_frame,
            text="清除列表",
            command=self.clear_files,
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            border_width=1,
            hover_color=("gray70", "gray30")
        )
        clear_btn.pack(side="left", padx=5)

        # Search section
        search_frame = ctk.CTkFrame(self.search_tab)
        search_frame.pack(fill="x", padx=10, pady=10, expand=False)

        search_label = ctk.CTkLabel(
            search_frame,
            text="搜索配置",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        search_label.pack(anchor="w", padx=10, pady=5)

        # Search input section
        search_input_frame = ctk.CTkFrame(search_frame)
        search_input_frame.pack(fill="x", padx=10, pady=5, expand=False)

        search_text_label = ctk.CTkLabel(search_input_frame, text="搜索内容:")
        search_text_label.pack(side="left", padx=5)

        search_entry = ctk.CTkEntry(
            search_input_frame,
            textvariable=self.search_term,
            placeholder_text="输入搜索内容...",
            height=32
        )
        search_entry.pack(side="left", padx=5, fill="x", expand=True)

        # Search options
        options_frame = ctk.CTkFrame(search_frame)
        options_frame.pack(fill="x", padx=10, pady=10)

        case_check = ctk.CTkCheckBox(
            options_frame,
            text="区分大小写",
            variable=self.case_sensitive
        )
        case_check.pack(side="left", padx=5)

        search_btn = ctk.CTkButton(
            options_frame,
            text="搜索",
            command=self.search,
            fg_color="#007acc",
            hover_color="#005fa3",
            width=120
        )
        search_btn.pack(side="right", padx=5)

        # Results section
        results_frame = ctk.CTkFrame(self.search_tab)
        results_frame.pack(fill="both", padx=10, pady=10, expand=True)

        results_label = ctk.CTkLabel(
            results_frame,
            text="搜索结果",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        results_label.pack(anchor="w", padx=10, pady=5)

        # Use a notebook for results
        self.result_notebook_frame = ctk.CTkFrame(results_frame)
        self.result_notebook_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # This will be created when needed with the search results
        self.result_notebook = None

    def _create_bottom_bar(self):
        """Bottom status bar"""
        self.status_bar = ctk.CTkFrame(self, height=28, corner_radius=0)
        self.status_bar.pack(side="bottom", fill="x")

        self.status_label = ctk.CTkLabel(
            self.status_bar,
            textvariable=self.status,
            anchor="w",
            font=ctk.CTkFont(size=12)
        )
        self.status_label.pack(fill="x", padx=20)

    def _toggle_theme(self):
        """Toggle between light and dark mode"""
        current = ctk.get_appearance_mode()
        new_mode = "Dark" if current == "Light" else "Light"
        ctk.set_appearance_mode(new_mode)

        # Update the listbox colors based on the new theme
        self.files_listbox.config(
            background="#ffffff" if new_mode == "Light" else "#2b2b2b",
            foreground="#000000" if new_mode == "Light" else "#ffffff"
        )

    def browse_files(self):
        filenames = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filenames:
            for filename in filenames:
                if filename not in self.file_paths:
                    self.file_paths.append(filename)
                    self.files_listbox.insert(tk.END, Path(filename).name)
            self.status.set(f"已添加 {len(filenames)} 个文件")

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            excel_files = glob.glob(os.path.join(folder, "*.xlsx")) + glob.glob(os.path.join(folder, "*.xls"))
            count = 0
            for file in excel_files:
                if file not in self.file_paths:
                    self.file_paths.append(file)
                    self.files_listbox.insert(tk.END, Path(file).name)
                    count += 1
            self.status.set(f"已从文件夹添加 {count} 个Excel文件")

    def clear_files(self):
        self.file_paths = []
        self.files_listbox.delete(0, tk.END)
        self.status.set("已清除文件列表")

    def search(self):
        # Clear previous results by removing the notebook widget if it exists
        if self.result_notebook is not None:
            self.result_notebook.destroy()

        file_paths = self.file_paths
        search_term = self.search_term.get()

        if not file_paths or not search_term:
            self.status.set("请选择Excel文件并提供搜索内容")
            return

        self.status.set("正在搜索...")
        self.update()  # Update the UI to show status change

        # Perform search
        results = search_excel_files(file_paths, search_term, self.case_sensitive.get())

        # Check if we got any results
        if not results:
            self.status.set("未找到匹配内容")
            return

        # Create a new notebook for results
        self.result_notebook = ctk.CTkTabview(self.result_notebook_frame)
        self.result_notebook.pack(fill="both", expand=True)

        # Display results
        total_files = len(results)
        total_sheets = 0
        total_rows = 0

        for file_name, file_results in results.items():
            # Create a tab for each file
            file_tab = self.result_notebook.add(file_name)

            if "error" not in file_results:
                file_sheet_count = len(file_results)
                file_row_count = sum(len(df) for df in file_results.values())

                # Create a nested tabview for sheets
                sheet_tabview = ctk.CTkTabview(file_tab)
                sheet_tabview.pack(fill="both", expand=True, padx=5, pady=5)

                for sheet_name, df in file_results.items():
                    # Create tab for each sheet
                    sheet_tab = sheet_tabview.add(f"{sheet_name} ({len(df)})")

                    # Create text area for displaying data
                    text_frame = ctk.CTkFrame(sheet_tab)
                    text_frame.pack(fill="both", expand=True, padx=5, pady=5)

                    text_area = ctk.CTkTextbox(text_frame, wrap="none", font=ctk.CTkFont(family="Consolas", size=12))
                    text_area.pack(fill="both", expand=True)
                    text_area.insert("1.0", df.to_string())
                    text_area.configure(state="disabled")

                    total_rows += len(df)

                total_sheets += file_sheet_count
            else:
                # Display error message
                text_area = ctk.CTkTextbox(file_tab, wrap="word", font=ctk.CTkFont(family="Consolas", size=12))
                text_area.pack(fill="both", expand=True, padx=10, pady=10)
                text_area.insert("1.0", f"处理文件时出错: {file_results['error']}")
                text_area.configure(state="disabled")

        self.status.set(f"在 {total_files} 个文件的 {total_sheets} 个表中找到 {total_rows} 行匹配内容")


if __name__ == "__main__":
    # Add exception handling
    try:
        app = ModernExcelSearchApp()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("严重错误", f"程序崩溃: {str(e)}")