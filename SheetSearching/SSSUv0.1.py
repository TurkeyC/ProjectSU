import pandas as pd

def search_excel(file_path, search_term):
    # 读取 Excel 文件，不将任何值视为缺失值
    xls = pd.ExcelFile(file_path)
    results = {}

    # 遍历所有 sheet
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, keep_default_na=False, na_values=[])
        # 搜索 term
        result = df[df.apply(lambda row: row.astype(str).str.contains(search_term).any(), axis=1)]
        if not result.empty:
            results[sheet_name] = result

    return results

if __name__ == "__main__":
    file_path = input("Enter the path to the Excel file: ")
    search_term = input("Enter the search term: ")
    results = search_excel(file_path, search_term)

    if results:
        for sheet, result in results.items():
            print(f"Results in sheet {sheet}:")
            print(result)
    else:
        print("No results found.")