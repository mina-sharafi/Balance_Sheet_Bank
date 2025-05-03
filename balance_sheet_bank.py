import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from datetime import datetime

Tk().withdraw()
file_path = askopenfilename(title = 'فایل تراز به تفکیک شعب را انتخاب کنید:' , filetypes=[("csv files","*.csv") , ("Excel files" , "*.xlsx *.xls")])
if file_path.endswith(".csv"):
    balance_sheet_by_branch = pd.read_csv(file_path, header=5)
elif file_path.endswith(".xlsx") or file_path.endswith(".xls"):
    balance_sheet_by_branch = pd.read_excel(file_path, header=5)
else:
    raise ValueError("Excel or CSV")

#balance_sheet_by_branch = pd.read_csv(r"C:\Users\Mina\Desktop\kholaseh daftar kol\balance_sheet_by_branch.csv", header=5)

balance_sheet_by_branch.columns = balance_sheet_by_branch.columns.str.strip()

balance_sheet_by_branch = balance_sheet_by_branch.iloc[:, 1:13].dropna()
balance_sheet_by_branch["سطح"] = balance_sheet_by_branch["سطح"].astype(int)
balance_sheet_by_branch["کد سرفصل"] = balance_sheet_by_branch["کد سرفصل"].astype(str)
balance_sheet_by_branch.iloc[:, 0] = balance_sheet_by_branch.iloc[:, 0].astype(str)

for col in balance_sheet_by_branch.columns[4:]:
    balance_sheet_by_branch[col] = balance_sheet_by_branch[col].replace(",", "", regex=True)
    balance_sheet_by_branch[col] = pd.to_numeric(balance_sheet_by_branch[col], errors='coerce')

path_stack_codes = []
path_stack_titles = []
output_paths_codes = []
output_paths_titles = []
levels = []

for _, row in balance_sheet_by_branch.iterrows():
    level = row["سطح"]
    code = row["کد سرفصل"]
    title = str(row["عنوان"]) if pd.notna(row["عنوان"]) else ""

    while len(path_stack_codes) >= level:
        path_stack_codes.pop()
        path_stack_titles.pop()

    path_stack_codes.append(code)
    path_stack_titles.append(title)

    output_paths_codes.append("_".join(path_stack_codes))
    output_paths_titles.append("_".join(path_stack_titles))
    levels.append(level)

balance_sheet_by_branch["کد نهایی"] = output_paths_codes
balance_sheet_by_branch["عنوان نهایی"] = output_paths_titles
balance_sheet_by_branch["سطح مسیر"] = levels

next_levels = levels[1:] + [0]
leaf_mask = [cur >= next for cur, next in zip(levels, next_levels)]
leaf_df = balance_sheet_by_branch[leaf_mask].copy()

now_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
output_path = f"خروجی_نهایی_برگ‌ها_{now_str}.xlsx"
leaf_df.to_excel(output_path, index=False)


print(leaf_df[["کد شعبه", "کد نهایی", "عنوان نهایی", "مانده(بد)", "مانده(بس)"]])
print(f"\n✅ خروجی شامل فقط برگ‌ها با مانده صحیح ذخیره شد: {output_path}")
