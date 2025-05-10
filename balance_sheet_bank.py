import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from datetime import datetime
from pandas import ExcelWriter

Tk().withdraw()

# open file dialog for main data
file_path = askopenfilename(title='فایل تراز به تفکیک شعب را انتخاب کنید:', filetypes=[("csv files", "*.csv"), ("Excel files", "*.xlsx *.xls")])
if file_path.endswith(".csv"):
    balance_sheet_by_branch = pd.read_csv(file_path, header=5)
elif file_path.endswith(".xlsx") or file_path.endswith(".xls"):
    balance_sheet_by_branch = pd.read_excel(file_path, header=5)
else:
    raise ValueError("Excel or CSV")


balance_sheet_by_branch.columns = balance_sheet_by_branch.columns.str.strip()

# drop extra cols and na rows
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

# loop each row and build full codes and titles
for _, row in balance_sheet_by_branch.iterrows():
    level = row["سطح"]
    code = row["کد سرفصل"]
    title = str(row["عنوان"]) if pd.notna(row["عنوان"]) else ""

    # pop till matching level
    while len(path_stack_codes) >= level:
        path_stack_codes.pop()
        path_stack_titles.pop()

    path_stack_codes.append(code)
    path_stack_titles.append(title)

    output_paths_codes.append("_".join(path_stack_codes))
    output_paths_titles.append("_".join(path_stack_titles))
    levels.append(level)

# add new cols for full code and title path
balance_sheet_by_branch["کد نهایی"] = output_paths_codes
balance_sheet_by_branch["عنوان نهایی"] = output_paths_titles
balance_sheet_by_branch["سطح مسیر"] = levels


next_levels = levels[1:] + [0]
leaf_mask = [cur >= next for cur, next in zip(levels, next_levels)]
leaf_df = balance_sheet_by_branch[leaf_mask].copy()


print("📄 لطفاً فایل اطلاعات تکمیلی شعب را انتخاب کنید.")
file_path_info = askopenfilename(title='فایل اطلاعات شعب را انتخاب کنید:', filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
if file_path_info.endswith(".csv"):
    branch_info = pd.read_csv(file_path_info)
else:
    branch_info = pd.read_excel(file_path_info)


branch_info.columns = branch_info.columns.str.strip()
branch_info = branch_info.rename(columns={branch_info.columns[0]: "کد شعبه"})


leaf_df["کد شعبه"] = leaf_df["کد شعبه"].astype(float).astype(int).astype(str)
branch_info["کد شعبه"] = branch_info["کد شعبه"].astype(float).astype(int).astype(str)

leaf_df = pd.merge(leaf_df, branch_info[["کد شعبه", "کد فناپ", "کد بانک مرکزی", "نام استان"]], on="کد شعبه", how="left")

# final output
now_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
output_path = f"خروجی_نهایی_برگ‌ها_{now_str}.xlsx"
leaf_df.to_excel(output_path, index=False)

# print summary (just for checkin the result)
print(leaf_df[["کد شعبه", "کد نهایی", "عنوان نهایی", "مانده(بد)", "مانده(بس)", "کد فناپ", "کد بانک مرکزی", "نام استان"]])
print(f"\n✅ خروجی شامل فقط برگ‌ها با مانده صحیح ذخیره شد: {output_path}")




# اضافه کردن کد استان (کد فناپ) به انتهای کد نهایی
leaf_df["کد نهایی با استان"] = leaf_df["کد نهایی"] + "_" + leaf_df["کد فناپ"].astype(str)

leaf_df["عنوان نهایی با استان"] = leaf_df["عنوان نهایی"] + "_" + leaf_df["نام استان"]

grouped_df = leaf_df.groupby(
    ["کد نهایی با استان", "عنوان نهایی با استان"], as_index=False
)[["مانده(بد)", "مانده(بس)", "مانده (بد)", "مانده (بس)"]].sum(min_count=1)

# حذف ردیف‌هایی که هم مانده(بد) و هم مانده(بس) صفر هستند
grouped_df = grouped_df[~((grouped_df["مانده(بد)"] == 0) & (grouped_df["مانده(بس)"] == 0))]

grouped_df["مانده(بد)_اصلاح‌شده"] = grouped_df["مانده(بد)"].clip(lower=0) + grouped_df["مانده(بس)"].clip(upper=0).abs()
grouped_df["مانده(بس)_اصلاح‌شده"] = grouped_df["مانده(بس)"].clip(lower=0) + grouped_df["مانده(بد)"].clip(upper=0).abs()

grouped_df["مانده(بد)_اصلاح‌شده"] = grouped_df["مانده(بد)_اصلاح‌شده"].astype(str)
grouped_df["مانده(بس)_اصلاح‌شده"] = grouped_df["مانده(بس)_اصلاح‌شده"].astype(str)

grouped_df = grouped_df[["کد نهایی با استان", "عنوان نهایی با استان", "مانده(بد)_اصلاح‌شده", "مانده(بس)_اصلاح‌شده"]]

output_grouped_path = f"خروجی_نهایی_تجمیعی_کد_با_استان_{now_str}.xlsx"
with ExcelWriter(output_grouped_path, engine="xlsxwriter") as writer:
    # شیت بدون عنوان
    grouped_df.to_excel(writer, index=False, header=False, sheet_name="تجمیعی بدون عنوان")
    # شیت با عنوان
    grouped_df.to_excel(writer, index=False, header=True, sheet_name="تجمیعی با عنوان")

print(f"\n✅ خروجی تجمیعی فیلترشده با اصلاح مقادیر منفی و تبدیل به متن در دو شیت ذخیره شد: {output_grouped_path}")
