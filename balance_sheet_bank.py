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

import pandas as pd
import matplotlib.pyplot as plt

# لیست کدهای مورد نظر — اینو تو پر کن با کدهایی که می‌خوای
selected_codes = [
    "2000_2001_2043_2150_2230","2000_2001_2043_2150_22300","2000_2001_2043_2150_2232","2000_2001_2043_2160_21601","2000_2001_2043_2160_21605",
    "2000_2001_2043_2160_21608","2000_2001_2043_2160_2161_21421","2000_2001_2043_2160_2190","2000_2001_2043_2160_3116","2000_2001_2030_2200",
    "2000_2001_2030_22001","2000_2001_2030_4244_4623","2000_2001_2032_2453_2639","1000_1008_1085_1250","2000_2001_2043_2150_2231","2000_2001_2043_2150_2233",
    "2000_2001_2043_2160_21602","2000_2001_2043_2160_21606","2000_2001_2043_2160_21609","2000_2001_2043_2160_2161_2142","2000_2001_2043_2160_3170",
    "2000_2001_2043_2160_3172","2000_2001_2043_3045","2000_2001_3251_35180","2000_2006_2056_2117","2000_2010_2093_2170","2000_2012_2099_2952_2662",
    "2000_2006_2126_44440","2000_2012_2113_21131","2000_2010_2109_2097","2000_2010_2109_2105","2000_2001_2030_2201","2000_2001_2030_4244_4624",
    "2000_2001_3251_35181","1000_1010_1112_11123","1000_1010_1112_11124","1000_1010_1112_1354","2000_2001_2032_2453_2648","2000_2008_20840_20841",
    "2000_2004_2048_2443","2000_2004_2050_3111","2000_2004_2107_21071","2000_2001_2034_2017_2020","2000_2001_2034_20342","2000_2001_2034_20347",
    "2000_2001_2034_20352","2000_2001_2034_20353","2000_2001_2034_2211","2000_2001_2034_2214","2000_2001_2034_2696","2000_2001_2034_3074",
    "2000_2001_2034_3076","2000_2001_2034_3079","2000_2001_2034_3140","2000_2001_2034_3177","2000_2001_2034_3210","2000_2001_2034_3266","2000_2001_2034_3272",
    "2000_2001_2034_3301","2000_2001_2034_3315","2000_2001_2034_3321","2000_2001_2034_5235","2000_2001_3251_35182","2000_2001_3251_35195","2000_2001_2036_2221_22212",
    "2000_2001_2036_2221_22215","2000_2001_2036_2221_3149","2000_2001_2036_2221_3274","2000_2001_2036_2221_3341","2000_2001_2036_2221_3353","2000_2001_2036_2221_3359",
    "2000_2001_2036_2221_3367","2000_2001_2036_2221_3370","2000_2001_2036_2222_2511","2000_2001_2036_2223_2517","2000_2001_2036_2223_25170","2000_2010_2093_2400_2588_2770",
    "2000_2001_2036_2221_22211","2000_2001_2036_2221_22214","2000_2001_2036_2221_2505","2000_2001_2036_2221_3247","2000_2001_2036_2221_3259","2000_2001_2036_2221_3269",
    "2000_2001_2036_2221_3273","2000_2001_2036_2221_3294","2000_2001_2036_2221_3340","2000_2001_2036_2221_3346","2000_2001_2036_2221_3352","2000_2001_2036_2221_3358",
    "2000_2001_2036_2221_3361","2000_2001_2036_2221_3363","2000_2001_2036_2221_3366","2000_2001_2036_2221_3368","2000_2001_2036_2222_2510","2000_2001_2036_2222_25100",
    "2000_2001_2036_2223_2516","2000_2001_2036_2223_25160","2000_2001_2035_31410","2000_2001_2035_2217","2000_2001_2035_31430","2000_2001_2034_2017_2018",
    "2000_2001_2034_2017_20181","2000_2001_2034_20341","2000_2001_2034_20346","2000_2001_2034_20350","2000_2001_2034_20351","2000_2001_2034_2210","2000_2001_2034_2213",
    "2000_2001_2034_2695","2000_2001_2034_3073","2000_2001_2034_3075","2000_2001_2034_3078","2000_2001_2034_3091_3085","2000_2001_2034_3139","2000_2001_2034_3176",
    "2000_2001_2034_3209","2000_2001_2034_3265","2000_2001_2034_3271","2000_2001_2034_3296","2000_2001_2034_3302","2000_2001_2034_3316","2000_2001_2034_3318",
    "2000_2001_2034_3320","2000_2001_2034_5234","2000_2010_2093_20935","2000_2001_2033_2452_2631","2000_2001_2033_2209_2607","1000_1010_1164_1114","2000_2010_2093_2179_23951",
    "2000_2010_2093_2179_2398_2568","2000_2010_2093_2179_2398_25680","2000_2010_2093_2179_2398_25681","2000_2010_2093_2179_2398_2572","2000_2010_2093_2179_2404_2578",
    "2000_2010_2093_2179_2404_25781","2000_2010_2093_2179_2404_25782","2000_2010_2093_2179_2398_2569","2000_2010_2093_2179_2398_25690","2000_2010_2093_2179_2404_2579",
    "2000_2010_2093_2179_2398_2570","2000_2010_2093_2179_2398_25700","2000_2010_2093_2179_2404_2580","2000_2010_2093_2179_2404_25800","2000_2010_2093_2179_2395",
    "2000_2010_2093_2179_23950","2000_2010_2093_2179_2401","2000_2010_2093_2179_24010","2000_2010_2093_2179_24011","2000_2010_2093_2179_4253_4687","2000_2010_2093_3198_3200",
    "2000_2010_2093_3198_3203","2000_2010_2093_2179_2396","2000_2010_2093_2179_23960","2000_2012_2115_4257_4660_2868","2000_2012_2115_4257_4660_28681",
    "2000_2012_2115_4257_4660_2869","2000_2012_2115_4257_4681_2871","2000_2012_2114_4211","2000_2012_2114_4213","2000_2012_2114_4215","2000_2001_20360_3371",
    "2000_2010_2093_3311","2000_2010_2093_20934","2000_2010_2094_20532","2000_2010_2094_20962"
]

# فیلتر فقط ردیف‌هایی که کد نهایی‌شون در لیست بالا هست
filtered_df = leaf_df[leaf_df["کد نهایی"].isin(selected_codes)].copy()

# محاسبه مانده خالص
filtered_df["مانده خالص"] = filtered_df["مانده(بس)"] - filtered_df["مانده(بد)"]

# گروه‌بندی بر اساس نام استان و جمع مانده خالص
grouped = filtered_df.groupby("نام استان")["مانده خالص"].sum().sort_values(ascending=False)

# پیدا کردن سه استان با بیشترین مانده
top3 = grouped.head(3).index

# رسم نمودار میله‌ای
colors = ["red" if province in top3 else "gray" for province in grouped.index]

plt.figure(figsize=(12, 6))
bars = plt.bar(grouped.index, grouped.values, color=colors)
plt.title("مانده سپرده مشتریان بر اساس استان")
plt.xlabel("نام استان")
plt.ylabel("سپرده مشتریان")
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()
