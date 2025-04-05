import pandas as pd
from colorama import Fore, Style

# مسیر فایل اکسل
file_path = r"C:\Users\hosseini.am\Desktop\Report_costofproduce_backup - Copy.xlsx"

# خواندن شیت‌های مورد نیاز
df_produce = pd.read_excel(file_path, sheet_name="dataofproduce", dtype=str, engine="openpyxl")
df_daily = pd.read_excel(file_path, sheet_name="DailyReport", engine="openpyxl", header=1, dtype=str)
df_machine_overhead = pd.read_excel(file_path, sheet_name="RawMaterialCost", usecols="AF:AJ", dtype=str, engine="openpyxl")
df_raw_material = pd.read_excel(file_path, sheet_name="RawMaterialCost", dtype={"کد کالای مواد اولیه": str}, engine="openpyxl")

# حذف فضای اضافی از نام ستون‌ها
df_produce.columns = df_produce.columns.str.strip()
df_daily.columns = df_daily.columns.str.strip()
df_machine_overhead.columns = df_machine_overhead.columns.str.strip()
df_raw_material.columns = df_raw_material.columns.str.strip()

# تبدیل مقادیر عددی به float
df_machine_overhead["max speed"] = pd.to_numeric(df_machine_overhead["max speed"], errors='coerce')
df_machine_overhead["سربار"] = pd.to_numeric(df_machine_overhead["سربار"], errors='coerce')
df_machine_overhead["ماه"] = pd.to_numeric(df_machine_overhead["ماه"], errors='coerce')

# دریافت اطلاعات از کاربر
user_input_kala = input(f"{Fore.CYAN}🔹 لطفاً کد جایگزین اصلی را وارد کنید: {Style.RESET_ALL}").strip()
total_production = float(input(f"{Fore.CYAN}🔹 لطفاً مقدار کل تولید خالص را وارد کنید (کیلوگرم): {Style.RESET_ALL}").strip())
waste_price = float(input(f"{Fore.CYAN}🔹 لطفاً قیمت هر کیلوگرم ضایعات را وارد کنید: {Style.RESET_ALL}").strip())

# فیلتر فرمول‌ها و سفارش‌ها
filtered_produce = df_produce[df_produce["کد جایگزین اصلی"] == user_input_kala]
formulas_used = filtered_produce["کد فرمول"].unique()
order_numbers = filtered_produce["شماره سفارش تولید"].unique()

results = []

if len(formulas_used) == 0:
    print(f"{Fore.RED}⛔ کد کالا {user_input_kala} در داده‌ها پیدا نشد!{Style.RESET_ALL}")
else:
    print(f"\n✅ کد فرمول‌های مرتبط با {Fore.GREEN}{user_input_kala}{Style.RESET_ALL}: {formulas_used}")
    print(f"✅ شماره سفارش‌های مرتبط: {Fore.GREEN}{order_numbers}{Style.RESET_ALL}")

    for formula in formulas_used:
        for order in order_numbers:
            print(f"\n🔹 {Fore.YELLOW}**محاسبه برای کد فرمول {formula} و شماره سفارش {order}:**{Style.RESET_ALL}")

            df_formula_data = df_produce[
                (df_produce["کد فرمول"] == formula) & 
                (df_produce["شماره سفارش تولید"] == order)
            ]

            if df_formula_data.empty:
                print(f"{Fore.RED}⛔ هیچ داده‌ای برای کد فرمول {formula} و شماره سفارش {order} یافت نشد!{Style.RESET_ALL}")
                continue

            machine_names = df_formula_data["نام دستگاه"].unique()

            for machine_name in machine_names:
                df_machine_data = df_formula_data[df_formula_data["نام دستگاه"] == machine_name]

                total_clean_production = df_machine_data["khales"].astype(float).sum()
                total_gross_production = df_machine_data["na_khales"].astype(float).sum()
                total_waste = df_machine_data["وزن کل ضایعات (kg)"].astype(float).sum()
                total_bobin_weight = df_machine_data["وزن بوبین کل"].astype(float).sum()
                order_month = int(df_machine_data["ماه"].astype(float).mode()[0])

                print(f"\n🔹 {Fore.BLUE}**نام دستگاه: {machine_name}**{Style.RESET_ALL}")
                print(f"📊 مقدار تولید خالص: {Fore.GREEN}{total_clean_production:.2f}{Style.RESET_ALL} کیلوگرم")
                print(f"📊 مقدار تولید ناخالص: {Fore.GREEN}{total_gross_production:.2f}{Style.RESET_ALL} کیلوگرم")
                print(f"📊 مقدار ضایعات: {Fore.RED}{total_waste:.2f}{Style.RESET_ALL} کیلوگرم")
                print(f"📊 مقدار وزن بوبین: {Fore.GREEN}{total_bobin_weight:.2f}{Style.RESET_ALL} کیلوگرم")
                print(f"📊 ماه تولید سفارش: {Fore.CYAN}{order_month}{Style.RESET_ALL}")

                machine_overhead_row = df_machine_overhead[
                    (df_machine_overhead["نام دستگاه"] == machine_name) & 
                    (df_machine_overhead["ماه"] == order_month)
                ]
                machine_overhead_cost = machine_overhead_row["سربار"].values[0] if not machine_overhead_row.empty else 0

                print(f"📊 هزینه سربار: 💰 {Fore.YELLOW}{machine_overhead_cost}{Style.RESET_ALL}")

                total_material_cost = 0  
                material_details = ""
                df_materials_info = df_daily[
                    (df_daily["کد فرمول"] == formula) & 
                    (df_daily["شماره سفارش"] == order)
                ][["کد کالای مواد اولیه", "استفاده مواد اولیه"]]

                if df_materials_info.empty:
                    print(f"{Fore.RED}⛔ هیچ ماده اولیه‌ای برای کد فرمول {formula} و شماره سفارش {order} یافت نشد!{Style.RESET_ALL}")
                    material_details = "⛔ یافت نشد"
                else:
                    df_materials_info = df_materials_info.merge(
                        df_raw_material[["کد کالای مواد اولیه", "قیمت واحد"]],
                        on="کد کالای مواد اولیه",
                        how="left"
                    )
                    df_materials_info["هزینه مواد اولیه"] = (
                        df_materials_info["استفاده مواد اولیه"].astype(float) * df_materials_info["قیمت واحد"].astype(float)
                    )
                    total_material_cost = df_materials_info["هزینه مواد اولیه"].sum()

                    print(f"\n📊 {Fore.MAGENTA}**بهای تمام‌شده مواد اولیه:**{Style.RESET_ALL}")
                    for _, row in df_materials_info.iterrows():
                        print(f"- {row['کد کالای مواد اولیه']}: {Fore.CYAN}{row['استفاده مواد اولیه']} kg{Style.RESET_ALL} × "
                              f"{Fore.YELLOW}{row['قیمت واحد']}{Style.RESET_ALL} = 💰 {Fore.GREEN}{row['هزینه مواد اولیه']:.2f}{Style.RESET_ALL}")
                        material_details += f"{row['کد کالای مواد اولیه']}: {row['استفاده مواد اولیه']}kg × {row['قیمت واحد']} = {row['هزینه مواد اولیه']:.0f}\n"

                total_cost = float(machine_overhead_cost) + total_material_cost + (total_waste * waste_price)
                print(f"\n✅ **بهای تمام‌شده کل: 💰 {Fore.GREEN}{total_cost:.2f}{Style.RESET_ALL}**")

                results.append({
                    "کد فرمول": formula,
                    "شماره سفارش": order,
                    "نام دستگاه": machine_name,
                    "تولید خالص (kg)": total_clean_production,
                    "تولید ناخالص (kg)": total_gross_production,
                    "ضایعات (kg)": total_waste,
                    "وزن بوبین کل (kg)": total_bobin_weight,
                    "ماه": order_month,
                    "جزئیات مواد اولیه": material_details.strip(),
                    "هزینه سربار (ریال)": float(machine_overhead_cost),
                    "هزینه مواد اولیه (ریال)": total_material_cost,
                    "هزینه ضایعات (ریال)": total_waste * waste_price,
                    "هزینه کل نهایی (ریال)": total_cost
                })

# نمایش جدول نهایی
df_results = pd.DataFrame(results)

if df_results.empty:
    print(f"{Fore.RED}❌ هیچ داده‌ای برای نمایش در جدول نهایی یافت نشد.{Style.RESET_ALL}")
else:
    print(f"\n\n📋 {Fore.CYAN}جدول نهایی خروجی:{Style.RESET_ALL}")
    print(df_results.to_string(index=False))
