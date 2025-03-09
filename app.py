import streamlit as st
import pandas as pd

# 📌 تنظیمات اولیه صفحه
st.set_page_config(page_title="محاسبه بهای تمام‌شده", layout="wide")

# 📌 مسیر فایل اکسل
file_path = r"C:\Users\hosseini.am\Desktop\Report_costofproduce_backup - Copy.xlsx"

# 📌 خواندن داده‌های اکسل
@st.cache_data
def load_data():
    df_produce = pd.read_excel(file_path, sheet_name="dataofproduce", dtype=str, engine="openpyxl")
    df_daily = pd.read_excel(file_path, sheet_name="DailyReport", engine="openpyxl", header=1, dtype=str)
    df_machine_overhead = pd.read_excel(file_path, sheet_name="RawMaterialCost", usecols="AF:AJ", dtype=str, engine="openpyxl")
    df_raw_material = pd.read_excel(file_path, sheet_name="RawMaterialCost", dtype={"کد کالای مواد اولیه": str}, engine="openpyxl")

    # حذف فضای اضافی از نام ستون‌ها
    df_produce.columns = df_produce.columns.str.strip()
    df_daily.columns = df_daily.columns.str.strip()
    df_machine_overhead.columns = df_machine_overhead.columns.str.strip()
    df_raw_material.columns = df_raw_material.columns.str.strip()

    return df_produce, df_daily, df_machine_overhead, df_raw_material

df_produce, df_daily, df_machine_overhead, df_raw_material = load_data()

# 📌 گرفتن ورودی از کاربر
st.sidebar.header("🔹 ورود اطلاعات")
user_input_kala = st.sidebar.text_input("📌 کد جایگزین اصلی را وارد کنید:")
total_production = st.sidebar.number_input("📌 مقدار کل تولید خالص (KG):", min_value=0.0, value=1000.0, step=100.0)
waste_price = st.sidebar.number_input("📌 قیمت هر کیلوگرم ضایعات (ریال):", min_value=0, value=500000, step=50000)

# 📌 دکمه اجرا
if st.sidebar.button("🚀 محاسبه بهای تمام‌شده"):
    # 📌 فیلتر داده‌ها بر اساس کد کالا
    filtered_produce = df_produce[df_produce["کد جایگزین اصلی"] == user_input_kala]
    formulas_used = filtered_produce["کد فرمول"].unique()
    order_numbers = filtered_produce["شماره سفارش تولید"].unique()

    if len(formulas_used) == 0:
        st.error(f"⛔ کد کالا {user_input_kala} در داده‌ها پیدا نشد یا هیچ کد فرمولی برای آن ثبت نشده است!")
    else:
        st.success(f"✅ کد فرمول‌های مرتبط با {user_input_kala}: {list(formulas_used)}")
        st.success(f"✅ شماره سفارش‌های مرتبط: {list(order_numbers)}")

        # 📌 پردازش و نمایش خروجی
        output_data = []
        for formula in formulas_used:
            for order in order_numbers:
                df_formula_data = df_produce[
                    (df_produce["کد فرمول"] == formula) & 
                    (df_produce["شماره سفارش تولید"] == order)
                ]

                if df_formula_data.empty:
                    output_data.append([formula, order, "⛔ هیچ داده‌ای برای این ترکیب یافت نشد!"])
                    continue

                # استخراج اطلاعات تولید
                total_clean_production = df_formula_data["khales"].astype(float).sum()
                total_gross_production = df_formula_data["na_khales"].astype(float).sum()
                total_waste = df_formula_data["وزن کل ضایعات (kg)"].astype(float).sum()
                order_month = int(df_formula_data["ماه"].astype(float).mode()[0])
                machine_name = df_formula_data["نام دستگاه"].iloc[0]

                # دریافت هزینه سربار دستگاه در آن ماه
                machine_overhead_row = df_machine_overhead[
                    (df_machine_overhead["نام دستگاه"] == machine_name) & 
                    (df_machine_overhead["ماه"] == order_month)
                ]
                machine_overhead_cost = machine_overhead_row["سربار"].values[0] if not machine_overhead_row.empty else 0

                # محاسبه هزینه مواد اولیه
                df_materials_info = df_daily[
                    (df_daily["کد فرمول"] == formula) & 
                    (df_daily["شماره سفارش"] == order)
                ][["کد کالای مواد اولیه", "استفاده مواد اولیه"]]

                if not df_materials_info.empty:
                    df_materials_info = df_materials_info.merge(
                        df_raw_material[["کد کالای مواد اولیه", "قیمت واحد"]],
                        on="کد کالای مواد اولیه",
                        how="left"
                    )
                    df_materials_info["هزینه مواد اولیه"] = (
                        df_materials_info["استفاده مواد اولیه"].astype(float) * df_materials_info["قیمت واحد"].astype(float)
                    )
                    total_material_cost = df_materials_info["هزینه مواد اولیه"].sum()
                else:
                    total_material_cost = 0

                # محاسبه هزینه کل
                total_cost = total_material_cost + (total_waste * waste_price) + machine_overhead_cost
                fi_total_clean = total_cost / total_clean_production if total_clean_production > 0 else 0

                output_data.append([
                    formula, order, machine_name, total_clean_production, total_gross_production,
                    total_waste, order_month, machine_overhead_cost, total_material_cost, total_cost, fi_total_clean
                ])

        # 📌 نمایش خروجی در جدول
        df_output = pd.DataFrame(output_data, columns=[
            "کد فرمول", "شماره سفارش", "نام دستگاه", "تولید خالص (KG)", "تولید ناخالص (KG)",
            "ضایعات (KG)", "ماه تولید", "سربار دستگاه", "هزینه مواد اولیه", "هزینه کل", "فی جذب شده خالص"
        ])
        st.dataframe(df_output)

        # 📌 امکان دانلود خروجی به عنوان اکسل
        @st.cache_data
        def convert_df(df):
            return df.to_csv(index=False).encode("utf-8")

        csv = convert_df(df_output)
        st.download_button(
            label="📥 دانلود خروجی به عنوان CSV",
            data=csv,
            file_name="output.csv",
            mime="text/csv"
        )
