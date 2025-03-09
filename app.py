import streamlit as st
import pandas as pd
import os

# 📌 تابع برای پردازش داده‌ها
@st.cache_data
def load_data(uploaded_file):
    if uploaded_file is not None:
        file_path = "uploaded_file.xlsx"
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # خواندن داده‌های اکسل
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
    return None, None, None, None

# 📌 رابط کاربری Streamlit
st.title("📊 محاسبه هزینه‌های تولید")

# ✅ آپلود فایل اکسل
uploaded_file = st.file_uploader("📂 لطفاً فایل اکسل را آپلود کنید", type=["xlsx"])
df_produce, df_daily, df_machine_overhead, df_raw_material = load_data(uploaded_file)

# ✅ دریافت ورودی از کاربر
user_input_kala = st.text_input("🔹 لطفاً کد جایگزین اصلی را وارد کنید:")
total_production = st.number_input("🔹 مقدار کل تولید خالص (کیلوگرم):", min_value=0.0, format="%.2f")
waste_price = st.number_input("🔹 قیمت هر کیلوگرم ضایعات:", min_value=0.0, format="%.2f")

# 🚀 دکمه پردازش
if st.button("🚀 محاسبه هزینه‌ها") and df_produce is not None:
    filtered_produce = df_produce[df_produce["کد جایگزین اصلی"] == user_input_kala]
    formulas_used = filtered_produce["کد فرمول"].unique()
    order_numbers = filtered_produce["شماره سفارش تولید"].unique()

    if len(formulas_used) == 0:
        st.error(f"⛔ کد کالا {user_input_kala} در داده‌ها پیدا نشد یا هیچ کد فرمولی برای آن ثبت نشده است!")
    else:
        st.success(f"✅ کد فرمول‌های مرتبط با {user_input_kala}: {formulas_used}")
        st.success(f"✅ شماره سفارش‌های مرتبط: {order_numbers}")

        # محاسبه بهای تمام‌شده برای هر کد فرمول و شماره سفارش
        for formula in formulas_used:
            for order in order_numbers:
                st.subheader(f"🔹 **محاسبه برای کد فرمول {formula} و شماره سفارش {order}:**")

                df_formula_data = df_produce[
                    (df_produce["کد فرمول"] == formula) & 
                    (df_produce["شماره سفارش تولید"] == order)
                ]

                if df_formula_data.empty:
                    st.warning(f"⛔ هیچ داده‌ای برای کد فرمول {formula} و شماره سفارش {order} یافت نشد!")
                    continue

                machine_names = df_formula_data["نام دستگاه"].unique()
                for machine_name in machine_names:
                    df_machine_data = df_formula_data[df_formula_data["نام دستگاه"] == machine_name]

                    total_clean_production = df_machine_data["khales"].astype(float).sum()
                    total_gross_production = df_machine_data["na_khales"].astype(float).sum()
                    total_waste = df_machine_data["وزن کل ضایعات (kg)"].astype(float).sum()
                    order_month = int(df_machine_data["ماه"].astype(float).mode()[0])

                    st.write(f"📊 **نام دستگاه: {machine_name}**")
                    st.write(f"📊 مقدار تولید خالص: {total_clean_production:.2f} کیلوگرم")
                    st.write(f"📊 مقدار تولید ناخالص: {total_gross_production:.2f} کیلوگرم")
                    st.write(f"📊 مقدار ضایعات: {total_waste:.2f} کیلوگرم")
                    st.write(f"📊 ماه تولید سفارش: {order_month}")

                    machine_overhead_row = df_machine_overhead[
                        (df_machine_overhead["نام دستگاه"] == machine_name) & 
                        (df_machine_overhead["ماه"] == order_month)
                    ]
                    machine_overhead_cost = machine_overhead_row["سربار"].values[0] if not machine_overhead_row.empty else 0

                    st.write(f"📊 هزینه سربار دستگاه {machine_name} در ماه {order_month}: 💰 {machine_overhead_cost}")

                    total_stoppage = df_machine_data["میزان کل توقف (دقیقه)"].astype(float).sum()
                    total_operation = df_machine_data["کارکرد (دقیقه)"].astype(float).sum()

                    st.write(f"📊 میزان توقفات: {total_stoppage} دقیقه")
                    st.write(f"📊 میزان کارکرد: {total_operation} دقیقه")

                    absorbed_overhead = (machine_overhead_cost * (total_stoppage + total_operation)) / (30 * 24 * 60)
                    st.write(f"📊 هزینه سربار جذب‌شده برای دستگاه {machine_name} در ماه {order_month}: 💰 {absorbed_overhead}")

                    waste_cost = total_waste * waste_price
                    total_material_cost = 0  

                    df_materials_info = df_daily[
                        (df_daily["کد فرمول"] == formula) & 
                        (df_daily["شماره سفارش"] == order)
                    ][["کد کالای مواد اولیه", "استفاده مواد اولیه"]]

                    if df_materials_info.empty:
                        st.warning(f"⛔ هیچ ماده اولیه‌ای برای این سفارش در DailyReport یافت نشد!")
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

                    total_cost = absorbed_overhead + waste_cost + total_material_cost
                    fi_total_clean = total_cost / total_clean_production if total_clean_production > 0 else 0

                    st.success(f"✅ **بهای تمام‌شده کل: 💰 {total_cost:.2f}**")
                    st.success(f"✅ **فی جذب شده خالص: {fi_total_clean:.2f}**")
