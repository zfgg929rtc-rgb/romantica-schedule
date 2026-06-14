import streamlit as st
import pandas as pd
from datetime import datetime

# ------------------------------
# ФУНКЦІЇ ДЛЯ ОБРОБКИ ДАНИХ
# ------------------------------

def extract_waiter_rows(df):
    """
    Знаходить рядки, які належать офіціантам, ігноруючи менеджерів та хостів.
    """
    waiter_rows = []
    waiter_names = []
    found_waiter_header = False
    
    for idx, row in df.iterrows():
        first_cell = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ""
        if first_cell == "офіціант":
            found_waiter_header = True
            continue
        if found_waiter_header:
            # Зупиняємося, якщо почався блок менеджера або хоста
            if first_cell in ["менеджер", "хост", ""] or (len(first_cell) > 0 and first_cell[0] == "х"):
                break
            if first_cell and first_cell != "":
                waiter_rows.append(idx)
                waiter_names.append(str(row.iloc[0]).strip())
    return waiter_rows, waiter_names

def process_sheet(df, sheet_name, target_waiter="вітя"):
    """
    Обробляє окрему сторінку Excel-файлу, витягує дати та зміни офіціантів.
    """
    date_row_idx = 0
    while date_row_idx < df.shape[0]:
        first_val = df.iloc[date_row_idx, 0]
        # Перевірка, чи є в першій комірці дата
        if isinstance(first_val, (datetime, pd.Timestamp)) or pd.to_datetime(first_val, errors='coerce') is not pd.NaT:
            break
        date_row_idx += 1
        
    if date_row_idx >= df.shape[0]:
        return None
    
    dates = []
    for col in range(1, df.shape[1]):
        val = df.iloc[date_row_idx, col]
        if pd.isna(val):
            continue
        try:
            if isinstance(val, (datetime, pd.Timestamp)):
                d = val.date()
            else:
                d = pd.to_datetime(val).date()
            dates.append((col, d))
        except:
            continue
    
    if not dates:
        return None
    
    waiter_rows, waiter_names = extract_waiter_rows(df)
    if not waiter_names:
        return None
    
    days_data = []
    for col, date in dates:
        weekday = date.weekday() 
        weekday_name = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Нд"][weekday]
        waiter_count = 0
        shifts = {} 
        for i, name in zip(waiter_rows, waiter_names):
            val = df.iloc[i, col] if col < df.shape[1] else None
            if pd.notna(val) and str(val).strip() != "":
                waiter_count += 1
                shifts[name] = str(val).strip()
            else:
                shifts[name] = ""
                
        vita_shift = ""
        for name, shift in shifts.items():
            if target_waiter in name.lower():
                vita_shift = shift
                break
                
        days_data.append({
            "date": date,
            "weekday_name": weekday_name,
            "weekday": weekday,
            "waiter_count": waiter_count,
            "shifts": shifts,
            "vita_shift": vita_shift
        })
    
    return {
        "sheet_name": sheet_name,
        "days_data": days_data,
        "waiter_names": waiter_names
    }

# ------------------------------
# ІНТЕРФЕЙС STREAMLIT
# ------------------------------
st.set_page_config(page_title="Romantica Schedule", layout="wide")

st.title("📅 Зведений графік офіціантів Romantica")
st.write("Завантажте файл Excel, щоб побудувати інтерактивні таблиці та розрахувати статистику.")

# Елементи керування в інтерфейсі
uploaded_file = st.file_uploader("Виберіть файл Excel", type=["xlsx"])
target_employee = st.text_input("Ім'я працівника для аналізу статистики:", value="вітя")

if uploaded_file is not None:
    xl = pd.ExcelFile(uploaded_file)
    sheets_data = {}
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet, header=None, dtype=str)
        if not df.empty and df.shape[1] >= 2:
            sheets_data[sheet] = df

    all_results = []
    for sheet_name, df in sheets_data.items():
        result = process_sheet(df, sheet_name, target_waiter=target_employee.lower())
        if result:
            all_results.append(result)

    if not all_results:
        st.warning("Не знайдено жодних даних про офіціантів або дати не розпізнані.")
    else:
        st.success("Дані успішно оброблено!")
        
        # Змінні для підрахунку загальної статистики
        total_vita_work = 0
        total_vita_weekend = 0
        total_vita_off = 0
        total_days_all = 0
        other_stats = []
        total_sheets = 0

        # Відображення таблиць для кожної сторінки Excel
        for res in all_results:
            sheet_name = res["sheet_name"]
            days_data = res["days_data"]
            waiter_names = res["waiter_names"]
            
            st.subheader(f"📆 Сторінка: {sheet_name}")
            
            # Формуємо структуру таблиці (список словників)
            table_rows = []
            for day in days_data:
                row = {
                    "Дата": day["date"].strftime("%d.%m.%Y"),
                    "День": day["weekday_name"],
                    "К-сть офіціантів": day["waiter_count"]
                }
                # Додаємо зміну кожного офіціанта в окрему колонку
                for name in waiter_names:
                    row[name] = day["shifts"].get(name, "-")
                
                table_rows.append(row)
                
                # Накопичуємо статистику для цільового працівника
                if day["vita_shift"]:
                    total_vita_work += 1
                    if day["weekday"] >= 5:  # 5 = Субота, 6 = Неділя
                        total_vita_weekend += 1
                else:
                    total_vita_off += 1
                total_days_all += 1
            
            # Конвертуємо список у Pandas DataFrame
            df_report = pd.DataFrame(table_rows)
            
            # Виводимо інтерактивну таблицю Streamlit
            st.dataframe(df_report, use_container_width=True, hide_index=True)
            
            # Рахуємо середню кількість змін для інших
            all_shifts = {name: 0 for name in waiter_names}
            for day in days_data:
                for name in waiter_names:
                    if day["shifts"].get(name):
                        all_shifts[name] += 1
            
            avg_shifts = sum(all_shifts.values()) / len(waiter_names) if waiter_names else 0
            other_stats.append(avg_shifts)
            total_sheets += 1
        
        # ------------------------------
        # ВІДОБРАЖЕННЯ СТАТИСТИКИ
        # ------------------------------
        st.markdown("---")
        st.subheader(f"📊 Статистика для працівника: {target_employee.capitalize()}")
        
        # Створюємо гарну сітку з 4 колонок для показників
        col1, col2, col3, col4 = st.columns(4)
        col1.metric(label="✅ Робочі дні", value=total_vita_work)
        col2.metric(label="❌ Вихідні дні", value=total_vita_off)
        col3.metric(label="📆 Робочі у вихідні (сб/нд)", value=total_vita_weekend)
        col4.metric(label="🗓️ Всього днів у графіку", value=total_days_all)
        
        # Розрахунок середніх значень
        avg_others = sum(other_stats) / len(other_stats) if other_stats else 0
        avg_vita = total_vita_work / total_sheets if total_sheets else 0
        
        # Порівняльний блок
        st.info(f"""
        📈 **Порівняльний аналіз:**
        * Середня кількість змін на одного офіціанта (загалом): **{avg_others:.1f}**
        * Середня кількість змін у **{target_employee.capitalize()}**: **{avg_vita:.1f}**
        """)