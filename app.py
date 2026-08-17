"""
Cropio Report System v2.0
Production structure foundation.

Modules inside one file:
- UI
- File validation
- Cropio work reader
- Fuel reader
- Template processor
- Excel generator
"""

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from copy import copy
from io import BytesIO


# ==========================
# CONFIG
# ==========================

st.set_page_config(
    page_title="Cropio Report System",
    page_icon="🚜",
    layout="wide"
)


# ==========================
# HELPERS
# ==========================

def normalize(value):
    if value is None:
        return ""
    return str(value).strip().lower()


def find_column(df, variants):
    for col in df.columns:
        col_name = normalize(col)
        for item in variants:
            if item.lower() in col_name:
                return col
    return None


# ==========================
# CROPIO WORK
# ==========================

def load_work_report(file):

    xls = pd.ExcelFile(file)
    df = pd.read_excel(
        file,
        sheet_name=xls.sheet_names[0]
    )

    machine = find_column(
        df,
        ["машина", "техніка", "machine", "vehicle"]
    )

    if not machine:
        raise Exception(
            "У звіті роботи не знайдено колонку техніки"
        )

    return df, machine


# ==========================
# FUEL ENGINE
# ==========================

def load_fuel_report(file):

    xls = pd.ExcelFile(file)

    result = None

    for header in range(0, 20):

        temp = pd.read_excel(
            file,
            sheet_name=xls.sheet_names[0],
            header=header
        )

        columns = " ".join(
            [str(x).lower() for x in temp.columns]
        )

        if "отрим" in columns or "маш" in columns:
            result = temp
            break

    if result is None:
        raise Exception(
            "Не знайдено таблицю видачі палива"
        )

    machine = find_column(
        result,
        ["отримувач", "машина", "техніка"]
    )

    amount = find_column(
        result,
        ["кількість", "літр", "паливо"]
    )

    source = find_column(
        result,
        ["азс", "джерело", "паливозаправник", "заправ"]
    )

    if not machine or not amount:
        raise Exception(
            "Не знайдені колонки палива"
        )

    return result, machine, amount, source


def prepare_fuel(df, machine, amount, source):

    df = df.copy()

    df[amount] = (
        df[amount]
        .astype(str)
        .str.replace(",", ".")
    )

    df[amount] = pd.to_numeric(
        df[amount],
        errors="coerce"
    )

    df = df.dropna(
        subset=[machine, amount]
    )

    if source is None:
        df["Джерело"] = "Паливо"
        source = "Джерело"

    return (
        df.groupby([machine, source])[amount]
        .sum()
        .reset_index(),
        source
    )


# ==========================
# EXCEL TEMPLATE ENGINE
# ==========================

def copy_style(ws, source_row, target_row):

    for col in range(1, ws.max_column + 1):

        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)

        if src.has_style:
            dst._style = copy(src._style)

        dst.number_format = src.number_format


def generate_excel(
        template,
        work,
        fuel,
        work_machine,
        fuel_machine,
        fuel_amount,
        fuel_source
):

    wb = load_workbook(template)
    ws = wb.active

    start_row = 2
    style_row = ws.max_row

    for row in ws.iter_rows(min_row=start_row):
        for cell in row:
            cell.value = None

    row_index = start_row
    used = set()

    for _, row in work.iterrows():

        if row_index > ws.max_row:
            ws.insert_rows(row_index)

        copy_style(
            ws,
            style_row,
            row_index
        )

        for col, value in enumerate(row.tolist(), 1):

            if col <= ws.max_column:
                ws.cell(
                    row_index,
                    col
                ).value = value

        machine = row[work_machine]

        key = normalize(machine)

        if key not in used:

            fuel_rows = fuel[
                fuel[fuel_machine]
                .astype(str)
                ==
                str(machine)
            ]

            for _, fuel_row in fuel_rows.iterrows():

                source = str(
                    fuel_row[fuel_source]
                )

                amount = fuel_row[fuel_amount]

                for cell in ws[1]:

                    if normalize(cell.value) == normalize(source):

                        target = ws.cell(
                            row_index,
                            cell.column
                        )

                        target.value = amount
                        target.font = Font(
                            color="FF0000"
                        )

            used.add(key)

        row_index += 1


    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output


# ==========================
# USER INTERFACE
# ==========================

st.title("🚜 Cropio Report System v2.0")

st.write(
    "Завантажте звіт роботи, паливо та Excel-шаблон"
)


work_file = st.file_uploader(
    "1. Звіт роботи техніки Cropio",
    type=["xlsx"]
)

fuel_file = st.file_uploader(
    "2. Звіт видачі палива Cropio",
    type=["xlsx"]
)

template_file = st.file_uploader(
    "3. Ваш Excel шаблон",
    type=["xlsx"]
)


if st.button("🟢 Сформувати звіт"):

    try:

        if not work_file or not fuel_file or not template_file:
            st.warning(
                "Потрібно завантажити всі три файли"
            )
            st.stop()

        work, work_machine = load_work_report(
            work_file
        )

        fuel, fuel_machine, fuel_amount, fuel_source = load_fuel_report(
            fuel_file
        )

        fuel, fuel_source = prepare_fuel(
            fuel,
            fuel_machine,
            fuel_amount,
            fuel_source
        )

        result = generate_excel(
            template_file,
            work,
            fuel,
            work_machine,
            fuel_machine,
            fuel_amount,
            fuel_source
        )

        st.success(
            f"✅ Готово. Робіт знайдено: {len(work)}"
        )

        st.download_button(
            "⬇️ Завантажити Excel",
            result,
            "Cropio_report.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(
            f"Помилка: {e}"
        )
