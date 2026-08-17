
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from copy import copy
from io import BytesIO


st.set_page_config(page_title="Cropio Report Generator", page_icon="🚜")

st.title("🚜 Автоматичне формування звіту Cropio")


def find_column(df, variants):
    for col in df.columns:
        text = str(col).lower()
        for v in variants:
            if v.lower() in text:
                return col
    return None


def read_work_file(file):
    xls = pd.ExcelFile(file)
    df = pd.read_excel(file, sheet_name=xls.sheet_names[0])

    machine = find_column(df, ["машина", "техніка", "machine", "vehicle"])
    if not machine:
        raise Exception("Не знайдено колонку техніки у звіті роботи")

    return df, machine


def read_fuel_file(file):
    xls = pd.ExcelFile(file)

    for header in range(0, 20):
        df = pd.read_excel(file, sheet_name=xls.sheet_names[0], header=header)
        text = " ".join(map(str, df.columns)).lower()
        if "отрим" in text or "маш" in text or "тех" in text:
            break

    machine = find_column(df, ["отримувач", "машина", "техніка"])
    amount = find_column(df, ["кількість", "літр", "паливо"])
    source = find_column(df, ["джерело", "азс", "заправ"])

    if not machine or not amount:
        raise Exception("Не знайдено дані по паливу")

    return df, machine, amount, source


def prepare_fuel(df, machine, amount, source):
    df = df.copy()
    df[amount] = (
        df[amount]
        .astype(str)
        .str.replace(",", ".")
    )
    df[amount] = pd.to_numeric(df[amount], errors="coerce")
    df = df.dropna(subset=[machine, amount])

    if not source:
        df["Джерело"] = "Паливо"
        source = "Джерело"

    return (
        df.groupby([machine, source])[amount]
        .sum()
        .reset_index()
    )


def get_fuel_dict(fuel, machine_col, source_col, amount_col, machine):
    data = fuel[fuel[machine_col].astype(str) == str(machine)]

    result = {}
    for _, row in data.iterrows():
        result[str(row[source_col])] = row[amount_col]

    return result


def copy_style(ws, source_row, target_row):
    for col in range(1, ws.max_column + 1):
        a = ws.cell(source_row, col)
        b = ws.cell(target_row, col)

        if a.has_style:
            b._style = copy(a._style)


def create_report(work, fuel, template, work_machine,
                  fuel_machine, fuel_source, fuel_amount):

    wb = load_workbook(template)
    ws = wb.active

    start_row = 2
    style_row = ws.max_row

    # очищення старих даних
    for row in ws.iter_rows(min_row=start_row):
        for cell in row:
            cell.value = None

    row_num = start_row
    used = set()

    for _, row in work.iterrows():

        if row_num > ws.max_row:
            ws.insert_rows(row_num)

        copy_style(ws, style_row, row_num)

        values = list(row.values)

        for col, value in enumerate(values[:ws.max_column], 1):
            ws.cell(row_num, col).value = value

        machine = row[work_machine]

        if machine not in used:
            fuel_values = get_fuel_dict(
                fuel,
                fuel_machine,
                fuel_source,
                fuel_amount,
                machine
            )

            if fuel_values:
                for source, amount in fuel_values.items():
                    for c in range(1, ws.max_column + 1):
                        if str(ws.cell(1, c).value) == source:
                            cell = ws.cell(row_num, c)
                            cell.value = amount
                            cell.font = Font(color="FF0000")

            used.add(machine)

        row_num += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output


work_file = st.file_uploader("📄 1. Звіт роботи техніки Cropio", type="xlsx")
fuel_file = st.file_uploader("⛽ 2. Звіт видачі палива Cropio", type="xlsx")
template_file = st.file_uploader("📑 3. Ваш Excel-шаблон", type="xlsx")


if st.button("🟢 Сформувати звіт"):

    try:
        if not work_file or not fuel_file or not template_file:
            st.error("Завантажте всі три файли")
            st.stop()

        work, work_machine = read_work_file(work_file)
        fuel, fuel_machine, fuel_amount, fuel_source = read_fuel_file(fuel_file)

        fuel = prepare_fuel(
            fuel,
            fuel_machine,
            fuel_amount,
            fuel_source
        )

        result = create_report(
            work,
            fuel,
            template_file,
            work_machine,
            fuel_machine,
            fuel_source,
            fuel_amount
        )

        st.success("✅ Звіт сформовано")

        st.download_button(
            "⬇️ Завантажити Excel",
            result,
            "Cropio_report.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Помилка: {e}")
