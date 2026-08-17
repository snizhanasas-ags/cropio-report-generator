# Cropio Report Generator v3
# Production version foundation
# This file is prepared for the real Cropio workflow:
# - work report
# - fuel report
# - Excel template preservation

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from copy import copy
from io import BytesIO


st.set_page_config(
    page_title="Cropio Report Generator v3",
    page_icon="🚜"
)


def norm(v):
    if v is None:
        return ""
    return str(v).strip().lower()


def find_col(columns, names):
    for c in columns:
        for n in names:
            if n in norm(c):
                return c
    return None


def read_work(file):
    df = pd.read_excel(
        file,
        sheet_name="Machine tasks",
        header=1
    )

    machine = find_col(
        df.columns,
        ["машина", "machine"]
    )

    if not machine:
        raise Exception("Не знайдено техніку у звіті роботи")

    return df, machine


def read_fuel(file):
    df = pd.read_excel(
        file,
        sheet_name=0,
        header=8
    )

    machine = find_col(
        df.columns,
        ["отримувач"]
    )

    source = find_col(
        df.columns,
        ["азс", "паливозаправник"]
    )

    amount = find_col(
        df.columns,
        ["кількість"]
    )

    if not machine or not amount:
        raise Exception("Не знайдено дані палива")

    df[amount] = pd.to_numeric(
        df[amount],
        errors="coerce"
    )

    df = df.dropna(subset=[amount])

    fuel = (
        df.groupby([machine, source])[amount]
        .sum()
        .reset_index()
    )

    return fuel, machine, source, amount


def copy_row(ws, source, target):
    for c in range(1, ws.max_column + 1):
        a = ws.cell(source, c)
        b = ws.cell(target, c)

        if a.has_style:
            b._style = copy(a._style)

        b.number_format = a.number_format


def build_report(template, work, fuel, work_machine,
                 fuel_machine, fuel_source, fuel_amount):

    wb = load_workbook(template)
    ws = wb.active

    header_row = 2
    start_row = 3

    style_row = start_row

    # clear old data only
    for row in ws.iter_rows(min_row=start_row):
        for cell in row:
            cell.value = None

    current = start_row
    written_fuel = set()

    for _, row in work.iterrows():

        if current > ws.max_row:
            ws.insert_rows(current)

        copy_row(ws, style_row, current)

        for i, value in enumerate(row.tolist(), 1):
            if i <= ws.max_column:
                ws.cell(current, i).value = value

        machine = row[work_machine]

        if norm(machine) not in written_fuel:

            f = fuel[
                fuel[fuel_machine].astype(str)
                ==
                str(machine)
            ]

            for _, fr in f.iterrows():

                # Find existing fuel columns only
                source_name = str(fr[fuel_source])

                for cell in ws[header_row]:
                    if norm(cell.value) == norm(source_name):
                        target = ws.cell(
                            current,
                            cell.column
                        )
                        target.value = fr[fuel_amount]
                        target.font = Font(color="FF0000")

            written_fuel.add(norm(machine))

        current += 1


    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output


st.title("🚜 Cropio Report Generator v3")

work_file = st.file_uploader(
    "1. Звіт роботи техніки",
    type="xlsx"
)

fuel_file = st.file_uploader(
    "2. Видача палива",
    type="xlsx"
)

template_file = st.file_uploader(
    "3. Excel шаблон",
    type="xlsx"
)


if st.button("🟢 Сформувати звіт"):

    try:
        work, work_machine = read_work(work_file)

        fuel, fuel_machine, fuel_source, fuel_amount = read_fuel(fuel_file)

        result = build_report(
            template_file,
            work,
            fuel,
            work_machine,
            fuel_machine,
            fuel_source,
            fuel_amount
        )

        st.success(
            f"Готово. Робіт: {len(work)}"
        )

        st.download_button(
            "⬇️ Завантажити Excel",
            result,
            "Cropio_report_v3.xlsx"
        )

    except Exception as e:
        st.error(str(e))
