"""
Cropio Report Generator FINAL v1
"""

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from copy import copy
from io import BytesIO


st.set_page_config(
    page_title="Cropio Report Generator",
    page_icon="🚜",
    layout="wide"
)


def normalize(value):
    if value is None:
        return ""
    return str(value).strip().lower()


def find_column(columns, variants):
    for col in columns:
        txt = normalize(col)
        for v in variants:
            if v in txt:
                return col
    return None


def find_header(file, sheet, max_rows=30):
    raw = pd.read_excel(
        file,
        sheet_name=sheet,
        header=None
    )

    for row in range(min(max_rows, len(raw))):
        values = " ".join(
            normalize(x) for x in raw.iloc[row].tolist()
        )

        if (
            "маш" in values
            or "тех" in values
            or "отрим" in values
            or "водій" in values
        ):
            return row

    return 0


def read_work(file):

    xls = pd.ExcelFile(file)

    sheet = "Machine tasks" if "Machine tasks" in xls.sheet_names else xls.sheet_names[0]

    header = find_header(file, sheet)

    df = pd.read_excel(
        file,
        sheet_name=sheet,
        header=header
    )

    machine = find_column(
        df.columns,
        [
            "маш",
            "тех",
            "vehicle",
            "machine",
            "агрег"
        ]
    )

    if not machine:
        raise Exception(
            "Не знайдено колонку техніки у звіті роботи"
        )

    return df, machine


def read_fuel(file):

    xls = pd.ExcelFile(file)

    sheet = xls.sheet_names[0]

    header = find_header(file, sheet)

    df = pd.read_excel(
        file,
        sheet_name=sheet,
        header=header
    )

    machine = find_column(
        df.columns,
        [
            "отрим",
            "маш",
            "тех"
        ]
    )

    source = find_column(
        df.columns,
        [
            "азс",
            "паливозаправ",
            "джерело"
        ]
    )

    amount = find_column(
        df.columns,
        [
            "кіль",
            "літр",
            "палив"
        ]
    )

    if not machine or not amount:
        raise Exception(
            "Не знайдено структуру палива"
        )

    df[amount] = pd.to_numeric(
        df[amount],
        errors="coerce"
    )

    df = df.dropna(
        subset=[amount]
    )

    return df, machine, source, amount


def copy_style(ws, src_row, dst_row):

    for col in range(1, ws.max_column + 1):

        src = ws.cell(src_row, col)
        dst = ws.cell(dst_row, col)

        if src.has_style:
            dst._style = copy(src._style)

        dst.number_format = src.number_format
        dst.alignment = copy(src.alignment)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.font = copy(src.font)


def fuel_summary(fuel, machine_col, source_col, amount_col):

    result = {}

    for _, row in fuel.iterrows():

        machine = str(row[machine_col])

        if machine not in result:
            result[machine] = {}

        source = str(row[source_col]) if source_col else "Паливо"

        result[machine][source] = (
            result[machine].get(source, 0)
            +
            float(row[amount_col])
        )

    return result


def generate_report(template, work, fuel_data, work_machine):

    wb = load_workbook(template)

    ws = wb.active

    header_row = 2
    first_data_row = 3

    style_row = first_data_row

    # clear only data area
    for row in ws.iter_rows(min_row=first_data_row):
        for cell in row:
            cell.value = None

    current = first_data_row
    written = set()

    for _, row in work.iterrows():

        if current > ws.max_row:
            ws.insert_rows(current)

        copy_style(
            ws,
            style_row,
            current
        )

        values = list(row.values)

        for col, value in enumerate(values, start=1):

            if col <= ws.max_column:
                ws.cell(current, col).value = value

        machine = str(row[work_machine])

        if machine not in written:

            if machine in fuel_data:

                for source, amount in fuel_data[machine].items():

                    for cell in ws[header_row]:

                        if normalize(cell.value) == normalize(source):

                            target = ws.cell(
                                current,
                                cell.column
                            )

                            target.value = amount
                            target.font = Font(
                                color="FF0000"
                            )

            written.add(machine)

        current += 1


    # Add machines existing only in fuel
    existing = set(
        str(x)
        for x in work[work_machine].tolist()
    )

    for machine, fuels in fuel_data.items():

        if machine not in existing:

            copy_style(
                ws,
                style_row,
                current
            )

            ws.cell(
                current,
                4
            ).value = machine

            for source, amount in fuels.items():

                for cell in ws[header_row]:

                    if normalize(cell.value) == normalize(source):

                        ws.cell(
                            current,
                            cell.column
                        ).value = amount

                        ws.cell(
                            current,
                            cell.column
                        ).font = Font(
                            color="FF0000"
                        )

            current += 1


    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output


st.title("🚜 Cropio Report Generator FINAL")


work_file = st.file_uploader(
    "1. Звіт роботи техніки",
    type=["xlsx"]
)

fuel_file = st.file_uploader(
    "2. Видача палива",
    type=["xlsx"]
)

template_file = st.file_uploader(
    "3. Excel шаблон",
    type=["xlsx"]
)


if st.button("🟢 Сформувати звіт"):

    try:

        work, work_machine = read_work(work_file)

        fuel, fuel_machine, fuel_source, fuel_amount = read_fuel(fuel_file)

        fuels = fuel_summary(
            fuel,
            fuel_machine,
            fuel_source,
            fuel_amount
        )

        result = generate_report(
            template_file,
            work,
            fuels,
            work_machine
        )

        st.success(
            f"✅ Звіт сформовано. Робіт: {len(work)}"
        )

        st.download_button(
            "⬇️ Завантажити готовий Excel",
            result,
            "Cropio_final_report.xlsx"
        )

    except Exception as e:

        st.error(
            str(e)
        )
