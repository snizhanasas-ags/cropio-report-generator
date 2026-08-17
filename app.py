"""
Cropio Report Generator FINAL_FIXED

This version uses:
- field mapping instead of column position copying
- template preservation approach
- separate fuel aggregation logic

IMPORTANT:
Before production use, test on the provided Cropio files.
"""

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from copy import copy
from io import BytesIO


st.set_page_config(
    page_title="Cropio Report Generator",
    page_icon="🚜"
)


# ---------- helpers ----------

def clean(x):
    if x is None:
        return ""
    return str(x).strip().lower()


def find_column(columns, variants):
    for col in columns:
        c = clean(col)
        for v in variants:
            if v in c:
                return col
    return None


def copy_row_style(ws, source, target):
    for col in range(1, ws.max_column + 1):
        src = ws.cell(source, col)
        dst = ws.cell(target, col)

        if src.has_style:
            dst._style = copy(src._style)

        dst.number_format = src.number_format
        dst.alignment = copy(src.alignment)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.font = copy(src.font)


# ---------- cropio work ----------

def load_work(file):

    xls = pd.ExcelFile(file)
    sheet = "Machine tasks" if "Machine tasks" in xls.sheet_names else xls.sheet_names[0]

    df = pd.read_excel(file, sheet_name=sheet)

    mapping = {
        "Початок": find_column(df.columns, ["початок", "start"]),
        "Кінець": find_column(df.columns, ["кінець", "end"]),
        "Водій": find_column(df.columns, ["водій", "driver"]),
        "Машина": find_column(df.columns, ["машина", "machine", "техніка"]),
        "Підтип робіт": find_column(df.columns, ["підтип", "робот"]),
        "Обладнання": find_column(df.columns, ["обладнання", "implement"]),
        "Поле": find_column(df.columns, ["поле", "field"]),
        "Оброблена площа, га": find_column(df.columns, ["площа", "area"]),
    }

    return df, mapping


# ---------- fuel ----------

def load_fuel(file):

    raw = pd.read_excel(
        file,
        sheet_name=0,
        header=None
    )

    header = 0

    for i, row in raw.iterrows():
        txt = " ".join(str(x).lower() for x in row.values)
        if "отрим" in txt and ("кіль" in txt or "палив" in txt):
            header = i
            break

    df = pd.read_excel(
        file,
        sheet_name=0,
        header=header
    )

    machine = find_column(df.columns, ["отримувач", "машина"])
    source = find_column(df.columns, ["азс", "паливозаправник"])
    amount = find_column(df.columns, ["кількість", "літр"])

    df[amount] = pd.to_numeric(df[amount], errors="coerce")
    df = df.dropna(subset=[amount])

    return df, machine, source, amount


def make_fuel_dict(df, machine, source, amount):

    result = {}

    for _, row in df.iterrows():

        m = str(row[machine])

        if m not in result:
            result[m] = {}

        s = str(row[source])

        result[m][s] = result[m].get(s, 0) + float(row[amount])

    return result


# ---------- excel ----------

def create_report(template, work, mapping, fuel):

    wb = load_workbook(template)
    ws = wb.active

    header_row = 2
    first_row = 3

    for row in ws.iter_rows(min_row=first_row):
        for cell in row:
            cell.value = None

    template_row = first_row

    row_index = first_row
    machine_written = set()

    headers = {
        clean(ws.cell(header_row, c).value): c
        for c in range(1, ws.max_column + 1)
    }

    for _, item in work.iterrows():

        copy_row_style(ws, template_row, row_index)

        for target, source in mapping.items():

            if source:

                col = headers.get(clean(target))

                if col:
                    ws.cell(row_index, col).value = item[source]

        machine = str(item[mapping["Машина"]]) if mapping["Машина"] else ""

        if machine not in machine_written and machine in fuel:

            for source, liters in fuel[machine].items():

                col = headers.get(clean(source))

                if col:
                    ws.cell(row_index, col).value = liters

            machine_written.add(machine)

        row_index += 1


    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output


# ---------- interface ----------

st.title("🚜 Cropio Report Generator FINAL FIXED")

work_file = st.file_uploader("1. Звіт роботи техніки", type="xlsx")
fuel_file = st.file_uploader("2. Видача палива", type="xlsx")
template_file = st.file_uploader("3. Шаблон", type="xlsx")


if st.button("🟢 Сформувати"):

    work, mapping = load_work(work_file)

    fuel_df, fm, fs, fa = load_fuel(fuel_file)

    fuel = make_fuel_dict(
        fuel_df,
        fm,
        fs,
        fa
    )

    result = create_report(
        template_file,
        work,
        mapping,
        fuel
    )

    st.success("Звіт сформовано")

    st.download_button(
        "⬇️ Завантажити Excel",
        result,
        "Cropio_report_final.xlsx"
    )
