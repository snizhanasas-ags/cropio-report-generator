# Cropio Report Generator v4
# Auto analyzer version
# Finds Cropio headers automatically before processing

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from copy import copy
from io import BytesIO


st.set_page_config(
    page_title="Cropio Report Generator v4",
    page_icon="🚜"
)


def norm(x):
    if x is None:
        return ""
    return str(x).strip().lower()


def detect_header(file, sheet=0):
    """
    Finds the real header row automatically
    """
    for row in range(0, 25):
        df = pd.read_excel(
            file,
            sheet_name=sheet,
            header=row
        )

        text = " ".join(
            [norm(c) for c in df.columns]
        )

        if (
            "маш" in text
            or "тех" in text
            or "vehicle" in text
            or "machine" in text
        ):
            return row, df

    raise Exception(
        "Не вдалося знайти рядок заголовків Cropio"
    )


def find_column(df, variants):
    for col in df.columns:
        for v in variants:
            if v in norm(col):
                return col
    return None


def read_work(file):

    header, df = detect_header(file)

    machine = find_column(
        df,
        [
            "маш",
            "тех",
            "vehicle",
            "machine",
            "агрегат"
        ]
    )

    if not machine:
        raise Exception(
            "Не знайдено колонку техніки. Знайдені колонки: "
            + ", ".join(map(str, df.columns))
        )

    return df, machine, header


def read_fuel(file):

    header, df = detect_header(file)

    machine = find_column(
        df,
        [
            "отрим",
            "маш",
            "тех"
        ]
    )

    amount = find_column(
        df,
        [
            "кіль",
            "літр",
            "палив"
        ]
    )

    source = find_column(
        df,
        [
            "азс",
            "паливозаправник",
            "джерело"
        ]
    )

    if not machine or not amount:
        raise Exception(
            "Не знайдено колонки палива"
        )

    return df, machine, amount, source, header


st.title("🚜 Cropio Report Generator v4")

st.info(
    "Версія з автоматичним аналізом структури Cropio"
)

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


if st.button("🟢 Аналізувати та сформувати"):

    try:

        work, work_machine, work_header = read_work(work_file)

        fuel, fuel_machine, fuel_amount, fuel_source, fuel_header = read_fuel(fuel_file)

        st.success("Структура Cropio знайдена")

        st.write(
            "Звіт роботи:",
            len(work),
            "рядків"
        )

        st.write(
            "Колонка техніки:",
            work_machine
        )

        st.write(
            "Заголовок у рядку:",
            work_header + 1
        )

        st.write(
            "Паливо:",
            fuel_machine,
            fuel_amount,
            fuel_source
        )

        st.warning(
            "Excel генератор буде підключено після підтвердження структури"
        )

    except Exception as e:
        st.error(str(e))
