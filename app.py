import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from copy import copy
from io import BytesIO


st.set_page_config(
    page_title="Cropio Report Generator",
    page_icon="🚜",
    layout="wide"
)


st.title("🚜 Автоматичне формування звіту Cropio")

st.write(
    """
    Завантажте:
    1. Звіт роботи техніки Cropio
    2. Звіт видачі палива Cropio
    3. Ваш Excel-шаблон готового звіту
    """
)


work_file = st.file_uploader(
    "📄 1. Звіт роботи техніки Cropio",
    type=["xlsx"]
)


fuel_file = st.file_uploader(
    "⛽ 2. Звіт видачі палива Cropio",
    type=["xlsx"]
)


template_file = st.file_uploader(
    "📑 3. Готовий шаблон звіту",
    type=["xlsx"]
)
def find_column(df, variants):
    """
    Пошук колонки за ключовими словами
    """
    for col in df.columns:
        name = str(col).lower().strip()

        for variant in variants:
            if variant.lower() in name:
                return col

    return None


def read_cropio_work(file):
    """
    Читання звіту роботи техніки
    """

    excel = pd.ExcelFile(file)

    # Беремо перший лист, якщо структура стандартна
    sheet = excel.sheet_names[0]

    df = pd.read_excel(
        file,
        sheet_name=sheet
    )


    machine = find_column(
        df,
        ["машина", "техніка", "vehicle", "machine"]
    )

    if not machine:
        raise Exception(
            "Не знайдено колонку техніки у звіті роботи"
        )


    return df, machine



def read_cropio_fuel(file):
    """
    Читання звіту палива
    """

    excel = pd.ExcelFile(file)

    sheet = excel.sheet_names[0]


    # пробуємо різні рядки заголовків
    for header in range(0, 15):

        df = pd.read_excel(
            file,
            sheet_name=sheet,
            header=header
        )

        cols = " ".join(
            [str(x).lower() for x in df.columns]
        )


        if (
            "отрим" in cols
            or "машин" in cols
            or "технік" in cols
        ):
            break


    machine = find_column(
        df,
        [
            "отримувач",
            "машина",
            "техніка",
            "machine"
        ]
    )


    amount = find_column(
        df,
        [
            "кількість",
            "літр",
            "паливо",
            "amount"
        ]
    )


    source = find_column(
        df,
        [
            "джерело",
            "азс",
            "місце",
            "заправ"
        ]
    )


    if not machine:
        raise Exception(
            "Не знайдено техніку у звіті палива"
        )


    return df, machine, amount, source
