import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(
    page_title="Cropio Report Generator",
    page_icon="🚜"
)

st.title("🚜 Автоматичне формування звіту Cropio")

st.write(
    "Завантажте два звіти Cropio та готовий Excel-шаблон"
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
    "3. Готовий шаблон звіту",
    type=["xlsx"]
)


if st.button("🟢 Сформувати звіт"):

    if not work_file or not fuel_file or not template_file:
        st.error("Потрібно завантажити всі три файли")

    else:
        st.success("Файли завантажені. Можна формувати звіт.")

        st.write("Файл роботи:", work_file.name)
        st.write("Файл палива:", fuel_file.name)
        st.write("Шаблон:", template_file.name)
