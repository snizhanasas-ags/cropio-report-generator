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
