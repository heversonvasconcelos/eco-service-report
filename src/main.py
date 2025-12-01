# main.py
import datetime
import io
import os
import re
import zipfile
from typing import Optional

import pandas as pd
import streamlit as st
from dateutil.relativedelta import relativedelta
from jinja2 import Template

# --- I18N Translations ---
TRANSLATIONS = {
    "en": {
        "page_title": "Eco Service Reports", "main_title": "♻️ Eco Service Report Generator",
        "main_subtitle": "Upload your customer data (XLS/XLSX) to generate monthly environmental impact reports.",
        "view_format_expander": "View Expected Spreadsheet Format",
        "format_intro": "The application expects an Excel file with a two-level header structure where the month/year acts as a parent header for its corresponding data columns.",
        "format_header1": "- **First Header Row**: This row defines the months. It should be placed **directly above the first column of that month's data** (e.g., above `1st WEEK`). The month can be a date (`01/01/2025`) or text (`january of 2025`).",
        "format_header2": "- **Second Header Row**: This contains the actual column titles like `ID`, `FULL NAME`, `1st WEEK`, and `TOTAL`.",
        "format_columns": "The columns `ID` and `FULL NAME` should appear before the month-specific columns. The app will automatically find them, and also the correct `TOTAL` column for the month you select.",
        "format_example_title": "**Example Layout:**",
        "format_example_content": """
|      | ... | FULL NAME     | 01/01/2025 (or "january of 2025") |           |           |           | 01/02/2025 ...
|------|-----|---------------|-----------------------------------|-----------|-----------|-----------|----------------
| ID   | ... | FULL NAME     | 1st WEEK                          | 2nd WEEK  | ...       | TOTAL     | 1st WEEK       ...
|------|-----|---------------|-----------------------------------|-----------|-----------|-----------|----------------
| 1    | ... | Client A      | 30.0                              | 40.0      | ...       | 150.5     | 40.0           ...
| 2    | ... | Client B      | 50.0                              | 50.0      | ...       | 200.0     | 55.0           ...
        """,
        "upload_label": "Upload Spreadsheet", "file_uploaded_success": "File uploaded: {file_name}",
        "process_error": "An error occurred while processing the spreadsheet: {error}",
        "no_month_header_error": "Could not find any valid month headers in the first row. Please ensure months are formatted as 'DD/MM/YYYY' or 'Month of Year' (e.g., 'january of 2025').",
        "select_month_header": "1. Select Report Month",
        "select_month_label": "Which month do you want to generate reports for?",
        "map_columns_header": "2. Map Spreadsheet Columns",
        "map_columns_caption": "The columns for ID, Name, and Total are detected automatically based on their default names.",
        "customer_id_label": "Customer ID Column", "customer_name_label": "Customer Name Column",
        "waste_total_label": "Waste Total Column for {month_name}",
        "review_data_header": "3. Review Data and Generate Reports",
        "review_data_caption": "Review the processed data below. Select a row to preview its report.",
        "preview_header": "Preview: Report for {customer_name}",
        "prepare_reports_button": "Prepare All Reports (.zip)", "download_reports_button": "Download All Reports",
        "zip_filename": "{month_name}-Reports.zip",
        "no_data_warning": "No valid data found for the selected columns. Please check your file and selections.",
        "months": ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
        "month_of_year_format": "{month_name} OF {year}", "month_year_format": "%B-%Y", "last_month_na": "N/A",
        "report_id_na": "N/A",
    },
    "pt": {
        "page_title": "Relatórios Eco Service", "main_title": "♻️ Gerador de Relatórios Eco Service",
        "main_subtitle": "Faça o upload dos dados dos seus clientes (XLS/XLSX) para gerar relatórios mensais de impacto ambiental.",
        "view_format_expander": "Ver Formato da Planilha Esperado",
        "format_intro": "A aplicação espera um arquivo Excel com uma estrutura de cabeçalho de dois níveis, onde o mês/ano atua como um cabeçalho pai para suas colunas de dados correspondentes.",
        "format_header1": "- **Primeira Linha do Cabeçalho**: Esta linha deve definir os meses. Deve ser colocada **diretamente acima da primeira coluna de dados daquele mês** (ex: acima de `1º SEMANA`). O mês pode ser uma data (`01/01/2025`) ou texto (`janeiro de 2025`).",
        "format_header2": "- **Segunda Linha do Cabeçalho**: Esta contém os títulos reais das colunas, como `ID`, `NOME COMPLETO`, `1º SEMANA` e `TOTAL`.",
        "format_columns": "As colunas `ID` e `NOME COMPLETO` devem aparecer antes das colunas específicas do mês. A aplicação as encontrará automaticamente, assim como a coluna `TOTAL` correta para o mês que você selecionar.",
        "format_example_title": "**Exemplo de Layout:**",
        "format_example_content": """
|      | ... | NOME COMPLETO | 01/01/2025 (ou "janeiro de 2025") |           |           |           | 01/02/2025 ...
|------|-----|---------------|-----------------------------------|-----------|-----------|-----------|----------------
| ID   | ... | NOME COMPLETO | 1º SEMANA                         | 2º SEMANA | ...       | TOTAL     | 1º SEMANA      ...
|------|-----|---------------|-----------------------------------|-----------|-----------|-----------|----------------
| 1    | ... | Cliente A     | 30.0                              | 40.0      | ...       | 150.5     | 40.0           ...
| 2    | ... | Cliente B     | 50.0                              | 50.0      | ...       | 200.0     | 55.0           ...
        """,
        "upload_label": "Fazer Upload da Planilha", "file_uploaded_success": "Arquivo enviado: {file_name}",
        "process_error": "Ocorreu um erro ao processar a planilha: {error}",
        "no_month_header_error": "Não foi possível encontrar cabeçalhos de mês válidos na primeira linha. Por favor, garanta que os meses estejam formatados como 'DD/MM/AAAA' ou 'Mês de Ano' (ex: 'janeiro de 2025').",
        "select_month_header": "1. Selecione o Mês do Relatório",
        "select_month_label": "Para qual mês você deseja gerar os relatórios?",
        "map_columns_header": "2. Mapeie as Colunas da Planilha",
        "map_columns_caption": "As colunas de ID, Nome e Total são detectadas automaticamente com base em seus nomes padrão.",
        "customer_id_label": "Coluna de ID do Cliente", "customer_name_label": "Coluna de Nome do Cliente",
        "waste_total_label": "Coluna de Total de Resíduos para {month_name}",
        "review_data_header": "3. Revise os Dados e Gere os Relatórios",
        "review_data_caption": "Revise os dados processados abaixo. Selecione uma linha para pré-visualizar seu relatório.",
        "preview_header": "Pré-visualização: Relatório para {customer_name}",
        "prepare_reports_button": "Preparar Todos os Relatórios (.zip)",
        "download_reports_button": "Baixar Todos os Relatórios", "zip_filename": "Relatorios-{month_name}.zip",
        "no_data_warning": "Nenhum dado válido encontrado para as colunas selecionadas. Por favor, verifique seu arquivo e seleções.",
        "months": ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"],
        "month_of_year_format": "{month_name} DE {year}", "month_year_format": "%B-%Y", "last_month_na": "N/D",
        "report_id_na": "N/D",
    }
}

# --- Constants ---
DEFAULT_XLS_COLUMNS = {"id": "ID", "name": "NOME COMPLETO", "total": "TOTAL"}
TEMPLATE_FILE_HTML = 'templates/service-report-preview-template.html'
TEMPLATE_FILE_SVG = 'templates/service-report-template.svg'
FERTILIZER_FACTOR, CO2_AVOIDED_FACTOR, DRIVING_DISTANCE_CO2_CONVERSION_FACTOR, TREES_EQUIVALENT_CO2_ABSORPTION_FACTOR, WATER_LITERS_FACTOR_BASE, WATER_LITERS_MULTIPLIER = 0.38, 0.77, 0.096, 0.35, 0.214, 12
MONTH_PT_TO_NUM = {name.lower(): i + 1 for i, name in enumerate(TRANSLATIONS["pt"]["months"])}

# --- Date Handling ---
def format_local_date(date_obj: datetime.datetime, lang: str, t) -> str:
    t_months = TRANSLATIONS[lang]["months"]
    month_name = t_months[date_obj.month - 1]
    return t("month_of_year_format").format(month_name=month_name, year=date_obj.year).upper()

def parse_month_string(month_str: str) -> Optional[datetime.datetime]:
    month_str_lower = month_str.lower()
    for separator in [" de ", " of "]:
        if separator in month_str_lower:
            parts = month_str_lower.split(separator)
            if len(parts) == 2:
                month_name, year_str = parts
                month_num = next((i + 1 for i, m in enumerate(TRANSLATIONS["en"]["months"]) if m.lower() == month_name), None)
                if not month_num:
                    month_num = next((i + 1 for i, m in enumerate(TRANSLATIONS["pt"]["months"]) if m.lower() == month_name), None)
                if month_num:
                    try: return datetime.datetime(int(year_str), month_num, 1)
                    except (ValueError, TypeError): continue
    return None

# --- File Handling and Data Processing ---
def upload_xls_file(t) -> Optional[st.runtime.uploaded_file_manager.UploadedFile]:
    uploaded_file = st.file_uploader(t("upload_label"), type=["xls", "xlsx"])
    if uploaded_file:
        st.success(t("file_uploaded_success").format(file_name=uploaded_file.name))
        return uploaded_file
    return None

def get_default_index(col_name_key: str, available_columns: list[str]) -> int:
    default_col_name = DEFAULT_XLS_COLUMNS.get(col_name_key)
    if not default_col_name: return 0
    if col_name_key == 'total':
        for i, col in enumerate(available_columns):
            if col.startswith(default_col_name): return i
    elif default_col_name in available_columns:
        return available_columns.index(default_col_name)
    return 0

# --- Report Generation ---
def create_report_html(template_path: str, customer_data: pd.Series, report_month_str: str, lang: str, t) -> str:
    dir_path = os.path.dirname(os.path.realpath(__file__))
    template_path = os.path.join(dir_path, template_path)
    with open(template_path, 'r', encoding='utf-8') as f:
        jinja_template = Template(f.read())

    report_date = parse_month_string(report_month_str)
    if report_date:
        last_month_date = report_date - relativedelta(months=1)
        last_month_year_str = format_local_date(last_month_date, lang, t)
        report_id_date_str = report_date.strftime('%Y-%m')
    else:
        last_month_year_str, report_id_date_str = t("last_month_na"), t("report_id_na")

    report_id = f"{customer_data['customer_id']}-{report_id_date_str}"
    customer_total_float = customer_data['customer_total']
    fertilizer_kg, co2_avoided = customer_total_float * FERTILIZER_FACTOR, customer_total_float * CO2_AVOIDED_FACTOR
    driving_distance, trees_equivalent, water_liters = co2_avoided / DRIVING_DISTANCE_CO2_CONVERSION_FACTOR, co2_avoided / TREES_EQUIVALENT_CO2_ABSORPTION_FACTOR, customer_total_float * WATER_LITERS_FACTOR_BASE * WATER_LITERS_MULTIPLIER

    template_vars = {
        'current_month_year': report_month_str.upper(), 'last_month_year': last_month_year_str, 'report_id': report_id,
        'customer_name': customer_data['customer_name'], 'customer_waste_kg': f"{customer_total_float:.2f}",
        'fertilizer_kg': f"{fertilizer_kg:.2f}", 'co2_avoided': f"{co2_avoided:.2f}",
        'driving_distance': f"{driving_distance:.2f}", 'trees_equivalent': f"{trees_equivalent:.0f}",
        'water_liters': f"{water_liters:.2f}"
    }
    return jinja_template.render(template_vars)

@st.cache_data
def generate_single_report_preview(customer_data: pd.Series, report_month_str: str, lang: str) -> str:
    def t(key): return TRANSLATIONS[lang].get(key, key)
    return create_report_html(TEMPLATE_FILE_HTML, customer_data, report_month_str, lang, t)

def sanitize_filename(text: str) -> str:
    """Sanitizes a value to be safe for use as a filename by converting it to a string first."""
    text = str(text).replace(" ", "_")
    return re.sub(r"[^a-zA-Z0-9_]", "", text)[:50]

@st.cache_data
def generate_reports_zip(customers_df: pd.DataFrame, report_month_str: str, lang: str) -> bytes:
    def t(key): return TRANSLATIONS[lang].get(key, key)
    buf = io.BytesIO()
    report_date = parse_month_string(report_month_str)
    filename_date_str = report_date.strftime('%m%Y') if report_date else "data"
    with zipfile.ZipFile(buf, 'a', zipfile.ZIP_DEFLATED) as zf:
        for _, customer_row in customers_df.iterrows():
            customer_name_sanitized = sanitize_filename(customer_row['customer_name'])
            report_filename = f"{customer_row['customer_id']}_{customer_name_sanitized}_{filename_date_str}.svg"
            report_svg_content = create_report_html(TEMPLATE_FILE_SVG, customer_row, report_month_str, lang, t)
            zf.writestr(report_filename, report_svg_content)
    return buf.getvalue()

# --- Streamlit UI Components ---
def display_customer_data_and_actions(customers_df: pd.DataFrame, report_month_str: str, lang: str, t):
    st.subheader(t("review_data_header"))
    st.caption(t("review_data_caption"))
    event = st.dataframe(customers_df, key="data_selection", on_select="rerun", selection_mode="single-row", hide_index=True, use_container_width=True)
    if event.selection and event.selection["rows"]:
        selected_row_index = event.selection["rows"][0]
        selected_customer = customers_df.iloc[selected_row_index]
        st.subheader(t("preview_header").format(customer_name=selected_customer['customer_name']))
        with st.container(border=True):
            st.html(generate_single_report_preview(selected_customer, report_month_str, lang))
        st.divider()
    if not customers_df.empty:
        report_date = parse_month_string(report_month_str)
        zip_filename_month = format_local_date(report_date, lang, t) if report_date else "Reports"
        zip_filename = t("zip_filename").format(month_name=zip_filename_month)
        if st.button(t("prepare_reports_button"), use_container_width=True, type="primary"):
            zip_data = generate_reports_zip(customers_df, report_month_str, lang)
            st.download_button(label=t("download_reports_button"), data=zip_data, file_name=zip_filename, mime="application/zip", icon="📦", use_container_width=True)

def process_spreadsheet(xls_file: st.runtime.uploaded_file_manager.UploadedFile, lang: str, t):
    try:
        months_row_df, main_df_headers = pd.read_excel(xls_file, nrows=1, header=None), pd.read_excel(xls_file, header=1, nrows=0)
        xls_file.seek(0)
        available_columns = [str(col).strip() for col in main_df_headers.columns]
        month_map = {}
        for i, cell_value in enumerate(months_row_df.iloc[0]):
            if pd.notna(cell_value):
                month_str = None
                if isinstance(cell_value, datetime.datetime): month_str = format_local_date(cell_value, lang, t)
                elif isinstance(cell_value, str) and parse_month_string(cell_value): month_str = cell_value.strip().upper()
                if month_str: month_map[month_str] = i
        if not month_map:
            st.error(t("no_month_header_error"))
            return
        st.subheader(t("select_month_header"))
        selected_month_name = st.selectbox(t("select_month_label"), options=list(month_map.keys()))
        st.subheader(t("map_columns_header"))
        st.caption(t("map_columns_caption"))
        col1, col2 = st.columns(2)
        with col1: customer_id_col_name = st.selectbox(t("customer_id_label"), options=available_columns, index=get_default_index("id", available_columns), disabled=True)
        with col2: customer_name_col_name = st.selectbox(t("customer_name_label"), options=available_columns, index=get_default_index("name", available_columns), disabled=True)
        month_start_idx, month_keys = month_map[selected_month_name], list(month_map.keys())
        current_month_list_idx = month_keys.index(selected_month_name)
        next_month_start_idx = len(available_columns)
        if current_month_list_idx + 1 < len(month_keys): next_month_start_idx = month_map[month_keys[current_month_list_idx + 1]]
        month_specific_columns = available_columns[month_start_idx:next_month_start_idx]
        
        month_name_for_label = selected_month_name.split(' ')[0].title()
        customer_total_col_name = st.selectbox(t("waste_total_label").format(month_name=month_name_for_label), options=month_specific_columns, index=get_default_index("total", month_specific_columns), disabled=True)
        
        id_col_idx, name_col_idx, total_col_idx = available_columns.index(customer_id_col_name), available_columns.index(customer_name_col_name), available_columns.index(customer_total_col_name)
        raw_headers = main_df_headers.columns
        customers_df = pd.read_excel(xls_file, header=1, usecols=[id_col_idx, name_col_idx, total_col_idx], dtype=str).rename(columns={raw_headers[id_col_idx]: 'customer_id', raw_headers[name_col_idx]: 'customer_name', raw_headers[total_col_idx]: 'customer_total'})
        xls_file.seek(0)
        customers_df = customers_df.dropna(how='all')
        customers_df['customer_id'] = pd.to_numeric(customers_df['customer_id'], errors='coerce')
        customers_df['customer_total'] = pd.to_numeric(customers_df['customer_total'], errors='coerce')
        customers_df = customers_df.dropna(subset=['customer_id', 'customer_total'])
        customers_df['customer_id'] = customers_df['customer_id'].astype(int)
        st.divider()
        if not customers_df.empty: display_customer_data_and_actions(customers_df, selected_month_name, lang, t)
        else: st.warning(t("no_data_warning"))
    except Exception as e: st.error(t("process_error").format(error=e))

# --- Main Application Logic ---
def main():
    if 'language' not in st.session_state:
        st.session_state.language = 'en'

    lang = st.session_state.language
    def t(key): return TRANSLATIONS[lang].get(key, key)

    st.set_page_config(page_title=t("page_title"), layout="wide", initial_sidebar_state="collapsed")

    _, col_en, col_pt = st.columns([10, 1, 1])
    if col_en.button('🇬🇧 English'):
        st.session_state.language = 'en'
        st.rerun()
    if col_pt.button('🇧🇷 Português'):
        st.session_state.language = 'pt'
        st.rerun()

    st.title(t("main_title"))
    st.markdown(t("main_subtitle"))

    with st.expander(t("view_format_expander")):
        st.markdown(t("format_intro"))
        st.markdown(t("format_header1"))
        st.markdown(t("format_header2"))
        st.markdown(t("format_columns"))
        st.markdown(t("format_example_title"))
        st.code(t("format_example_content"), language='markdown')

    uploaded_file = upload_xls_file(t)
    if uploaded_file:
        st.divider()
        process_spreadsheet(uploaded_file, lang, t)

if __name__ == '__main__':
    main()
