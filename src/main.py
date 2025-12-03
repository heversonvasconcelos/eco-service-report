# main.py
import base64
import datetime
import io
import os
import re
import unicodedata
import zipfile
from typing import Optional

import pandas as pd
import streamlit as st
from jinja2 import Template

# --- I18N Translations ---
TRANSLATIONS = {
    "en": {
        "page_title": "Eco Service Reports", "main_title": "♻️ Eco Service Report Generator",
        "main_subtitle": "Upload your customer data (XLS/XLSX) to generate monthly environmental impact reports.",
        "view_format_expander": "View Expected Spreadsheet Format",
        "format_intro": "The application expects an Excel file with a single header row containing dates.",
        "format_header1": "- **Header Row**: Contains column titles. Columns representing weeks should use the date format `DD/MM/YYYY` (e.g., `07/11/2025`).",
        "format_header2": "",
        "format_columns": "The columns `ID` and `NAME` should appear before the date columns.",
        "format_example_title": "**Example Layout:**",
        "format_example_content": """
| ID   | NAME          | 07/11/2025 | 14/11/2025 | ... | TOTAL      | 05/12/2025 ...
|------|---------------|------------|------------|-----|------------|----------------
| 1    | Client A      | 30.0       | 40.0       | ... | 150.5      | 40.0       ...
| 2    | Client B      | 50.0       | 50.0       | ... | 200.0      | 55.0       ...
        """,
        "upload_label": "Upload Spreadsheet", "file_uploaded_success": "File uploaded: {file_name}",
        "process_error": "An error occurred while processing the spreadsheet: {error}",
        "no_month_header_error": "Could not find any valid date columns (DD/MM/YYYY) in the header.",
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
        "months": ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
                   "November", "December"],
        "month_of_year_format": "{month_name} {year}", "month_year_format": "%B-%Y", "last_month_na": "N/A",
        "report_id_na": "N/A",
        # Report Template Translations
        "report_title": "Environmental Impact Report", "report_id_label": "Report ID:",
        "metric_waste": "Total Waste Diverted", "metric_fertilizer": "Organic Fertilizer",
        "metric_co2": "CO₂ Avoided", "metric_driving": "Driving Distance",
        "metric_trees": "Trees Equivalent", "metric_water": "Water Saved",
        "report_footer": "Thank you for making a positive impact on the environment! | Generated on {date}"
    },
    "pt": {
        "page_title": "Relatórios Eco Service", "main_title": "♻️ Gerador de Relatórios Eco Service",
        "main_subtitle": "Faça o upload dos dados dos seus clientes (XLS/XLSX) para gerar relatórios mensais de impacto ambiental.",
        "view_format_expander": "Ver Formato da Planilha Esperado",
        "format_intro": "A aplicação espera um arquivo Excel com uma única linha de cabeçalho contendo datas.",
        "format_header1": "- **Linha de Cabeçalho**: Contém os títulos das colunas. Colunas que representam semanas devem usar o formato de data `DD/MM/AAAA` (ex: `07/11/2025`).",
        "format_header2": "",
        "format_columns": "As colunas `ID` e `NOME` devem aparecer antes das colunas de data.",
        "format_example_title": "**Exemplo de Layout:**",
        "format_example_content": """
| ID   | NOME          | 07/11/2025 | 14/11/2025 | ... | TOTAL      | 05/12/2025 ...
|------|---------------|------------|------------|-----|------------|----------------
| 1    | Cliente A     | 30.0       | 40.0       | ... | 150.5      | 40.0       ...
| 2    | Cliente B     | 50.0       | 50.0       | ... | 200.0      | 55.0       ...
        """,
        "upload_label": "Fazer Upload da Planilha", "file_uploaded_success": "Arquivo enviado: {file_name}",
        "process_error": "Ocorreu um erro ao processar a planilha: {error}",
        "no_month_header_error": "Não foi possível encontrar colunas de data (DD/MM/AAAA) no cabeçalho.",
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
        "months": ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro",
                   "Novembro", "Dezembro"],
        "month_of_year_format": "{month_name} {year}", "month_year_format": "%B-%Y", "last_month_na": "N/D",
        "report_id_na": "N/D",
        # Report Template Translations
        "report_title": "Relatório de Impacto Ambiental", "report_id_label": "ID do Relatório:",
        "metric_waste": "Total de Resíduos Desviados", "metric_fertilizer": "Fertilizante Orgânico",
        "metric_co2": "CO₂ Evitado", "metric_driving": "Distância de Carro",
        "metric_trees": "Árvores Equivalentes", "metric_water": "Água Economizada",
        "report_footer": "Obrigado por causar um impacto positivo no meio ambiente! | Gerado em {date}"
    }
}

# --- Constants ---
DEFAULT_XLS_COLUMNS = {"id": "ID", "name": "NOME", "total": "TOTAL"}
TEMPLATE_FILE_HTML = 'templates/service-report-preview-template.html'
TEMPLATE_FILE_SVG = 'templates/service-report-template.svg'
FERTILIZER_FACTOR, CO2_AVOIDED_FACTOR, DRIVING_DISTANCE_CO2_CONVERSION_FACTOR, TREES_EQUIVALENT_CO2_ABSORPTION_FACTOR, WATER_LITERS_FACTOR_BASE, WATER_LITERS_MULTIPLIER = 0.38, 0.77, 0.096, 0.35, 0.214, 12
MONTH_PT_TO_NUM = {name.lower(): i + 1 for i, name in enumerate(TRANSLATIONS["pt"]["months"])}
MONTH_EN_TO_NUM = {name.lower(): i + 1 for i, name in enumerate(TRANSLATIONS["en"]["months"])}
MONTH_NAME_TO_NUM = {**MONTH_PT_TO_NUM, **MONTH_EN_TO_NUM}


# --- Date Handling ---
def format_local_date(date_obj: datetime.datetime, lang: str, t) -> str:
    t_months = TRANSLATIONS[lang]["months"]
    month_name = t_months[date_obj.month - 1]
    return t("month_of_year_format").format(month_name=month_name, year=date_obj.year)


def parse_month_string(month_str: any) -> Optional[datetime.datetime]:
    if isinstance(month_str, datetime.datetime):
        # If it's already a datetime object, extract year and month, set day to 1.
        return datetime.datetime(month_str.year, month_str.month, 1)
    
    month_str = str(month_str).strip()
    
    # Try parsing DD/MM/YYYY
    try:
        # Use a regex to extract day, month, year from DD/MM/YYYY, then reformat to YYYY-MM-DD for datetime parsing
        match = re.match(r'(\d{2})/(\d{2})/(\d{4})', month_str)
        if match:
            day, month, year = map(int, match.groups())
            return datetime.datetime(year, month, day)
    except (ValueError, TypeError):
        pass

    # Try parsing DD-MM-YYYY
    try:
        return datetime.datetime.strptime(month_str, '%d-%m-%Y')
    except (ValueError, TypeError):
        pass

    # Try parsing YYYY-MM-DD (e.g., from sanitized strings or other sources)
    try:
        return datetime.datetime.strptime(month_str, '%Y-%m-%d')
    except (ValueError, TypeError):
        pass

    # Try parsing YYYY-MM-DD HH:MM:SS (original format)
    try:
        return datetime.datetime.strptime(month_str, '%Y-%m-%d %H:%M:%S')
    except (ValueError, TypeError):
        pass

    # After trying all specific date formats, if not parsed, try month name + year
    month_str_lower = month_str.lower()
    for month_name, month_num in MONTH_NAME_TO_NUM.items():
        if month_str_lower.startswith(month_name):
            # The remaining part should contain the year, possibly with noise
            remaining_str = month_str_lower[len(month_name):].strip()
            
            # Extract only digits for the year from the remaining string
            year_str = re.sub(r'\D', '', remaining_str)
            
            if year_str.isdigit() and len(year_str) == 4: # Assume 4-digit year
                try:
                    return datetime.datetime(int(year_str), month_num, 1)
                except (ValueError, TypeError):
                    pass
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
@st.cache_data
def get_ods_images() -> dict[str, str]:
    images = {}
    image_files = {
        "ods_2": "ods-2.svg", "ods_3": "ods-3.svg", "ods_6": "ods-6.svg",
        "ods_11": "ods-11.svg", "ods_12": "ods-12.svg", "ods_13": "ods-13.svg", "ods_15": "ods-15.svg"
    }
    dir_path = os.path.dirname(os.path.realpath(__file__))
    images_dir = os.path.join(dir_path, 'images')
    for key, filename in image_files.items():
        file_path = os.path.join(images_dir, filename)
        try:
            with open(file_path, "rb") as f:
                encoded = base64.b64encode(f.read()).decode("utf-8")
                images[key] = f"data:image/svg+xml;base64,{encoded}"
        except FileNotFoundError:
            images[key] = ""
    return images

def create_report_html(template_path: str, customer_data: pd.Series, report_month_str: str, lang: str, t) -> str:
    dir_path = os.path.dirname(os.path.realpath(__file__))
    template_path = os.path.join(dir_path, template_path)
    with open(template_path, 'r', encoding='utf-8') as f:
        jinja_template = Template(f.read())

    report_date = parse_month_string(report_month_str)
    if report_date:
        report_id_date_str = report_date.strftime('%Y-%m')
    else:
        report_id_date_str = t("report_id_na")

    current_date = format_local_date(datetime.datetime.now(), lang, t)
    report_id = f"{customer_data['customer_id']}-{report_id_date_str}"
    customer_total_float = customer_data['customer_total']
    fertilizer_kg, co2_avoided = customer_total_float * FERTILIZER_FACTOR, customer_total_float * CO2_AVOIDED_FACTOR
    driving_distance, trees_equivalent, water_liters = co2_avoided / DRIVING_DISTANCE_CO2_CONVERSION_FACTOR, co2_avoided / TREES_EQUIVALENT_CO2_ABSORPTION_FACTOR, customer_total_float * WATER_LITERS_FACTOR_BASE * WATER_LITERS_MULTIPLIER

    template_vars = {
        'report_title': t('report_title'),
        'report_date': report_month_str,
        'report_id_label': t('report_id_label'),
        'report_id': report_id,
        'current_date': current_date,
        'customer_name': customer_data['customer_name'],
        'metric_waste': t('metric_waste'),
        'customer_waste_kg': f"{customer_total_float:.2f}",
        'metric_fertilizer': t('metric_fertilizer'),
        'fertilizer_kg': f"{fertilizer_kg:.2f}",
        'metric_co2': t('metric_co2'),
        'co2_avoided': f"{co2_avoided:.2f}",
        'metric_driving': t('metric_driving'),
        'driving_distance': f"{driving_distance:.2f}",
        'metric_trees': t('metric_trees'),
        'trees_equivalent': f"{trees_equivalent:.0f}",
        'metric_water': t('metric_water'),
        'water_liters': f"{water_liters:.2f}",
        'report_footer': t('report_footer').format(date=current_date)
    }
    template_vars.update(get_ods_images())
    return jinja_template.render(template_vars)


@st.cache_data
def generate_single_report_preview(customer_data: pd.Series, report_month_str: str, lang: str) -> str:
    def t(key): return TRANSLATIONS[lang].get(key, key)

    return create_report_html(TEMPLATE_FILE_HTML, customer_data, report_month_str, lang, t)


def sanitize_filename(text: str) -> str:
    """Sanitizes a value to be safe for use as a filename by converting it to a string first."""
    text = str(text).strip()
    # Normalize unicode characters to their closest ASCII equivalents
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    # Replace all non-alphanumeric (except underscore) characters with a single underscore
    text = re.sub(r'[^a-zA-Z0-9_]+', '_', text)
    # Remove leading/trailing underscores
    text = text.strip('_')
    return text[:50]


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
    event = st.dataframe(customers_df, key="data_selection", on_select="rerun", selection_mode="single-row",
                         hide_index=True, use_container_width=True)
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
        zip_filename = sanitize_filename(t("zip_filename").format(month_name=zip_filename_month))
        if st.button(t("prepare_reports_button"), use_container_width=True, type="primary"):
            zip_data = generate_reports_zip(customers_df, report_month_str, lang)
            st.download_button(label=t("download_reports_button"), data=zip_data, file_name=zip_filename,
                               mime="application/zip", icon="📦", use_container_width=True)


@st.cache_data
def load_excel_data(file_content: bytes):
    excel_file = io.BytesIO(file_content)
    # Read using the first row (header=0) as header
    full_df = pd.read_excel(excel_file, header=0, dtype=str)
    return full_df.columns.tolist(), full_df


def process_spreadsheet(xls_file: st.runtime.uploaded_file_manager.UploadedFile, lang: str, t):
    try:
        file_content = xls_file.getvalue()
        raw_columns, full_df = load_excel_data(file_content)
        
        available_columns = [str(col).strip() for col in raw_columns]
        month_map = {}
        
        # Identify months based on TOTAL columns and their preceding date column
        for i, col_name in enumerate(available_columns):
            if "TOTAL" in col_name.upper():
                if i > 0:
                    prev_col_name = available_columns[i-1]
                    parsed_date = parse_month_string(prev_col_name)
                    if parsed_date:
                        month_str = format_local_date(parsed_date, lang, t)
                        month_map[month_str] = {
                            "total_idx": i,
                            "total_col_name": col_name,
                            "date_source": prev_col_name
                        }
                        
        if not month_map:
            st.error(t("no_month_header_error"))
            return

        st.subheader(t("select_month_header"))
        selected_month_name = st.selectbox(t("select_month_label"), options=list(month_map.keys()))
        
        st.subheader(t("map_columns_header"))
        st.caption(t("map_columns_caption"))
        col1, col2 = st.columns(2)
        
        with col1:
            customer_id_col_name = st.selectbox(t("customer_id_label"), options=available_columns,
                                                index=get_default_index("id", available_columns))
        with col2:
            customer_name_col_name = st.selectbox(t("customer_name_label"), options=available_columns,
                                                  index=get_default_index("name", available_columns))
        
        selected_month_data = month_map[selected_month_name]
        total_col_idx = selected_month_data["total_idx"]
        customer_total_col_name = selected_month_data["total_col_name"]
        
        month_name_for_label = selected_month_name.split(' ')[0].title()
        st.info(f"{t('waste_total_label').format(month_name=month_name_for_label)}: **{customer_total_col_name}**")

        id_col_idx = available_columns.index(customer_id_col_name)
        name_col_idx = available_columns.index(customer_name_col_name)
        
        customers_df = full_df.iloc[:, [id_col_idx, name_col_idx, total_col_idx]].copy()
        customers_df.columns = ['customer_id', 'customer_name', 'customer_total']
        
        customers_df = customers_df.dropna(how='all')
        customers_df['customer_id'] = pd.to_numeric(customers_df['customer_id'], errors='coerce')
        customers_df['customer_total'] = pd.to_numeric(customers_df['customer_total'], errors='coerce')
        customers_df = customers_df.dropna(subset=['customer_id', 'customer_total'])

        # Remove rows where customer_name is NaN, 'nan', 'None', or empty
        customers_df = customers_df[~customers_df['customer_name'].astype(str).str.strip().str.lower().isin(['nan', 'none', ''])]
        
        customers_df['customer_id'] = customers_df['customer_id'].astype(int)
        
        st.divider()
        if not customers_df.empty:
            display_customer_data_and_actions(customers_df, selected_month_name, lang, t)
        else:
            st.warning(t("no_data_warning"))
    except Exception as e:
        st.error(t("process_error").format(error=e))


# --- Main Application Logic ---
def main():
    if 'language' not in st.session_state:
        st.session_state.language = 'pt'

    lang = st.session_state.language

    def t(key):
        return TRANSLATIONS[lang].get(key, key)

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
