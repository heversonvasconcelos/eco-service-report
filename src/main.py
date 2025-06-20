# main.py
import io
import zipfile
import datetime  # Combined datetime import
import locale
from typing import Tuple, Optional, Any, Dict

import pandas as pd
import streamlit as st
from jinja2 import Template

# --- Constants ---
# Consider moving these to a separate config file or Streamlit secrets for more complex apps
DEFAULT_XLS_COLUMNS = {
    "id": "ID",
    "name": "NOME COMPLETO",
    "total": "TOTAL16"  # Example: Be specific if the column name is fixed
}
TEMPLATE_FILE_HTML = 'templates/service-report-preview-template.html'
TEMPLATE_FILE_SVG = 'templates/service-report-template.svg'  # If you plan to use SVG again

# --- Report Calculation Constants ---
FERTILIZER_FACTOR = 0.38
CO2_AVOIDED_FACTOR = 0.77
DRIVING_DISTANCE_CO2_CONVERSION_FACTOR = 0.096  # kg CO2 per km (example value)
TREES_EQUIVALENT_CO2_ABSORPTION_FACTOR = 0.35  # kg CO2 per tree per month (example value)
WATER_LITERS_FACTOR_BASE = 0.214
WATER_LITERS_MULTIPLIER = 12


# --- Helper Functions ---

def get_translated_month_year(month_year_format: str = '%B DE %Y') -> str:
    """
    Gets the current month and year, translated to Portuguese (Brazil).
    """
    try:
        # Ensure locale is set, but be mindful if this runs in a shared environment
        # It's often better to handle localization at the presentation layer if possible
        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
        current_month_year = datetime.datetime.now().strftime(month_year_format)
        return current_month_year.upper()
    except locale.Error:
        st.warning("Locale 'pt_BR.UTF-8' not available. Using default system locale for month name.")
        # Fallback to a non-localized or English format if pt_BR is not available
        return datetime.datetime.now().strftime(month_year_format).upper()


# --- File Handling and Data Processing ---

def upload_xls_file() -> Optional[st.runtime.uploaded_file_manager.UploadedFile]:
    """
    Displays a file uploader for XLS/XLSX files and returns the uploaded file.
    """
    uploaded_file = st.file_uploader("Upload XLS/XLSX file", type=["xls", "xlsx"])
    if uploaded_file is not None:
        st.success(f"File uploaded: {uploaded_file.name}")
        return uploaded_file
    return None


def select_columns_from_xls(xls_file: st.runtime.uploaded_file_manager.UploadedFile) -> tuple[tuple[Any, int], tuple[
    Any, int], tuple[Any, int]] | None:
    """
    Allows the user to select columns for ID, Name, and Total from the uploaded XLS file.
    Returns the selected column names.
    """
    try:
        # Read only the header and first row to get column names
        # skiprows=[0] might be problematic if the actual headers are on the first row.
        # If headers are on row 0, remove skiprows=[0]. If data starts after a blank row, adjust.
        # Assuming headers are on the first row (index 0) and data starts on the second (index 1)
        header_df = pd.read_excel(xls_file, nrows=1, skiprows=[0])  # Reads the first row for headers
        xls_file.seek(0)  # Reset file pointer if reading again

        header_df = header_df.rename(columns=lambda x: x.strip())
        available_columns = [col.strip() for col in header_df.columns]

        st.subheader("Map Your Spreadsheet Columns")
        st.caption("Select the columns from your uploaded file that correspond to the required fields.")

        # Use a more robust way to find default indices
        def get_default_index(col_name_key: str) -> int:
            default_col_name = DEFAULT_XLS_COLUMNS.get(col_name_key)
            if default_col_name and default_col_name in available_columns:
                return available_columns.index(default_col_name)
            return 0  # Default to the first column if not found

        customer_id_col = st.selectbox(
            "Customer ID Column",
            options=available_columns,
            index=get_default_index("id"),
            help="Select the column containing unique customer identifiers."
        )
        customer_id_col = customer_id_col, available_columns.index(customer_id_col)

        customer_name_col = st.selectbox(
            "Customer Name Column",
            options=available_columns,
            index=get_default_index("name"),
            help="Select the column containing the full name of the customer."
        )
        customer_name_col = customer_name_col, available_columns.index(customer_name_col)

        customer_total_col = st.selectbox(
            "Waste Total Column",
            options=available_columns,
            index=get_default_index("total"),
            help="Select the column containing the total waste amount."
        )
        customer_total_col = customer_total_col, available_columns.index(customer_total_col)

        return customer_id_col, customer_name_col, customer_total_col
    except Exception as e:
        st.error(f"Error reading columns from XLS: {e}")
        return None


def load_and_prepare_data(xls_file: st.runtime.uploaded_file_manager.UploadedFile,
                          id_col: Tuple[Any, int], name_col: Tuple[Any, int], total_col: Tuple[Any, int]) -> Optional[
    pd.DataFrame]:
    """
    Reads data from the XLS file using selected columns, cleans, and renames it.
    """
    try:
        # Specify columns to use and their data types
        # skiprows=1 assumes the first row is the header row. Adjust if different.
        customers_data = pd.read_excel(
            xls_file,
            skiprows=1,  # Skip the header row as we've already processed it
            usecols=[id_col[1], name_col[1], total_col[1]],
            dtype={id_col: str, name_col: str, total_col: str}  # Keep as str initially for safety
        )
        xls_file.seek(0)  # Reset file pointer

        if customers_data.empty:
            st.warning("The selected columns in the XLS file are empty or do not contain valid data.")
            return None

        # Clean and rename columns
        customers_data = customers_data.rename(columns=lambda x: x.strip())
        customers_data = customers_data.rename(
            columns={
                id_col[0]: 'customer_id',
                name_col[0]: 'customer_name',
                total_col[0]: 'customer_total'
            }
        )
        customers_data = customers_data.dropna()

        # Attempt to convert to numeric, handle errors
        customers_data['customer_id'] = pd.to_numeric(customers_data['customer_id'], downcast='integer',
                                                      errors='coerce')
        customers_data['customer_total'] = pd.to_numeric(customers_data['customer_total'], errors='coerce')
        customers_data = customers_data.dropna(subset=['customer_total'])  # Remove rows where conversion failed

        return customers_data

    except Exception as e:
        st.error(f"Error processing XLS data: {e}")
        return None


# --- Report Generation ---

def create_report_html(template_path: str, customer_data: pd.Series) -> str:
    """
    Renders an HTML report for a single customer using a Jinja2 template.
    """
    with open(template_path, 'r', encoding='utf-8') as f:
        template_content = f.read()
        jinja_template = Template(template_content)

    current_month_year_display = get_translated_month_year('%B DE %Y')
    current_month_year_id = datetime.datetime.now().strftime('%Y-%m')  # For report ID consistency

    report_id = f"{customer_data['customer_id']}-{current_month_year_id}"
    customer_name = customer_data['customer_name']
    customer_total_float = customer_data['customer_total']

    # Calculations with named constants
    fertilizer_kg = customer_total_float * FERTILIZER_FACTOR
    co2_avoided = customer_total_float * CO2_AVOIDED_FACTOR
    driving_distance = (
            co2_avoided / DRIVING_DISTANCE_CO2_CONVERSION_FACTOR) if DRIVING_DISTANCE_CO2_CONVERSION_FACTOR else 0
    trees_equivalent = (
            co2_avoided / TREES_EQUIVALENT_CO2_ABSORPTION_FACTOR) if TREES_EQUIVALENT_CO2_ABSORPTION_FACTOR else 0
    water_liters = customer_total_float * WATER_LITERS_FACTOR_BASE * WATER_LITERS_MULTIPLIER

    template_vars = {
        'month_year': current_month_year_display,
        'report_id': report_id,
        'customer_name': customer_name,
        'customer_waste_kg': f"{customer_total_float:.2f}",
        'fertilizer_kg': f"{fertilizer_kg:.2f}",
        'co2_avoided': f"{co2_avoided:.2f}",
        'driving_distante': f"{driving_distance:.2f}",  # Typo in original template variable? 'distante' vs 'distance'
        'trees_equivalent': f"{trees_equivalent:.0f}",
        'water_liters': f"{water_liters:.2f}"
    }
    return jinja_template.render(template_vars)


@st.cache_data  # Cache the generation of the preview
def generate_single_report_preview(customer_data: pd.Series) -> str:
    """
    Generates HTML for a single customer report preview.
    customer_data is a Pandas Series representing one row.
    """
    # No need to re-assign customer_data.costumer_id etc.
    # Access values directly: customer_data['customer_id']
    return create_report_html(TEMPLATE_FILE_HTML, customer_data)


@st.cache_data  # Cache the generation of the zip file
def generate_reports_zip(customers_df: pd.DataFrame) -> bytes:
    """
    Generates all customer reports and packages them into a zip file.
    """
    buf = io.BytesIO()
    current_month_year_filename = get_translated_month_year('%B-%Y')  # For filenames

    with zipfile.ZipFile(buf, 'a', zipfile.ZIP_DEFLATED) as zf:
        for _, customer_row in customers_df.iterrows():
            report_filename = f"{customer_row['customer_id']}-{current_month_year_filename}.svg"
            report_html = create_report_html(TEMPLATE_FILE_SVG, customer_row)
            zf.writestr(report_filename, report_html)
            # st.toast(f"Report generated for {customer_row['customer_name']} at {report_filename}") # Optional: for progress
    return buf.getvalue()


# --- Streamlit UI Components ---

def display_customer_data_and_actions(customers_df: pd.DataFrame):
    """
    Displays customer data in a Streamlit dataframe and handles selection/download.
    """
    st.subheader("Confirm Customer Data and Generate Reports")
    st.caption("Review the processed data. Select a row to preview its report.")

    event = st.dataframe(
        customers_df,
        key="data_selection",  # Changed key to avoid potential conflicts
        on_select="rerun",
        selection_mode="single-row",  # Corrected: was ["single-row"]
        hide_index=True,
        use_container_width=True
    )

    if event.selection and event.selection["rows"]:
        selected_row_index = event.selection["rows"][0]
        selected_customer = customers_df.iloc[selected_row_index]

        st.subheader(f"Preview: Report for {selected_customer['customer_name']}")
        with st.container(border=True):
            preview_html = generate_single_report_preview(selected_customer)
            st.html(preview_html)
        st.divider()

    if not customers_df.empty:
        if st.button("Prepare All Reports (.zip)", use_container_width=True):
            zip_filename = get_translated_month_year('%B-%Y') + '-Reports.zip'
            zip_data = generate_reports_zip(customers_df)

            st.download_button(
                label="Download",
                data=zip_data,
                file_name=zip_filename,
                mime="application/zip",
                icon=":material/download:",
                use_container_width=True
            )
            st.toast(f"All reports ready for download: {zip_filename}")  # Optional feedback


# --- Main Application Logic ---
def main():
    """
    Main function to run the Streamlit application.
    """
    st.set_page_config(page_title="Eco Service Reports", layout="wide")
    st.title("♻️ Eco Service Report Generator")
    st.markdown("Upload your customer data (XLS/XLSX) to generate monthly environmental impact reports.")

    uploaded_file = upload_xls_file()

    if uploaded_file:
        st.divider()
        selected_columns = select_columns_from_xls(uploaded_file)

        if selected_columns:
            id_col, name_col, total_col = selected_columns
            customers_df = load_and_prepare_data(uploaded_file, id_col, name_col, total_col)

            if customers_df is not None and not customers_df.empty:
                st.divider()
                display_customer_data_and_actions(customers_df)
            elif customers_df is not None and customers_df.empty:
                st.info("No valid customer data found after processing. Please check your file and column selections.")


if __name__ == '__main__':
    main()
