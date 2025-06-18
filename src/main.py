# main.py
from datetime import datetime

import pandas as pd
import streamlit as st
from jinja2 import Template


def upload_xls_file():
    uploaded_file = st.file_uploader("Upload XLS/XLSX file", type=["xls", "xlsx"])
    if uploaded_file is not None:
        st.success(f"File uploaded: {uploaded_file.name}")
        return uploaded_file
    return None


def filter_columns_from_xls(xls_file):
    data = pd.read_excel(xls_file, skiprows=[0], nrows=2)

    costumer_data = data.rename(columns=lambda x: x.strip())

    costumer_id = st.selectbox(
        "ID",
        index=(costumer_data.columns.get_loc("ID")),
        options=costumer_data.columns
    )
    costumer_id = costumer_data.columns.get_loc(costumer_id)

    costumer_name = st.selectbox(
        "NOME COMPLETO",
        index=(costumer_data.columns.get_loc("NOME COMPLETO")),
        options=costumer_data.columns
    )
    costumer_name = costumer_data.columns.get_loc(costumer_name)

    costumer_total = st.selectbox(
        "TOTAL",
        index=(costumer_data.columns.get_loc("TOTAL16")),
        options=costumer_data.columns
    )
    costumer_total = costumer_data.columns.get_loc(costumer_total)

    return costumer_id, costumer_name, costumer_total


def read_xls_file(xls_file):
    costumers_fields = filter_columns_from_xls(xls_file)
    # Read data from the XLS file
    costumers_data = pd.read_excel(xls_file,
                                   skiprows=[0],
                                   usecols=costumers_fields,
                                   dtype={0: str, 1: str, 2: str})  # Read the first sheet

    # Check if the data is empty
    if costumers_data.empty:
        print("Error: The XLS file is empty or does not contain valid data.")
        return

    costumers_data = costumers_data.dropna()
    costumers_data = costumers_data.rename(columns=lambda x: x.strip())
    costumers_data = costumers_data.rename(columns={costumers_fields[0]: 'costumer_id',
                                                    costumers_fields[1]: 'costumer_name',
                                                    costumers_fields[2]: 'costumer_total'})

    present_customers_data_to_confirm(costumers_data)


def present_customers_data_to_confirm(costumers_data):
    event = st.dataframe(
        costumers_data,
        key="data",
        on_select="rerun",
        selection_mode=["single-row"],
        hide_index=True,
    )

    if event.selection and event.selection.rows:
        selected_costumer = costumers_data.iloc[event.selection.rows[0]]
        # selected_month = month_data.iloc[event.selection.rows[0]]
        preview_costumer_report(selected_costumer)


def get_costumer_report_file_name(costumer_id):
    month_year = get_translated_month_year('%B-%Y')
    costumer_report_file_name = costumer_id + '-' + month_year + '.html'
    return costumer_report_file_name


def preview_costumer_report(costumer_data):
    costumer_data.costumer_id = costumer_data[0]
    costumer_data.costumer_name = costumer_data[1]
    costumer_data.costumer_total = costumer_data[2]

    # message = container.header(f":green[ {costumer_data.costumer_name} ]")
    # container.markdown(":green[Você está nos ajudando a transformar o mundo mais sustentável!] :sunglasses:")
    # container.divider()
    # container.subheader(f"ESTE FOI SEU IMPACTO :blue[POSITIVO] EM {month_year}")
    # container.markdown(
    #     f"Deixou de enviar para o aterro sanitário :green[{float(costumer_data.costumer_total):.2f}] kg :clap:")
    # container.markdown(
    #     f"Que se transformou em :green[{(float(costumer_data.costumer_total) * 0.38):.2f}] kg de adubo orgânico :recycle:")
    # container.markdown(
    #     f"Também foi evitado o lançamento de :green[{(float(costumer_data.costumer_total) * 0.77):.2f}] kg de CO² na atmosfera :partly_sunny:")
    # container.markdown(
    #     f"Equivale á :green[{((float(costumer_data.costumer_total) * 0.77) / 0.096):.2f}] km rodados de carro :car:")
    # container.markdown(
    #     f"Também equivale ao sequestro de CO² de :green[{((float(costumer_data.costumer_total) * 0.77) / 0.35):.0f}] árvores :deciduous_tree:")
    # container.markdown(
    #     f"Com a sua iniciativa evitamos a contaminação de :green[{((float(costumer_data.costumer_total) * 0.214) * 12):.2f}] litros de água :droplet: :national_park:")

    # if st.button("Prepare download"):
    #     st.download_button(
    #         label="Download text",
    #         data="TESTE",
    #         file_name="message.txt",
    #         on_click="ignore",
    #         type="primary",
    #         icon=":material/download:",
    #     )

    costumer_report_file_name = get_costumer_report_file_name(costumer_data.costumer_id)

    container = st.container(border=True)
    html = generate_monthly_report('templates/service-report-template.v0.html', costumer_data,
                                   costumer_report_file_name)
    container.html(html)

    st.divider()

    html = generate_monthly_report('templates/service-report-template.html', costumer_data, costumer_report_file_name)
    if st.download_button(
            label="Download reports",
            data=html,
            file_name=costumer_report_file_name,
            mime="text/html",
            icon=":material/download:",
    ):
        st.toast(f"Report generated for {costumer_data.costumer_name} at {costumer_report_file_name}")


def generate_monthly_report(html_template_file, costumer_data, report_file_name):
    # Load the Jinja2 template
    with open(html_template_file, 'r') as template_file:
        template_content = template_file.read()
        jinja_template = Template(template_content)

    month_year = get_translated_month_year()
    report_id = f"{costumer_data.costumer_id + '-' + datetime.datetime.now().strftime('%Y-%m')}"
    fertilizer_kg = f"{(float(costumer_data.costumer_total) * 0.38):.2f}"
    co2_avoided = f"{(float(costumer_data.costumer_total) * 0.77):.2f}"
    driving_distante = f"{((float(costumer_data.costumer_total) * 0.77) / 0.096):.2f}"
    trees_equivalent = f"{((float(costumer_data.costumer_total) * 0.77) / 0.35):.0f}"
    water_liters = f"{((float(costumer_data.costumer_total) * 0.214) * 12):.2f}"

    # Render the template
    htmlstr = jinja_template.render(title=report_file_name,
                                    month_year=month_year,
                                    report_id=report_id,
                                    costumer_name=costumer_data.costumer_name,
                                    costumer_waste_kg=costumer_data.costumer_total,
                                    fertilizer_kg=fertilizer_kg,
                                    co2_avoided=co2_avoided,
                                    driving_distante=driving_distante,
                                    trees_equivalent=trees_equivalent,
                                    water_liters=water_liters)

    return htmlstr


def read_html_file(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        html_content = f.read()
    return html_content


def generate_service_report():
    st.header(f":green[ NOSSOS NÚMEROS EM ]")


import datetime
import locale


def get_translated_month_year(month_year_format: str = '%B DE %Y'):
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')  # Set to Portuguese (Brazil)
        current_month_year = datetime.datetime.now().strftime(month_year_format)
        return current_month_year.upper()
    except locale.Error:
        return 'Language not supported'


if __name__ == '__main__':
    file = upload_xls_file()
    if file is not None:
        # Process the uploaded XLS file
        read_xls_file(file)
