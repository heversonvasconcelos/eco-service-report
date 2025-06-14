# main.py

import pandas as pd
import streamlit as st


def upload_xls_file():
    uploaded_file = st.file_uploader("Upload XLS/XLSX file", type=["xls", "xlsx"])
    if uploaded_file is not None:
        st.success(f"File uploaded: {uploaded_file.name}")
        return uploaded_file
    return None


def filter_columns_from_xls(xls_file):
    data = pd.read_excel(xls_file, skiprows=[0])

    data = data.dropna()
    data = data.drop([2, len(data) - 1], axis=0, errors='ignore')  # Drop the third row and the last row if they exist
    data = data.rename(columns=lambda x: x.strip())

    costumer_id = st.selectbox(
        "ID",
        index=(data.columns.get_loc("ID")),
        options=data.columns
    )
    costumer_id = data.columns.get_loc(costumer_id)

    costumer_name = st.selectbox(
        "NOME COMPLETO",
        index=(data.columns.get_loc("NOME COMPLETO")),
        options=data.columns
    )
    costumer_name = data.columns.get_loc(costumer_name)

    costumer_total = st.selectbox(
        "TOTAL",
        index=(data.columns.get_loc("TOTAL16")),
        options=data.columns
    )
    costumer_total = data.columns.get_loc(costumer_total)

    return costumer_id, costumer_name, costumer_total


def read_xls_file(xls_file):
    cols = filter_columns_from_xls(xls_file)
    # Read data from the XLS file
    data = pd.read_excel(xls_file,
                         skiprows=[0],
                         usecols=cols,
                         dtype={0: str, 1: str, 2: str})  # Read the first sheet

    # Check if the data is empty
    if data.empty:
        print("Error: The XLS file is empty or does not contain valid data.")
        return

    data = data.dropna()
    data = data.rename(columns=lambda x: x.strip())
    data = data.rename(columns={cols[0]: 'costumer_id', cols[1]: 'costumer_name', cols[2]: 'costumer_total'})

    present_customers_data_to_confirm(data, cols)


def present_customers_data_to_confirm(data, cols):
    event = st.dataframe(
        data,
        key="data",
        on_select="rerun",
        selection_mode=["single-row"],
        hide_index=True,
    )

    if event.selection and event.selection.rows:
        generate_costumer_report(data.iloc[event.selection.rows[0]])


def generate_costumer_report(costumer_data):
    costumer_data.costumer_id = costumer_data[0]
    costumer_data.costumer_name = costumer_data[1]
    costumer_data.costumer_total = costumer_data[2]

    st.header(f":green[ {costumer_data.costumer_name} ]")
    st.markdown(":green[Você está nos ajudando a transformar o mundo mais sustentável!] :sunglasses:")
    st.divider()
    st.subheader("ESTE FOI SEU IMPACTO :blue[POSITIVO] EM JUNHO DE 2025")
    st.markdown(f"Deixou de enviar para o aterro sanitário :green[{float(costumer_data.costumer_total):.2f}] kg :clap:")
    st.markdown(
        f"Que se transformou em :green[{(float(costumer_data.costumer_total) * 0.38):.2f}] kg de adubo orgânico :recycle:")
    st.markdown(
        f"Também foi evitado o lançamento de :green[{(float(costumer_data.costumer_total) * 0.77):.2f}] kg de CO² na atmosfera :partly_sunny:")
    st.markdown(
        f"Equivale á :green[{((float(costumer_data.costumer_total) * 0.77) / 0.096):.2f}] km rodados de carro :car:")
    st.markdown(
        f"Também equivale ao sequestro de CO² de :green[{((float(costumer_data.costumer_total) * 0.77) / 0.35):.0f}] árvores :deciduous_tree:")
    st.markdown(
        f"Com a sua iniciativa evitamos a contaminação de :green[{((float(costumer_data.costumer_total) * 0.214) * 12):.2f}] litros de água :droplet: :national_park:")


def generate_service_report():
    st.header(f":green[ NOSSOS NÚMEROS EM ]")


if __name__ == "__main__":
    file = upload_xls_file()
    if file is not None:
        # Process the uploaded XLS file
        read_xls_file(file)
