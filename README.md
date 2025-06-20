# ♻️ Eco Service Report Generator

[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Framework](https://img.shields.io/badge/Framework-Streamlit-red)](https://streamlit.io)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A web application built with Streamlit to automatically generate monthly environmental impact reports for customers from an XLS or XLSX file. This tool helps businesses quantify and communicate the positive environmental impact of their services, such as waste recycling or composting.

---

## ✨ Features

- **Easy File Upload:** Upload customer data via XLS/XLSX files.
- **Intuitive Column Mapping:** Map your spreadsheet columns to required fields.
- **Automated Impact Calculations:** Calculates:
  - Organic fertilizer produced (kg)
  - CO₂ emissions avoided (kg)
  - Equivalent distance not driven by car (km)
  - Equivalent number of trees sequestering CO₂
  - Water saved (liters)
- **Dynamic Report Generation:** Creates individual, styled HTML or SVG reports for each customer.
- **Live Preview:** Instantly preview a selected customer's report in the app.
- **Bulk Download:** Download all reports in a single ZIP file.
- **Localized:** Report dates are translated to Portuguese (Brazil).

---

## 📋 Requirements

- Python 3.8+
- pip

---

## 🚀 Installation & Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/heversonvasconcelos/eco-service-report.git
   cd eco-service-report
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application:**
   ```bash
   streamlit run src/main.py
   ```

---

## 📝 Usage

1. Upload your XLS/XLSX file with customer data.
2. Map the columns to the required fields (ID, Name, Total).
3. Review and preview individual reports.
4. Download all reports as a ZIP file.

---

## 📁 Project Structure

- `src/main.py` — Main Streamlit application.
- `templates/` — HTML and SVG report templates.
- `requirements.txt` — Python dependencies.

---

## 📄 License

This project is licensed under the MIT License.