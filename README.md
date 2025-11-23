# Automated Sales Reporting

## 📌 Overview
This project demonstrates how to **automate the entire monthly reporting workflow** using Python. It transforms raw sales data into a professional, presentation-ready Excel report with a single click.

Designed to simulate a real-world Admin/Analyst task, this tool replaces hours of manual copy-pasting with efficient code.

## 🚀 Key Features
1.  **Automated Data Enrichment**:
    - Calculates `Total Sales`, `Commission` (5%), and `Performance Rating` automatically.
2.  **Smart Excel Formatting**:
    - Generates **Official Excel Tables** with "Olive Green" styling.
    - Applies **Currency Formatting** (IDR) and **Center Alignment** programmatically.
    - **Auto-fits column widths** to ensure data is always readable.
3.  **Executive Summary Sheet**:
    - Creates a separate sheet with **Pivot Tables** (Sales by Region, Category, Top Products).
    - Includes **Total Rows** for quick insights.
4.  **Email Distribution Draft**:
    - Automatically generates a **ready-to-send email body** in the terminal with key metrics (Total Revenue, Top Region, etc.).

## 🛠️ Tech Stack
- **Python 3.x**
- **Pandas**: For data manipulation and aggregation.
- **OpenPyXL**: For advanced Excel formatting and styling.

## 📂 Project Structure
```text
Project_4_Automated_Reporting/
├── main.py                 # The orchestrator script
├── src/
│   ├── data_generator.py   # Creates realistic dummy data
│   └── report_generator.py # Core logic for report creation & formatting
├── data/
│   ├── input/              # Raw CSV files
│   └── output/             # Generated Excel reports
└── requirements.txt        # Project dependencies
```

## ⚡ How to Run
1.  **Install Dependencies**:
    ```bash
    pip install -r requirements.txt
    ```
2.  **Run the Automation**:
    ```bash
    python main.py
    ```
3.  **Check the Results**:
    - Open `data/output/Monthly_Report.xlsx` to see the Excel report.
    - Check your **Terminal** to see the generated Email Draft.

## 📈 Business Value
- **Efficiency**: Reduces reporting time from hours to seconds.
- **Accuracy**: Eliminates human error in calculation and formatting.
- **Consistency**: Ensures every report looks exactly the same, every month.

## 👤 Author
**Nabila Salvaningtyas**
