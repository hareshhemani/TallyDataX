# TallyDataX 🚀
> **High-Speed Tally Prime / ERP 9 Data Extractor, Tax Reconciliation & Excel Reporting Suite**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Author](https://img.shields.io/badge/Author-Haresh%20Kumar%20Hemani-orange.svg)](https://github.com/hareshhemani)
[![Website](https://img.shields.io/badge/Website-taxonline24.in-success.svg)](https://www.taxonline24.in)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-brightgreen.svg)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.100+-009688.svg)](https://fastapi.tiangolo.com/)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)]()

---

## 📌 Overview & Purpose

**TallyDataX** is an enterprise-grade automated data extraction and financial auditing tool designed for **Tally Prime** and **Tally.ERP 9**. 

In standard accounting workflows, exporting thousands of multi-currency, multi-tax vouchers with parent ledger hierarchies and addresses is notoriously slow, repetitive, and often results in broken formatting or missing fields.

**TallyDataX solves this by:**
- Connecting directly to Tally's native HTTP XML/TDL interface on port 9000.
- Extracting **25,000+ yearly vouchers** and full ledger hierarchies in **under 5 seconds**.
- Generating pixel-perfect **22-column audit-ready Excel reports** with **live dynamic formulas** (=K+O+P, =SUM(...), =SUMIF(...)).
- Providing an intuitive, modern browser-based dashboard for accountants, auditors, and business owners.
- Offering a **standalone Windows .exe** distribution so end-clients can run the tool with zero setup or Python installation.

---

## ✨ Key Features

### 1. ⚡ High-Performance XML / TDL Engine
- Communicates directly with the running Tally instance via low-overhead XML requests.
- Auto-detects currently loaded companies and active accounting periods.
- Capable of batch processing 10,000 to 25,000+ voucher records in seconds without freezing Tally.
- Offline demo mode included for interface testing when Tally is closed.

### 2. 🌳 Full Ledger Group Breadcrumb Hierarchy
- Extracts complete parent group paths for every transaction ledger:
  - *Example:* Primary > Current Liabilities > Sundry Creditors > Trade Payables
- Enables auditors and analysts to slice and categorize transactions by high-level accounting groups immediately.

### 3. 🔍 Smart Tax & GST Reconciliation
- Categorizes tax lines into dedicated columns (**CGST**, **SGST / UTGST**, **IGST**).
- Intelligent matching handles diverse ledger naming conventions:
  - Central Tax: CGST, CENTRAL GST, CENTRAL TAX
  - State / UT Tax: SGST, STATE GST, STATE TAX, UTGST
  - Integrated Tax: IGST, INTEGRATED GST, INTEGRATED TAX
- Computes GST TOTAL, TCS/ (TDS), Taxable Value, and Invoice Amount.

### 4. 📊 22-Column Excel Reports with Active Formulas
- Generates professional .xlsx workbooks using xlsxwriter and openpyxl.
- Real active Excel formulas (not hardcoded static values):
  - **Invoice Amount**: =K{row}+O{row}+P{row} (Taxable Value + GST Total + TCS/TDS)
  - **Subtotals & Grand Totals**: Active =SUM(...) formulas across all numerical columns.
- Formatted with frozen header rows, clean accounting number formatting (#,##0.00), borders, and contrasting color palettes.

### 5. 📁 Multi-File Excel Merger
- Built-in drag-and-drop tool to combine multiple Excel files (.xlsx, .xls, .csv).
- Options to merge into a single consolidated sheet or separate tabs within one workbook.

### 6. 🖥️ Modern Web UI & One-Click Desktop App
- Clean, responsive user interface built with Tailwind CSS.
- Fast client-side filtering, search, multi-ledger selector, and date range filters.
- Can run as a local web service or as a single-file portable Windows application (TallyDataX_App.exe).

---

## 📋 22-Column Report Structure

The generated Excel report and web grid include the following standard columns:

| Col # | Column Name | Description / Formula |
| :---: | :--- | :--- |
| **A** | S No | Sequential row number |
| **B** | Date | Transaction date (DD-MM-YYYY) |
| **C** | Vendor Name | Party / Supplier / Customer ledger name |
| **D** | GSTN | Party GST Identification Number |
| **E** | Voucher No. | Tally voucher number |
| **F** | Address | Party billing / mailing address |
| **G** | Invoice No | Reference supplier invoice / bill number |
| **H** | Ledger | Primary transacted ledger |
| **I** | Ledger Group | Full hierarchy path (Primary > Group > Subgroup) |
| **J** | Particulars | Item description or narration details |
| **K** | Taxable Value | Base transaction value |
| **L** | CGST | Central Goods & Services Tax |
| **M** | SGST | State / UT Goods & Services Tax |
| **N** | IGST | Integrated Goods & Services Tax |
| **O** | GST TOTAL | Total tax amount |
| **P** | TCS/ (TDS) | Tax Collected / Deducted at Source |
| **Q** | Invoice Amount | Active Formula: =K{row}+O{row}+P{row} |
| **R** | Amount paid | Payment amount cleared |
| **S** | Date of Payment | Date of payment transaction |
| **T** | Mode of Payment | Bank / Cash / Cheque / RTGS |
| **U** | Ledger balance | Running ledger balance |
| **V** | Closing Balance | Account closing balance |

---

## 🛠️ Architecture & Tech Stack

`mermaid
graph LR
    A[Tally Prime / ERP 9<br/>Port 9000 XML Server] <-->|HTTP POST XML / TDL| B[FastAPI Backend<br/>tally_client.py]
    B <-->|Excel Generation| C[Report Engine<br/>xlsxwriter / openpyxl]
    B <-->|REST API / JSON| D[Web Frontend<br/>Tailwind CSS / SheetJS]
    B -.->|PyInstaller| E[Standalone .EXE<br/>TallyDataX_App.exe]
`

- **Backend**: Python 3.10+, [FastAPI](https://fastapi.tiangolo.com/), [Uvicorn](https://www.uvicorn.org/)
- **Tally Connector**: 
equests, native TDL collections via XML over HTTP
- **Excel Engine**: xlsxwriter (ultra-fast generation) & openpyxl (template & formula handling)
- **Frontend**: HTML5, Tailwind CSS, SheetJS (xlsx), Font Awesome
- **Packaging**: PyInstaller (single-file Windows portable executable)

---

## ⚙️ Tally Prime Configuration

Before extracting data, enable Tally Prime's XML Connectivity:

1. Open **Tally Prime**.
2. Press **F1 (Help)** -> **Settings** -> **Connectivity**.
3. Under **Client/Server configuration**:
   - Set **TallyPrime acts as**: Both (or Server)
   - Set **Port**: 9000
4. Press **Ctrl + A** to save and restart Tally Prime.
5. Open your desired company in Tally.

---

## 🚀 Getting Started

### Option A: Run via Batch File (Recommended for Windows)
Simply double-click:
`at
start_app.bat
`
This launcher checks Python dependencies, starts the FastAPI server, and opens http://127.0.0.1:8000 in your default browser.

---

### Option B: Run from Source

1. **Clone the repository:**
   ```bash
   git clone https://github.com/hareshhemani/tally-extractor.git
   cd tally-extractor
   ```

2. **Install dependencies:**
   ` ash
   pip install fastapi uvicorn requests openpyxl xlsxwriter python-multipart
   `

3. **Start the application:**
   `ash
   python app.py
   `

4. **Access the Dashboard:**
   Open your browser and navigate to:
   `
   http://127.0.0.1:8000
   `

---

### Option C: Build Standalone Windows Executable (.exe)

To generate a portable .exe for distribution to clients:

`ash
pip install pyinstaller
python build_exe.py
`
The compiled output will be generated in:
- dist/TallyDataX_App.exe
- dist/TallyDataX_App.zip

Clients can run this .exe directly without installing Python or any dependencies.

---

## 📂 Project Structure

```
tallyExtractor/
├── app.py              # FastAPI server, REST routes & browser auto-launcher
├── tally_client.py     # Tally XML/TDL protocol client & voucher parsing engine
├── report_engine.py    # High-speed Excel generator (xlsxwriter & openpyxl)
├── index.html          # Single-page frontend dashboard (Tailwind CSS)
├── build_exe.py        # PyInstaller packaging script with dependency bundling
├── start_app.bat       # Windows 1-click execution script
├── app_icon.ico        # Application desktop icon
├── app_icon.png        # Web app logo & favicon
├── README.md           # Project documentation and user guide
├── LICENSE             # MIT Open Source License
├── COPYRIGHT.md        # Copyright and trademark disclaimers
└── .gitignore          # Git exclusion rules
```

---

## 👨‍💻 Author & Developer Credits

- **Author & Lead Developer**: **Haresh Kumar Hemani**
- **Website**: [www.taxonline24.in](https://www.taxonline24.in)
- **GitHub Profile**: [@hareshhemani](https://github.com/hareshhemani)
- **Official Contact**: [contact@taxonline24.in](mailto:contact@taxonline24.in)
- **Role**: Software Architecture, Tally XML/TDL Protocol Engine, High-Speed Excel Automation & Dashboard UI
- **Contributions & Support**: Open for feature requests, issues, and collaborations.

---

## 📄 License & Disclaimer

- **License**: Distributed under the **MIT License**. Copyright (c) 2026 **Haresh Kumar Hemani**. See [LICENSE](LICENSE) for details.
- **Trademark Notice**: *Tally*, *TallyPrime*, and *Tally.ERP 9* are registered trademarks of **Tally Solutions Pvt. Ltd.** This project is an independent open-source utility and is not affiliated with or endorsed by Tally Solutions Pvt. Ltd. See [COPYRIGHT.md](COPYRIGHT.md).
