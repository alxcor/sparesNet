# SparesNet

![SparesNet Banner](images/header.png)

> **SparesNet** is a portable desktop application designed to evaluate the life-cycle status and availability of Siemens Industry automation products using their part numbers (MLFB).

Starting from the older macro-enabled Excel solutions (`SparesWeb.xls`, no longer operational), **SparesNet** provides a modern standalone interface powered by Electron/Vite/Puppeteer and dynamic spreadsheets.

Web Page: [alxcor.github.io/sparesnet](https://alxcor.github.io/sparesnet)

---

## Interface Overview

![SparesNet Main Interface](images/sparesnet_workspace.png)

- **Clear All**: Wipe either the entire spreadsheet dataset
- **Clear Data**: Delete data cells content while retaining original codes (first column).
- **Import from Excel**: Import data from an existing Excel file (only from 'Data' spreadsheet, if available, first spreadsheet if not).
- **Export to Excel**: Export data into a new Excel file.
- **Read All with successors (Wizard)**: Sort all rows by MLFB codes, read data for all rows, check for successors, read data for successors also, generate and display a report.
- **Sort by MLFB**: Sort all rows by MLFB codes.
- **Sort by Column**: Sort all rows by the selected column.
- **Read Row**: Read data only for current row.
- **Read All**: Read data only for all rows (without adding rows for successors).
- **Successor for Row**: Check current row and, if a successor is available but not present in grid, inserts a new row successor and reads data from web.
- **Successor for All**: Check all rows for successors, insert a new rows (if needed) and reads data for all.
- **Preview report**: Generate a HTML report for current data.
- **Export report to HTML**: Generate a HTML report for current data and save it.
- **Export report to PDF**: Generate a HTML report for current data and save it as PDF file.
- **Spares On Web**: Quick-launch button to jump straight into the component's entry page on *Siemens Spares on Web (SOW)*
- **Industry Mall**: Quick-launch button to jump straight into the component's entry page on *Siemens Industry Mall*
- **SIOS**: Quick-launch button to jump straight into the component's entry page on *Siemens Industry Online Support (SIOS)*

---

## How To Use

1. **Input Part Numbers**: Input component codes (MLFB) into the first column of the grid spreadsheet area manually, or load them from an existing Excel via **Import from Excel**.
2. **Get Data from SIOS**: Click **Read All with successors** to start a complete background download queue, search for successors, add successors to grid, download data also for successors, then generate a report.

---

## Change Log

- **v26.06.20 (June 2026)**: Electron/Vite/Puppeteer platform transition. Introduced dark theme overrides, fully responsive workspace sizing computations, and a unified wizard execution sequence.