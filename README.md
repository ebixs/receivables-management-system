<div align="center">

# 🌳 Receivables Management & Reconciliation System

### A Python-based receivables management and reconciliation system with Excel automation, business-rule processing, reporting, and a management dashboard.

</div>

## Overview

This project provides a practical workflow for processing receivables data from Excel, applying reconciliation and business rules, producing structured Excel outputs, and reviewing results through a management dashboard.

The public repository contains the application source code and dashboard components. Real operational/customer data is intentionally excluded.

## Main Components

| Component | Purpose |
|---|---|
| **Desktop Application** | Simple GUI for selecting an Excel input file, running the processing engine, opening the latest output, and launching the dashboard. |
| **Python Engine** | Core reconciliation and business-rule processing implemented in Python. |
| **Management Dashboard** | Standalone HTML dashboard for analysis, reporting, and visualization of processed results. |

## Repository Structure

```text
receivables-management-system/
│
├── README.md
│
├── python/
│   ├── app_desktop.py
│   ├── wosool_engine.py
│   ├── requirements.txt
│   ├── build_exe.bat
│   └── app_icon.ico
│
├── dashboard/
│   └── dashboard_modiriati_FINAL.html
│
└── dashboard-modiriati/
    ├── app.js
    ├── template.html
    ├── dashboard_modiriati_FINAL.html
    └── logo_base64.txt
```

## Features

- Excel-based receivables processing
- Reconciliation of installment/term records
- Business-rule processing for new, paid, duplicate, and special-status records
- Automated Excel output generation
- Secondary reporting sheets
- Desktop GUI built with Tkinter
- Standalone management dashboard
- Persian-language reporting and visualization
- Windows executable build workflow with PyInstaller
- Separation of source code from operational data

## Technology Stack

- Python
- pandas
- openpyxl
- Tkinter
- HTML / CSS / JavaScript
- Chart.js
- SheetJS
- PyInstaller

## Requirements

For development/building:

- Windows
- Python 3.9+
- pip

Install Python dependencies:

```bash
pip install -r python/requirements.txt
```

## Build the Windows Application

From the `python` directory on Windows:

```text
build_exe.bat
```

The script installs the required packages, builds the executable with PyInstaller, and assembles a `release` directory containing the application and dashboard.

## Using the Application

1. Start the desktop application.
2. Select the source Excel workbook.
3. Run the reconciliation process.
4. Review the generated log and Excel output.
5. Open the management dashboard and load the generated result when required.

> Operational Excel files containing real customer or financial information are not included in this public repository.

## Data & Privacy

This repository is intended as a **public software/project showcase**.

Do not commit:

- Real customer/student records
- Telephone numbers or other personal identifiers
- Financial transaction data
- Internal operational Excel workbooks
- Generated reports containing real data
- Local configuration or temporary files

Use anonymized/sample data for demonstrations.

## Business Logic

The processing engine applies project-specific reconciliation rules to identify changes between source datasets, preserve required records, and generate reporting outputs.

The exact operational rules are implemented in `python/wosool_engine.py`.

## Dashboard

The dashboard is a standalone HTML interface that reads processed Excel output locally in the browser and provides management-oriented views such as:

- receivables summaries
- installment counts
- amounts
- status breakdowns
- personnel/coach workload
- monthly trends
- reconciliation-related indicators

## Roadmap

- [x] Desktop GUI
- [x] Excel processing engine
- [x] Management dashboard
- [x] Windows executable build
- [ ] Database-backed storage
- [ ] Automated scheduled reporting
- [ ] Additional tests and CI validation

## Author

**ابراهیم سلیمی**

