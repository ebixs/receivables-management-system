# Receivables Management System

A Python-based receivables management and reconciliation system with Excel automation, business-rule processing, reporting, and a management dashboard.

> **Public portfolio version:** source code and documentation are prepared for demonstration and software-engineering review. Real customer, employee, financial, and organizational records are intentionally excluded.

## Overview

This project automates a recurring receivables workflow built around Excel-based operational data. It combines a desktop application, a reusable Python processing engine, and a browser-based management dashboard.

### Core capabilities

- Excel-based data processing with `pandas` and `openpyxl`
- Normalization and exact matching of records
- Duplicate and mismatch detection
- Business-rule based assignment and status handling
- Automated Excel output generation
- Persian/Jalali date handling
- Desktop GUI for non-technical users
- Management dashboard for summaries and reporting
- Reusable processing engine that can be called independently from the GUI

## Architecture

```
Excel input
    │
    ▼
Desktop application
    │
    ▼
Python processing engine
    │
    ├── validation & normalization
    ├── matching & reconciliation
    ├── business rules
    ├── mismatch reporting
    └── Excel output
             │
             ▼
      Management dashboard
```

## Technology Stack

- Python 3.9+
- pandas
- openpyxl
- Tkinter
- HTML / CSS / JavaScript
- Excel automation

## Project Structure

```
receivables-management-system/
├── README.md
├── .gitignore
├── python/
│   ├── app_desktop.py
│   ├── wosool_engine.py
│   ├── requirements.txt
│   └── build_exe.bat
├── dashboard/
│   └── dashboard_modiriati_FINAL.html
└── dashboard-modiriati/
    ├── app.js
    ├── dashboard_modiriati_FINAL.html
    └── template.html
```

## Running the Python Application

Create a virtual environment and install the dependencies:

```bash
python -m venv .venv

# Windows
.venv\\Scripts\\activate

# Linux / macOS
source .venv/bin/activate

pip install -r python/requirements.txt
```

Run the desktop application:

```bash
python python/app_desktop.py
```

## Data Privacy

This repository intentionally contains **no real customer records, payment files, phone numbers, personal identifiers, or operational Excel workbooks**.

For a production deployment, input/output Excel files should remain outside the Git repository and should be handled according to the organization's security and privacy requirements.

## Design Principles

- Keep business rules explicit and auditable.
- Separate data processing from the user interface.
- Preserve source data during reconciliation.
- Report mismatches instead of silently overwriting conflicting records.
- Keep operational data outside source control.

## Future Improvements

- Database-backed storage instead of operational Excel files
- Automated scheduled reporting
- Authentication and role-based access
- Automated tests and CI
- Containerized deployment

## Author

**Ebrahim Salimi**

GitHub: https://github.com/ebixs

## License

This repository is provided as a portfolio/demo project. A production license and organizational usage terms should be defined before redistribution.
