# Excel Updater (Python + Excel Automation)

This Python script automatically updates a master Excel file using data from a weekly dump.

## Features
- Filters current month's records
- Updates "State" to "Closed"
- Handles Parent Incident logic
- Uses pandas and openpyxl

## How to Run

```bash
pip install pandas openpyxl
python update_excel.py
