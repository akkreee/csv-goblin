# Simple Database Handling System

A small desktop application for managing records through a Tkinter interface. Records are stored locally in `records.csv`, so the project does not require a database server or SQL.

## Features

- Add records and prevent duplicate IDs
- Select a record to load it into the form, then update it
- Delete a record by ID or from the table's right-click menu
- Search all fields as you type; searches with multiple words require every word to match
- Show all records and open the CSV file in your system's default application
- Export records to an Excel `.xlsx` file
- Customize the fields and colors in the source code

## Requirements

- Python 3
- Tkinter (included with most standard Python installations; on some Linux distributions it may need to be installed separately)
- `openpyxl` for Excel export (optional)

Install the optional Excel dependency with:

```bash
python -m pip install openpyxl
```

## Run

Clone or download this repository, open a terminal in the project directory, and run:

```bash
python dm.py
```

On Windows, you can also use `py dm.py` if `python` is not the Python launcher command on your system.

The application reads and writes `records.csv` beside `dm.py`. If the file is missing, it creates one with the configured column headers.

## Customize the fields

In `dm.py`, edit the `FIELDS` list near the top of the file:

```python
FIELDS = ["ID", "Name", "Class", "Phone", "Email"]
```

The first field is the unique key used to identify records. Keep the CSV header in `records.csv` consistent with the configured fields. Back up the CSV before changing its structure if it contains records you need to keep.

## Data and privacy

All records are stored in plain text in `records.csv` on the local machine. The application does not encrypt or transmit this data. Avoid committing real personal or sensitive records to a public repository; use a sanitized sample CSV when publishing the project.

## Project files

- `dm.py` — Tkinter application and CSV data handling
- `records.csv` — local data file with the configured headers

