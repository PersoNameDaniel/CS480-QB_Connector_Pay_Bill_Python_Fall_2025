# Pay Bill Connector

## Setup Project
Once you forked and cloned the repo, run:
```bash
poetry install
```
to install dependencies.
Then write code in the src/ folder.

## Quality Check
To setup pre-commit hook (you only need to do this once):
```bash
poetry run pre-commit install
```
To manually run pre-commit checks:
```bash
poetry run pre-commit run --all-file
```
To manually run ruff check and auto fix:
```bash
poetry run ruff check --fix
```

## Build
The CLI command to build the project to a .exe run
```bash
poetry run pyinstaller --name "pay_bills" --onefile --console src\build_exe.py
```

## Run
To run the created .exe file imediately after building (assuming company data is in parent folder)
```bash
cd .\dist
.\pay_bills.exe --workbook ..\company_data.xlsx
```

If a particular sheet needs to be use --sheet <sheet name>
default is both

To change output JSON file use --output <file name>
default is report.json

To skip quickbooks use --skip-qb
default is to not skip quickbooks

Sample output:
```bash
{
  "same_records_count": 0,
  "conflicts": [
    {
      "type": "data_mismatch",
      "excel_id": "44783-46191",
      "qb_id": "44783-46191",
      "excel_date": "2024-03-22T00:00:00",
      "qb_date": "2024-03-22T00:00:00",
      "excel_amount": 14000.0,
      "qb_amount": 145500.0,
      "excel_vendor": "ATT(cell phone)",
      "qb_vendor": "ATT(cell phone)"
    }
    {
      "type": "missing_in_excel",
      "excel_id": "",
      "qb_id": "44783-46191",
      "excel_date": "",
      "qb_date": "2024-03-22T00:00:00",
      "excel_amount": ,
      "qb_amount": 145500.0,
      "excel_vendor": "",
      "qb_vendor": "ATT(cell phone)"
    }
  ],
  "added_bill_payments": [
    {
      "source": "excel",
      "id": "44092",
      "date": "2024-01-09T00:00:00",
      "amount_to_pay": 400.0,
      "vendor": "Citi Card - COSTCO"
    }
  ]
}
```
