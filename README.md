# Bloomberg Automation

Python pipeline for reading the daily Bloomberg metadata workbook, ranking stocks by index, and generating dated output files.

## Current Status

Step 1 is complete.
- Installed and verified `xlwings`, `pandas`, `schedule`, and `openpyxl` in the local `.venv`.

Step 2 is partially complete.
- Built [Bloomberg_Data/bloomberg_pull.py](</c:/Users/shant/Desktop/Shantanu Stuff/ULJK Stuff/Bloomberg Automation/Bloomberg_Data/bloomberg_pull.py>) to read the uploaded metadata workbook.
- Recreated the ranking logic from the Excel `LET` formula using Python.
- Output currently saves the Top 10 per `Index_Clean` into `Bloomberg_Data/Daily_Snapshots/YYYY_MM_DD.csv`.
- Workbook path is intentionally left empty in code and must be passed with `--input` on the Bloomberg machine.

## Runtime Note

Do not use `/` inside a Windows filename. A file such as `Bloomberg_Daily_04/22/2026.xlsx` is invalid as a single filename on Windows.

Use either:
- `C:\Users\Shantanu\Desktop\Shantanu\Daily_File\Bloomberg_Daily_04_22_2026.xlsx`
- or a folder structure that separates the date into folders

## Current Command

```powershell
.\.venv\Scripts\python.exe Bloomberg_Data\bloomberg_pull.py --input "C:\Path\To\Bloomberg_Daily_04_22_2026.xlsx" --date 2026-04-22
```

## Output Structure

```text
Bloomberg_Data/
|-- Daily_Snapshots/
|-- Reports/
`-- bloomberg_pull.py
```

## Work Log

2026-04-22
- Inspected the uploaded metadata workbook structure.
- Confirmed the workbook contains 28 indices in one sheet using `Index_Clean`.
- Mapped Python fields to the Excel logic:
  `Company Name`, `CMP`, `FS`, `TS`, `Total`, `RSI`, `30WMA`, `ROCE FY26`, `ROE FY26`, `EPS 5Y CAGR`, `D/E`.
- Implemented Top 10 ranking by `Total` within each index.
- Generated a snapshot CSV successfully from the sample workbook.
- Updated the script so the workbook path remains blank by default and is provided only at runtime.

## Next Work

- Step 3: Compare today's snapshot with yesterday's snapshot.
- Step 4: Generate an HTML email report for rank movement.
- Step 5: Add Windows Task Scheduler setup instructions.
- Step 6: Test each component individually, then combine them.
