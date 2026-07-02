# DCSC MIS Consolidation

Python solution for the DCSC MIS Consolidation coding assignment.

## Features

- Recursively scans the `Export/` and `Import/` folders for `.xlsx` files.
- Automatically detects AWB-level detail sheets even when sheet names differ (`ANNEXURE`, `ANX`, `ANNEX 01-15`, `Data_Insert`, `Summary`, `Export TP_SC`, etc.).
- Normalizes different airline-specific column names into the required common schema.
- Derives metadata including `Direction`, `Airline`, `Month`, `Fortnight`, and `Source File`.
- Calculates Handling, X-Ray, Demurrage, and GST-inclusive Total Amounts using the available rate sheets.
- Generates a consolidated workbook containing:
  - `Export`
  - `Import`
  - `Month_Summary`
  - `Reconciliation`
- Produces a JSON run log for execution details.

## Requirements

- Python 3.10+
- openpyxl

## Run

```powershell
python dcsc_mis_consolidator.py `
  --input "C:\path\to\DCSC_MIS_Sample_Pack" `
  --output "DCSC_MIS_Consolidated_Output.xlsx" `
  --log "run_log.json"
```

For this submission, the consolidated workbook has already been generated and is included in the project folder.

## Output Workbook

The generated workbook contains the following sheets:

- **README** – Generation notes.
- **Export** – Consolidated Export AWB-level records.
- **Import** – Consolidated Import AWB-level records.
- **Month_Summary** – Airline × Direction × Month totals.
- **Reconciliation** – Comparison of parsed totals against Service Challan / SC totals.

## Assumptions

- Header rows are detected dynamically and are not assumed to start at row 1.
- Trailing blank rows and TOTAL rows are excluded from consolidation.
- Fortnight information is extracted from either the filename or sheet name, depending on the airline format.
- Column mappings are normalized using configurable aliases.

## Notes

The reconciliation sheet intentionally highlights mismatches as **CHECK** instead of suppressing them. Small differences may occur because some source challans use rounded flight-level chargeable weights while AWB-level calculations use exact line-level values, or because certain airline formats may require additional airline-specific parsing logic.