## 2024-05-24 - CSV/Formula Injection in Excel Exports
**Vulnerability:** User-controlled strings (like student names or IDs) were being directly written to an Excel file export using `openpyxl.Workbook.append()`.
**Learning:** Even desktop applications are vulnerable to CSV/Formula Injection if they export user-controlled data to Excel formats (`.xlsx` or `.csv`) without escaping characters that could trigger formulas (e.g. `=`, `+`, `-`, `@`, `\t`, `\r`, `\n`).
**Prevention:** Always sanitize user input in all spreadsheet/CSV exports by prepending a single quote (`'`) to strings that start with formula-triggering characters.

## 2026-09-17 - Unsanitized ID Fields in Excel Exports
**Vulnerability:** The "Sequence/ID" (`序号`) field from imported CSV/Excel data was directly appended to the output Excel without sanitization, allowing formula injection if the ID contains malicious spreadsheet commands.
**Learning:** Developers often assume fields labeled as "IDs", "Sequence", or "Numbers" only contain integers and omit sanitizing them. However, when these fields are imported from untrusted CSVs or spreadsheets, they can contain arbitrary string payloads.
**Prevention:** Treat ALL user-provided data from external files as untrusted strings, even if they represent logical integers/IDs. Apply `sanitize_excel_input` to them before exporting back to Excel.
