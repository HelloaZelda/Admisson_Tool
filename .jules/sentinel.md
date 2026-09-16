## 2024-05-24 - CSV/Formula Injection in Excel Exports
**Vulnerability:** User-controlled strings (like student names or IDs) were being directly written to an Excel file export using `openpyxl.Workbook.append()`.
**Learning:** Even desktop applications are vulnerable to CSV/Formula Injection if they export user-controlled data to Excel formats (`.xlsx` or `.csv`) without escaping characters that could trigger formulas (e.g. `=`, `+`, `-`, `@`, `\t`, `\r`, `\n`).
**Prevention:** Always sanitize user input in all spreadsheet/CSV exports by prepending a single quote (`'`) to strings that start with formula-triggering characters.
