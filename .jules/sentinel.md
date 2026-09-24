## 2024-05-24 - CSV/Formula Injection in Excel Exports
**Vulnerability:** User-controlled strings (like student names or IDs) were being directly written to an Excel file export using `openpyxl.Workbook.append()`.
**Learning:** Even desktop applications are vulnerable to CSV/Formula Injection if they export user-controlled data to Excel formats (`.xlsx` or `.csv`) without escaping characters that could trigger formulas (e.g. `=`, `+`, `-`, `@`, `\t`, `\r`, `\n`).
**Prevention:** Always sanitize user input in all spreadsheet/CSV exports by prepending a single quote (`'`) to strings that start with formula-triggering characters.

## 2024-09-24 - CSV/Formula Injection in Exported Fields
**Vulnerability:** The fields `序号` and `录取专业` were not being sanitized before being exported to Excel via `ws.append`.
**Learning:** Numeric-sounding fields (like sequence numbers `序号`) and optional text fields imported from untrusted sources are vulnerable as they are parsed as arbitrary strings and must be sanitized.
**Prevention:** Include all user-controlled strings exported to spreadsheets to be sanitized by `sanitize_excel_input` to avoid CSV/Formula injections.
