## 2024-05-24 - CSV/Formula Injection in Excel Exports
**Vulnerability:** User-controlled strings (like student names or IDs) were being directly written to an Excel file export using `openpyxl.Workbook.append()`.
**Learning:** Even desktop applications are vulnerable to CSV/Formula Injection if they export user-controlled data to Excel formats (`.xlsx` or `.csv`) without escaping characters that could trigger formulas (e.g. `=`, `+`, `-`, `@`, `\t`, `\r`, `\n`).
**Prevention:** Always sanitize user input in all spreadsheet/CSV exports by prepending a single quote (`'`) to strings that start with formula-triggering characters.

## 2024-05-24 - Missed CSV Injection in Excel Exports (Numeric/Sequence Fields)
**Vulnerability:** The sequence number ("序号") and output major field ("录取专业") were exported to Excel files via `openpyxl.Workbook.append()` without sanitization. Although "序号" implies a numeric sequence, in the underlying untrusted CSV data it is parsed as an arbitrary string which could trigger CSV/Formula Injection.
**Learning:** We must sanitize ALL user-controlled string fields, including fields that appear safe or numeric-sounding (like sequence numbers or IDs) because they originate from untrusted external sources and can carry payloads.
**Prevention:** Always apply the CSV injection sanitization wrapper to all string fields extracted from untrusted CSV/Excel inputs, and output fields derived from those inputs, prior to Excel export.

## 2026-10-23 - CSV/Formula Injection in Dictionary Keys used in Excel Exports
**Vulnerability:** The major name (`major`), which is derived from untrusted user input and used as a dictionary key for aggregation, was exported to an Excel file (`xlwt`) without sanitization.
**Learning:** Dictionary keys or aggregation headers imported from untrusted sources can carry injection payloads. Even though `xlwt` exports native `.xls` formats where strings are stored literally, project directives require all user-controlled data to be sanitized.
**Prevention:** Always apply the CSV injection sanitization wrapper (`sanitize_excel_input`) to all user-controlled string fields, including those used as dictionary keys or aggregation headers.
