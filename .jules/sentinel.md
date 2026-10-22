## 2024-05-24 - CSV/Formula Injection in Excel Exports
**Vulnerability:** User-controlled strings (like student names or IDs) were being directly written to an Excel file export using `openpyxl.Workbook.append()`.
**Learning:** Even desktop applications are vulnerable to CSV/Formula Injection if they export user-controlled data to Excel formats (`.xlsx` or `.csv`) without escaping characters that could trigger formulas (e.g. `=`, `+`, `-`, `@`, `\t`, `\r`, `\n`).
**Prevention:** Always sanitize user input in all spreadsheet/CSV exports by prepending a single quote (`'`) to strings that start with formula-triggering characters.

## 2024-05-24 - Missed CSV Injection in Excel Exports (Numeric/Sequence Fields)
**Vulnerability:** The sequence number ("序号") and output major field ("录取专业") were exported to Excel files via `openpyxl.Workbook.append()` without sanitization. Although "序号" implies a numeric sequence, in the underlying untrusted CSV data it is parsed as an arbitrary string which could trigger CSV/Formula Injection.
**Learning:** We must sanitize ALL user-controlled string fields, including fields that appear safe or numeric-sounding (like sequence numbers or IDs) because they originate from untrusted external sources and can carry payloads.
**Prevention:** Always apply the CSV injection sanitization wrapper to all string fields extracted from untrusted CSV/Excel inputs, and output fields derived from those inputs, prior to Excel export.
## 2026-10-22 - xlwt Excel Export Formula Injection Sanitization
**Vulnerability:** User-controlled dictionary keys (majors) imported from untrusted sources were exported to Excel without sanitization.
**Learning:** When exporting data using `xlwt` to native `.xls` formats, string values are stored literally, making CSV/Formula injection protections technically unnecessary. However, project directives mandate applying `sanitize_excel_input` to all *user-controlled* Excel exports, while never sanitizing hardcoded static strings (e.g., constant headers) alongside them to avoid 'security theater'.
**Prevention:** Always apply `sanitize_excel_input` to user-controlled data exported via `xlwt` per project directive, even if technically unnecessary, while strictly avoiding sanitization of static constants.
