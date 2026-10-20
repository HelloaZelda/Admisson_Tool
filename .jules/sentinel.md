## 2024-05-24 - CSV/Formula Injection in Excel Exports
**Vulnerability:** User-controlled strings (like student names or IDs) were being directly written to an Excel file export using `openpyxl.Workbook.append()`.
**Learning:** Even desktop applications are vulnerable to CSV/Formula Injection if they export user-controlled data to Excel formats (`.xlsx` or `.csv`) without escaping characters that could trigger formulas (e.g. `=`, `+`, `-`, `@`, `\t`, `\r`, `\n`).
**Prevention:** Always sanitize user input in all spreadsheet/CSV exports by prepending a single quote (`'`) to strings that start with formula-triggering characters.

## 2024-05-24 - Missed CSV Injection in Excel Exports (Numeric/Sequence Fields)
**Vulnerability:** The sequence number ("序号") and output major field ("录取专业") were exported to Excel files via `openpyxl.Workbook.append()` without sanitization. Although "序号" implies a numeric sequence, in the underlying untrusted CSV data it is parsed as an arbitrary string which could trigger CSV/Formula Injection.
**Learning:** We must sanitize ALL user-controlled string fields, including fields that appear safe or numeric-sounding (like sequence numbers or IDs) because they originate from untrusted external sources and can carry payloads.
**Prevention:** Always apply the CSV injection sanitization wrapper to all string fields extracted from untrusted CSV/Excel inputs, and output fields derived from those inputs, prior to Excel export.
## 2026-10-20 - CSV/Formula Injection in Dictionary Keys
**Vulnerability:** Untrusted string values used as dictionary keys ('major' from Excel input) were exported to Excel without sanitization, leading to CSV/Formula Injection.
**Learning:** Aggregation keys or headers that originate from untrusted sources carry the same injection payload risk as direct user input fields when exported to native Excel files.
**Prevention:** Apply input sanitization (e.g., prepending a single quote) to all user-controlled data exported to spreadsheets, including fields used as aggregation headers or keys.
