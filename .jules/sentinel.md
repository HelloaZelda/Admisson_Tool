## 2024-05-24 - CSV/Formula Injection in Excel Exports
**Vulnerability:** User-controlled strings (like student names or IDs) were being directly written to an Excel file export using `openpyxl.Workbook.append()`.
**Learning:** Even desktop applications are vulnerable to CSV/Formula Injection if they export user-controlled data to Excel formats (`.xlsx` or `.csv`) without escaping characters that could trigger formulas (e.g. `=`, `+`, `-`, `@`, `\t`, `\r`, `\n`).
**Prevention:** Always sanitize user input in all spreadsheet/CSV exports by prepending a single quote (`'`) to strings that start with formula-triggering characters.

## 2024-05-24 - Missed CSV Injection in Excel Exports (Numeric/Sequence Fields)
**Vulnerability:** The sequence number ("序号") and output major field ("录取专业") were exported to Excel files via `openpyxl.Workbook.append()` without sanitization. Although "序号" implies a numeric sequence, in the underlying untrusted CSV data it is parsed as an arbitrary string which could trigger CSV/Formula Injection.
**Learning:** We must sanitize ALL user-controlled string fields, including fields that appear safe or numeric-sounding (like sequence numbers or IDs) because they originate from untrusted external sources and can carry payloads.
**Prevention:** Always apply the CSV injection sanitization wrapper to all string fields extracted from untrusted CSV/Excel inputs, and output fields derived from those inputs, prior to Excel export.

## 2026-10-25 - CSV/Formula Injection via Dictionary Keys in Excel Export
**Vulnerability:** Untrusted user input used as dictionary keys (e.g., majors from student choices) was being exported to an Excel file without sanitization, leading to potential CSV/Formula Injection.
**Learning:** Even when user input is aggregated and used as dictionary keys or aggregation headers, it retains its untrusted nature. When exported back to spreadsheets, these keys can still carry injection payloads.
**Prevention:** Apply input sanitization to all user-controlled data exported to spreadsheets, explicitly including fields used as dictionary keys or aggregation headers, while avoiding sanitization of hardcoded static strings.
