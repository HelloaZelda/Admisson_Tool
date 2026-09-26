## 2024-05-24 - CSV/Formula Injection in Excel Exports
**Vulnerability:** User-controlled strings (like student names or IDs) were being directly written to an Excel file export using `openpyxl.Workbook.append()`.
**Learning:** Even desktop applications are vulnerable to CSV/Formula Injection if they export user-controlled data to Excel formats (`.xlsx` or `.csv`) without escaping characters that could trigger formulas (e.g. `=`, `+`, `-`, `@`, `\t`, `\r`, `\n`).
**Prevention:** Always sanitize user input in all spreadsheet/CSV exports by prepending a single quote (`'`) to strings that start with formula-triggering characters.

## 2024-05-24 - CSV Injection in Numeric/Arbitrary Fields
**Vulnerability:** The fields `student['序号']` and `student.get('录取专业', '')` in `src/gui/simple_main.py` were not sanitized before being exported to Excel via `ws.append()`.
**Learning:** Numeric-sounding fields (like sequence numbers or IDs) imported from untrusted sources, or derived arbitrary fields, can be parsed as strings in spreadsheet software. If they contain characters like `=`, `+`, `-`, or `@`, they can trigger formula execution.
**Prevention:** Always sanitize ALL string or potentially string-like fields when exporting to Excel or CSV, not just obvious text fields. Apply sanitization consistently to all user-controlled or derived data in the export function.
