## 2024-05-24 - CSV/Formula Injection in Excel Exports
**Vulnerability:** User-controlled strings (like student names or IDs) were being directly written to an Excel file export using `openpyxl.Workbook.append()`.
**Learning:** Even desktop applications are vulnerable to CSV/Formula Injection if they export user-controlled data to Excel formats (`.xlsx` or `.csv`) without escaping characters that could trigger formulas (e.g. `=`, `+`, `-`, `@`, `\t`, `\r`, `\n`).
**Prevention:** Always sanitize user input in all spreadsheet/CSV exports by prepending a single quote (`'`) to strings that start with formula-triggering characters.

## 2024-05-27 - [CSV/Formula Injection in Numeric-Sounding Excel Exports]
**Vulnerability:** Application exports data to Excel and is vulnerable to CSV/Formula Injection. Certain fields like `序号` (Sequence Number) and `分数` (Score) were not sanitized because they are numeric-sounding, but they can still contain injection payloads.
**Learning:** All user-controlled strings exported to spreadsheets must be sanitized, regardless of the column header implying it contains a number. Untrusted input will be parsed as arbitrary strings.
**Prevention:** Use a helper function (like `sanitize_excel_input`) consistently on ALL fields being exported to Excel when constructing rows.
