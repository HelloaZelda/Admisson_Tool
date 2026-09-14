## 2024-10-27 - [Excel Formula Injection]
**Vulnerability:** Excel export features vulnerable to CSV/Formula Injection. Malicious inputs (starting with '=', '+', '-', '@') exported to Excel can execute arbitrary formulas/commands when opened by a user.
**Learning:** The application exports data containing user-controlled inputs (student names, etc.) to Excel without sanitization. Openpyxl and other libraries execute formulas starting with these characters by default.
**Prevention:** Always sanitize data being exported to spreadsheets. Prepend a single quote (`'`) to values starting with danger characters (`=`, `+`, `-`, `@`, `\t`, `\r`, `\n`) to force them to be interpreted as text.
