## 2024-05-24 - Tkinter Treeview Bulk Deletion
**Learning:** In Tkinter, iterating over `Treeview.get_children()` and deleting items one by one in a Python loop (`for item in tree.get_children(): tree.delete(item)`) is extremely slow for large datasets because of repeated C/Python boundary crossings and UI updates.
**Action:** Always use the unpacked bulk deletion method: `tree.delete(*tree.get_children())` which passes all IDs in a single call to the underlying Tcl interpreter and operates significantly faster (O(1) from the Python side).

## 2024-05-24 - openpyxl Lazy Reading Memory Optimization
**Learning:** `openpyxl`'s default `load_workbook` instantiates heavy `Cell` objects for every cell, consuming massive amounts of memory and time for large sheets (e.g. 10k rows took ~34MB and 1.8s). Using `read_only=True` alongside `values_only=True` in `iter_rows` turns it into a lazy generator that yields basic Python types, dropping memory to ~2MB and execution time by ~30%. However, `read_only=True` keeps the file handle open, so `wb.close()` must be called explicitly when done.
**Action:** For read-only operations on large Excel files, always use `read_only=True` with `values_only=True` and ensure the workbook is properly closed afterwards.
