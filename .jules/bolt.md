## 2025-02-18 - Missing short-circuiting on depleted quotas
**Learning:** In the `assign_admissions` algorithm, iterating through preference choices and fallback adjustments for every unassigned student takes O(N*M) operations, even after all available quotas are exhausted. This creates a severe performance bottleneck when handling applicant pools significantly larger than available spots.
**Action:** When implementing admission assignment logic or similar resource allocation algorithms, always maintain a `total_remaining` count and short-circuit the processing loop (e.g., jump to `unassigned`) once the total available resources reach zero.
## 2024-05-19 - Pandas to_dict('records') Overhead
**Learning:** `pandas.DataFrame.to_dict(orient='records')` introduces significant boxing overhead (2-3x slower) compared to extracting columns natively when converting large DataFrames to lists of dictionaries in this codebase.
**Action:** Use `.tolist()` to extract raw lists of column data and combine them natively with `zip()` and `dict()` comprehensions when mapping large DataFrames to lists of dictionaries for performance-critical logic.
