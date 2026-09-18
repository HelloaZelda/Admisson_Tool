## 2025-02-18 - Missing short-circuiting on depleted quotas
**Learning:** In the `assign_admissions` algorithm, iterating through preference choices and fallback adjustments for every unassigned student takes O(N*M) operations, even after all available quotas are exhausted. This creates a severe performance bottleneck when handling applicant pools significantly larger than available spots.
**Action:** When implementing admission assignment logic or similar resource allocation algorithms, always maintain a `total_remaining` count and short-circuit the processing loop (e.g., jump to `unassigned`) once the total available resources reach zero.

## 2025-02-18 - Pandas to_dict('records') overhead
**Learning:** `pandas.DataFrame.to_dict(orient="records")` is a significant performance bottleneck in this codebase due to the internal boxing overhead of converting rows into Pandas Series before creating dictionaries.
**Action:** When converting large DataFrames to lists of dictionaries, bypass Pandas by unboxing columns to native Python lists using `.tolist()` and combining them using `zip()`. This approach is much faster (2-3x speedup) and safe for mixed data types.
