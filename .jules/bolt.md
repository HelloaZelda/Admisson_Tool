## 2025-02-18 - Missing short-circuiting on depleted quotas
**Learning:** In the `assign_admissions` algorithm, iterating through preference choices and fallback adjustments for every unassigned student takes O(N*M) operations, even after all available quotas are exhausted. This creates a severe performance bottleneck when handling applicant pools significantly larger than available spots.
**Action:** When implementing admission assignment logic or similar resource allocation algorithms, always maintain a `total_remaining` count and short-circuit the processing loop (e.g., jump to `unassigned`) once the total available resources reach zero.

## 2024-05-18 - Pandas to_dict Performance
**Learning:** Using `pandas.DataFrame.to_dict(orient="records")` has a high internal boxing overhead, making it significantly slower for large datasets compared to unboxing columns to lists.
**Action:** Unbox columns to native Python lists using `.tolist()` and zip them to form a list of dictionaries for faster DataFrame conversion.
