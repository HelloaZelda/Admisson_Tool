## 2025-02-18 - Missing short-circuiting on depleted quotas
**Learning:** In the `assign_admissions` algorithm, iterating through preference choices and fallback adjustments for every unassigned student takes O(N*M) operations, even after all available quotas are exhausted. This creates a severe performance bottleneck when handling applicant pools significantly larger than available spots.
**Action:** When implementing admission assignment logic or similar resource allocation algorithms, always maintain a `total_remaining` count and short-circuit the processing loop (e.g., jump to `unassigned`) once the total available resources reach zero.
## 2024-03-24 - Pandas to Dict Boxing Overhead
**Learning:** `DataFrame.to_dict(orient="records")` has significant boxing overhead, causing a severe performance bottleneck.
**Action:** When converting large DataFrames to lists of dicts, bypass pandas internal boxing by unboxing columns to native Python lists with `.tolist()` and zip them manually. This is 2-3x faster.
