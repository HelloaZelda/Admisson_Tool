## 2025-02-18 - Missing short-circuiting on depleted quotas
**Learning:** In the `assign_admissions` algorithm, iterating through preference choices and fallback adjustments for every unassigned student takes O(N*M) operations, even after all available quotas are exhausted. This creates a severe performance bottleneck when handling applicant pools significantly larger than available spots.
**Action:** When implementing admission assignment logic or similar resource allocation algorithms, always maintain a `total_remaining` count and short-circuit the processing loop (e.g., jump to `unassigned`) once the total available resources reach zero.
## 2024-05-19 - Pandas to_dict('records') Bottleneck
**Learning:** Converting large Pandas DataFrames to lists of dictionaries using `to_dict('records')` is a significant performance bottleneck due to internal boxing overhead.
**Action:** Bypass it by unboxing columns to native Python lists with `.tolist()` and combining them using `zip(*(...))` and `dict(zip(...))`. This provides a 2-3x speedup.
