## 2025-02-18 - Missing short-circuiting on depleted quotas
**Learning:** In the `assign_admissions` algorithm, iterating through preference choices and fallback adjustments for every unassigned student takes O(N*M) operations, even after all available quotas are exhausted. This creates a severe performance bottleneck when handling applicant pools significantly larger than available spots.
**Action:** When implementing admission assignment logic or similar resource allocation algorithms, always maintain a `total_remaining` count and short-circuit the processing loop (e.g., jump to `unassigned`) once the total available resources reach zero.

## 2025-02-18 - Pandas to_dict('records') conversion bottleneck
**Learning:** Using `pandas.DataFrame.to_dict(orient="records")` is surprisingly slow for large DataFrames because of internal type boxing overhead and slow Python iterations within the Pandas library. Profiling showed this step taking ~80% of total algorithm time.
**Action:** Unboxing columns to native python lists using `tolist()` and zipping them using `(dict(zip(columns, row)) for row in zip(*[df[col].tolist() for col in columns]))` completely bypasses Pandas' slow internal row conversion, yielding a 2-3x speedup. Use this pattern when you must convert large pandas structures back to list of dicts.
