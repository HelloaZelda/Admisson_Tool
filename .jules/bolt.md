
## 2024-05-24 - Exhausted Quota Bottleneck
**Learning:** In admission allocation algorithms, when the pool of applicants significantly exceeds the available quotas (a common real-world scenario), continuing to iterate over all majors for 'adjustment' assignments is a silent O(N*M) bottleneck. Even a simple `for major, q in list(remaining.items()):` becomes extremely costly when repeated 100k times after quotas are already zero.
**Action:** Always maintain a `total_remaining` counter alongside individual quotas in assignment algorithms. Use it to short-circuit preference loops and adjustment loops completely when no slots are left. Additionally, avoid wrapping `.items()` in `list()` inside high-frequency loops as the list construction overhead is non-negligible.
