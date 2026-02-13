"""
Core admission assignment logic.

This module is intentionally UI-agnostic so it can be reused by tkinter GUI,
tests, and any future CLI/API.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Mapping, Optional


@dataclass(frozen=True)
class AdmissionResult:
    students: List[Dict[str, Any]]
    remaining_quotas: Dict[str, int]

ADJUST_SUFFIX = "(调剂)"
INVALID_CHOICE_LABEL = "无效志愿"
UNASSIGNED_LABEL = "未分配"


def _norm_choice(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().upper()


def assign_admissions(
    students: Iterable[Mapping[str, Any]],
    quotas: Mapping[str, int],
    preference_mapping: Mapping[str, List[str]],
    *,
    score_key: str = "分数",
    sort_desc: bool = True,
    choice_key: str = "志愿选择",
    assigned_key: str = "录取专业",
    adjust_suffix: str = ADJUST_SUFFIX,
    invalid_choice_label: str = INVALID_CHOICE_LABEL,
    unassigned_label: str = UNASSIGNED_LABEL,
) -> AdmissionResult:
    """
    Assign admissions for students sorted by score (descending).

    Rules:
    - If choice code invalid => 标记为 invalid_choice_label
    - Else try 1st/2nd/3rd preference in order; assign first with remaining quota
    - Else try adjustment into any major with remaining quota (append adjust_suffix)
    - Else mark unassigned_label
    """

    remaining: Dict[str, int] = {k: int(v) for k, v in quotas.items()}

    # Copy input students into mutable dicts so callers can pass in mapping/rows safely.
    items: List[Dict[str, Any]] = [dict(s) for s in students]

    def score_of(s: Mapping[str, Any]) -> float:
        try:
            return float(s.get(score_key, 0))
        except Exception:
            return 0.0

    items.sort(key=score_of, reverse=sort_desc)

    for s in items:
        choice = _norm_choice(s.get(choice_key))

        # Distinguish between "blank choice" and "invalid code".
        # Blank: treat as no preferences, but still eligible for adjustment.
        # Invalid: mark explicitly.
        if choice and choice not in preference_mapping:
            s[assigned_key] = invalid_choice_label
            continue

        assigned_major: Optional[str] = None
        if choice:
            for major in preference_mapping[choice]:
                if remaining.get(major, 0) > 0:
                    remaining[major] -= 1
                    assigned_major = major
                    break

        if assigned_major is not None:
            s[assigned_key] = assigned_major
            continue

        # Adjustment: any remaining slot.
        for major, q in list(remaining.items()):
            if q > 0:
                remaining[major] -= 1
                s[assigned_key] = f"{major}{adjust_suffix}"
                break
        else:
            s[assigned_key] = unassigned_label

    return AdmissionResult(students=items, remaining_quotas=remaining)
