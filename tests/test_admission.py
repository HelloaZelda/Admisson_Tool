from __future__ import annotations

from src.core.admission import INVALID_CHOICE_LABEL, UNASSIGNED_LABEL, assign_admissions
from src.core.preferences import PREFERENCE_MAPPING


def test_invalid_choice_marked():
    students = [{"学号": "1", "分数": 100, "志愿选择": "Z"}]
    quotas = {"电子信息工程": 1, "通信工程": 1, "电磁场与无线技术": 1}
    r = assign_admissions(students, quotas, PREFERENCE_MAPPING)
    assert r.students[0]["录取专业"] == INVALID_CHOICE_LABEL


def test_blank_choice_goes_to_adjustment_or_unassigned():
    students = [{"学号": "1", "分数": 100, "志愿选择": ""}]
    quotas = {"电子信息工程": 0, "通信工程": 0, "电磁场与无线技术": 1}
    r = assign_admissions(students, quotas, PREFERENCE_MAPPING)
    assert r.students[0]["录取专业"] == "电磁场与无线技术(调剂)"


def test_sorted_by_score_desc_and_assign_preference():
    students = [
        {"学号": "low", "分数": 10, "志愿选择": "A"},
        {"学号": "high", "分数": 99, "志愿选择": "A"},
    ]
    quotas = {"电子信息工程": 1, "通信工程": 1, "电磁场与无线技术": 1}
    r = assign_admissions(students, quotas, PREFERENCE_MAPPING)
    assert [s["学号"] for s in r.students] == ["high", "low"]
    assert r.students[0]["录取专业"] == "电子信息工程"

def test_sort_ascending_when_using_rank():
    students = [
        {"学号": "r2", "排名": 2, "志愿选择": "A"},
        {"学号": "r1", "排名": 1, "志愿选择": "A"},
    ]
    quotas = {"电子信息工程": 2, "通信工程": 0, "电磁场与无线技术": 0}
    r = assign_admissions(students, quotas, PREFERENCE_MAPPING, score_key="排名", sort_desc=False)
    assert [s["学号"] for s in r.students] == ["r1", "r2"]


def test_fallback_to_next_preferences_then_adjust_then_unassigned():
    # A: 电子信息工程 > 通信工程 > 电磁场与无线技术
    students = [
        {"学号": "s1", "分数": 100, "志愿选择": "A"},
        {"学号": "s2", "分数": 90, "志愿选择": "A"},
        {"学号": "s3", "分数": 80, "志愿选择": "A"},
        {"学号": "s4", "分数": 70, "志愿选择": "A"},
    ]
    quotas = {"电子信息工程": 1, "通信工程": 1, "电磁场与无线技术": 0}
    r = assign_admissions(students, quotas, PREFERENCE_MAPPING)

    # s1 gets first choice, s2 gets second choice, s3 has no preference slots left and adjusts nowhere, so unassigned.
    assert r.students[0]["录取专业"] == "电子信息工程"
    assert r.students[1]["录取专业"] == "通信工程"
    assert r.students[2]["录取专业"] == UNASSIGNED_LABEL
    assert r.students[3]["录取专业"] == UNASSIGNED_LABEL

    # Now with an extra slot in a non-preference remaining major, adjustment should happen.
    quotas2 = {"电子信息工程": 1, "通信工程": 1, "电磁场与无线技术": 1}
    r2 = assign_admissions(students, quotas2, PREFERENCE_MAPPING)
    assert r2.students[2]["录取专业"].endswith("(调剂)") is False  # still has 3rd preference available
    assert r2.students[2]["录取专业"] == "电磁场与无线技术"
