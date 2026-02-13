"""
Preference definitions shared across CLI/GUI and algorithm layers.
"""

from __future__ import annotations

from typing import Dict, List


# A-F map to ordered major preferences.
PREFERENCE_MAPPING: Dict[str, List[str]] = {
    "A": ["电子信息工程", "通信工程", "电磁场与无线技术"],
    "B": ["电子信息工程", "电磁场与无线技术", "通信工程"],
    "C": ["电磁场与无线技术", "电子信息工程", "通信工程"],
    "D": ["电磁场与无线技术", "通信工程", "电子信息工程"],
    "E": ["通信工程", "电子信息工程", "电磁场与无线技术"],
    "F": ["通信工程", "电磁场与无线技术", "电子信息工程"],
}

