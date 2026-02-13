from __future__ import annotations

from typing import Any, Dict

from src.core.admission import assign_admissions
from src.core.preferences import PREFERENCE_MAPPING


class AdmissionAlgorithm:
    # Backwards-compatible alias.
    MAJOR_MAPPING = PREFERENCE_MAPPING
    
    def __init__(self, quotas):
        self.quotas = quotas.copy()
        self.remaining_quotas = quotas.copy()
    
    def process_admissions(self, student_data):
        """
        Process student admissions based on their ranking and preferences.
        
        Args:
            student_data (pd.DataFrame): DataFrame containing student information
                with columns: ['学号', '排名', '志愿选择']
        
        Returns:
            pd.DataFrame: DataFrame with admission results
        """
        # Ranking: smaller is better.
        sorted_students = student_data.sort_values("排名")
        rows = sorted_students.to_dict(orient="records")
        quotas: Dict[str, int] = {k: int(v) for k, v in self.remaining_quotas.items()}
        result = assign_admissions(
            rows,
            quotas,
            PREFERENCE_MAPPING,
            score_key="排名",
            sort_desc=False,
            choice_key="志愿选择",
        )

        self.remaining_quotas = result.remaining_quotas.copy()

        try:
            import pandas as pd  # type: ignore

            return pd.DataFrame(result.students)
        except Exception:
            # Fallback: preserve a useful structure even without pandas.
            return result.students
    
    def get_remaining_quotas(self):
        """Return the remaining quotas for each major."""
        return self.remaining_quotas.copy()
    
    def reset_quotas(self):
        """Reset the remaining quotas to their original values."""
        self.remaining_quotas = self.quotas.copy() 
