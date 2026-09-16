import pytest
from src.utils.security import sanitize_excel_input

def test_sanitize_excel_input():
    # Safe values
    assert sanitize_excel_input("Alice") == "Alice"
    assert sanitize_excel_input("12345") == "12345"
    assert sanitize_excel_input("") == ""
    assert sanitize_excel_input(None) is None
    assert sanitize_excel_input(100) == 100
    assert sanitize_excel_input(100.5) == 100.5

    # Unsafe values (CSV/Formula injection)
    assert sanitize_excel_input("=CMD('calc')") == "'=CMD('calc')"
    assert sanitize_excel_input("+1+1") == "'+1+1"
    assert sanitize_excel_input("-1-1") == "'-1-1"
    assert sanitize_excel_input("@SUM(A1:A10)") == "'@SUM(A1:A10)"
    assert sanitize_excel_input("\tPayload") == "'\tPayload"
    assert sanitize_excel_input("\rPayload") == "'\rPayload"
    assert sanitize_excel_input("\nPayload") == "'\nPayload"
