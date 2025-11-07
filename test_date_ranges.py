#!/usr/bin/env python3
"""Test date range parsing functionality."""

from models import parse_date_ranges
from excel_import import _parse_date_ranges as excel_parse_date_ranges


def test_parse_date_ranges():
    """Test the parse_date_ranges function."""

    print("Testing models.parse_date_ranges()...")
    print("-" * 60)

    # Test 1: Single date
    test1 = "25/12/2024"
    result1 = parse_date_ranges(test1)
    print(f"Input: {test1}")
    print(f"Output: {sorted(result1)}")
    assert result1 == {'2024-12-25'}, f"Expected {{'2024-12-25'}}, got {result1}"
    print("✓ Single date test passed\n")

    # Test 2: Multiple individual dates
    test2 = "25/12/2024, 26/12/2024, 01/01/2025"
    result2 = parse_date_ranges(test2)
    print(f"Input: {test2}")
    print(f"Output: {sorted(result2)}")
    assert result2 == {'2024-12-25', '2024-12-26', '2025-01-01'}, f"Expected 3 dates, got {result2}"
    print("✓ Multiple dates test passed\n")

    # Test 3: Simple date range (no spaces)
    test3 = "25/12/2024-27/12/2024"
    result3 = parse_date_ranges(test3)
    print(f"Input: {test3}")
    print(f"Output: {sorted(result3)}")
    assert result3 == {'2024-12-25', '2024-12-26', '2024-12-27'}, f"Expected 3 dates, got {result3}"
    print("✓ Date range (no spaces) test passed\n")

    # Test 4: Date range with spaces
    test4 = "25/12/2024 - 27/12/2024"
    result4 = parse_date_ranges(test4)
    print(f"Input: {test4}")
    print(f"Output: {sorted(result4)}")
    assert result4 == {'2024-12-25', '2024-12-26', '2024-12-27'}, f"Expected 3 dates, got {result4}"
    print("✓ Date range (with spaces) test passed\n")

    # Test 5: Mix of individual dates and ranges
    test5 = "25/12/2024, 28/12/2024-30/12/2024, 01/01/2025"
    result5 = parse_date_ranges(test5)
    print(f"Input: {test5}")
    print(f"Output: {sorted(result5)}")
    expected5 = {'2024-12-25', '2024-12-28', '2024-12-29', '2024-12-30', '2025-01-01'}
    assert result5 == expected5, f"Expected {expected5}, got {result5}"
    print("✓ Mixed dates and ranges test passed\n")

    # Test 6: Empty input
    test6 = ""
    result6 = parse_date_ranges(test6)
    print(f"Input: '{test6}'")
    print(f"Output: {result6}")
    assert result6 == set(), f"Expected empty set, got {result6}"
    print("✓ Empty input test passed\n")

    # Test 7: Longer range
    test7 = "01/01/2025-07/01/2025"
    result7 = parse_date_ranges(test7)
    print(f"Input: {test7}")
    print(f"Output: {sorted(result7)}")
    assert len(result7) == 7, f"Expected 7 dates, got {len(result7)}"
    print("✓ Week-long range test passed\n")

    print("\n" + "=" * 60)
    print("Testing excel_import._parse_date_ranges()...")
    print("=" * 60 + "\n")

    # Test Excel import version (returns list instead of set)
    test8 = "25/12/2024, 28/12/2024-30/12/2024"
    result8 = excel_parse_date_ranges(test8)
    print(f"Input: {test8}")
    print(f"Output: {result8}")
    assert len(result8) == 4, f"Expected 4 dates, got {len(result8)}"
    print("✓ Excel import date range test passed\n")

    print("=" * 60)
    print("ALL TESTS PASSED! ✓")
    print("=" * 60)


if __name__ == '__main__':
    test_parse_date_ranges()
