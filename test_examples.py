#!/usr/bin/env python3
"""
Example usage and testing script for Automatisk vakansberäkning
"""

from datetime import datetime, date, time
from vakant_karens_app import (
    OBClassifier,
    SwedishDateHelper,
    PersonnummerParser,
    load_config
)


def test_ob_classification():
    """Test OB classification with different dates and times"""
    print("=" * 60)
    print("Testing OB Classification")
    print("=" * 60)
    
    config = load_config()
    classifier = OBClassifier(config.holidays)
    
    test_cases = [
        (datetime(2025, 12, 25, 15, 0), "Helg-OB", "Christmas day afternoon"),
        (datetime(2025, 12, 23, 23, 0), "Natt", "Tuesday night"),
        (datetime(2025, 12, 26, 20, 0), "Helg-OB", "Friday evening 20:00"),
        (datetime(2025, 12, 29, 3, 0), "Helg-OB", "Monday morning 03:00 (after holiday)"),
        (datetime(2025, 12, 24, 14, 0), "Dag", "Wednesday afternoon"),
        (datetime(2025, 12, 24, 20, 30), "Kväll", "Wednesday evening 20:30"),
    ]
    
    for dt, expected, description in test_cases:
        result = classifier.classify(dt)
        status = "✓" if result == expected else "✗"
        print(f"{status} {description:40} -> {result:10} (expected: {expected})")
    
    print()


def test_personnummer_parsing():
    """Test personnummer normalization"""
    print("=" * 60)
    print("Testing Personnummer Parsing")
    print("=" * 60)
    
    test_cases = [
        ("9001011234", "199001011234", "10-digit from 1990"),
        ("5512311234", "195512311234", "10-digit from 1955"),
        ("199001011234", "199001011234", "Already 12-digit"),
    ]
    
    for input_pnr, expected, description in test_cases:
        result = PersonnummerParser.normalize(input_pnr)
        status = "✓" if result == expected else "✗"
        print(f"{status} {description:30} {input_pnr} -> {result}")
    
    print()


def test_swedish_date_helper():
    """Test Swedish month name parsing"""
    print("=" * 60)
    print("Testing Swedish Date Helper")
    print("=" * 60)
    
    month_tests = [
        ("december", 12),
        ("dec", 12),
        ("januari", 1),
        ("jan", 1),
        ("maj", 5),
    ]
    
    for month_name, expected in month_tests:
        result = SwedishDateHelper.parse_month_name(month_name)
        status = "✓" if result == expected else "✗"
        print(f"{status} {month_name:15} -> {result:2} (expected: {expected})")
    
    print()


def test_holiday_detection():
    """Test holiday detection logic"""
    print("=" * 60)
    print("Testing Holiday Detection")
    print("=" * 60)
    
    config = load_config()
    holidays = config.holidays
    
    test_dates = [
        (date(2025, 12, 25), True, "Christmas"),
        (date(2025, 12, 24), False, "Day before Christmas"),
        (date(2025, 12, 26), True, "Boxing Day"),
        (date(2026, 1, 1), True, "New Year"),
    ]
    
    for test_date, expected, description in test_dates:
        is_holiday = test_date in holidays
        is_before = SwedishDateHelper.is_day_before_holiday(test_date, holidays)
        is_after = SwedishDateHelper.is_day_after_holiday(test_date, holidays)
        
        status = "✓" if is_holiday == expected else "✗"
        print(f"{status} {test_date} ({description:20})")
        print(f"   Holiday: {is_holiday}, Before: {is_before}, After: {is_after}")
    
    print()


def example_workflow():
    """Show example workflow"""
    print("=" * 60)
    print("Example Workflow")
    print("=" * 60)
    
    print("""
# 1. Basic CLI usage:
python vakant_karens_app.py \\
  --sick_pdf Sjuklista_december_2025.pdf \\
  --payslips person1.pdf person2.pdf person3.pdf \\
  --out rapport.xlsx

# 2. With verbose logging:
python vakant_karens_app.py \\
  --sick_pdf Sjuklista_december_2025.pdf \\
  --payslips *.pdf \\
  --out rapport.xlsx \\
  --verbose

# 3. Web interface:
streamlit run vakant_karens_streamlit.py

# 4. Programmatic usage:
from vakant_karens_app import process_karens_calculation

process_karens_calculation(
    sick_pdf="sjuklista.pdf",
    payslip_paths=["p1.pdf", "p2.pdf"],
    output_xlsx="rapport.xlsx"
)
    """)


def main():
    """Run all tests"""
    print("\n")
    print("╔" + "═" * 58 + "╗")
    print("║" + " " * 6 + "Automatisk vakansberäkning - Test & Examples" + " " * 7 + "║")
    print("╚" + "═" * 58 + "╝")
    print("\n")
    
    test_ob_classification()
    test_personnummer_parsing()
    test_swedish_date_helper()
    test_holiday_detection()
    example_workflow()
    
    print("=" * 60)
    print("✓ All tests completed!")
    print("=" * 60)
    print("\nReady to process your files!")
    print("Run: streamlit run vakant_karens_streamlit.py")
    print()


if __name__ == "__main__":
    main()
