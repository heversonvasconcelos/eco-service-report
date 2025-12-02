import unittest
import datetime
import sys
import os

# Add the src directory to the path so we can import main
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))

from main import parse_month_string, sanitize_filename

class TestDateParsing(unittest.TestCase):

    def test_valid_date_dd_mm_yyyy(self):
        """Test parsing valid DD/MM/YYYY dates."""
        self.assertEqual(parse_month_string("01/01/2025"), datetime.datetime(2025, 1, 1))
        self.assertEqual(parse_month_string("15/06/2023"), datetime.datetime(2023, 6, 15))
        self.assertEqual(parse_month_string("31/12/2099"), datetime.datetime(2099, 12, 31))

    def test_valid_month_year_pt(self):
        """Test parsing valid 'Month Year' dates in Portuguese."""
        self.assertEqual(parse_month_string("Janeiro 2025"), datetime.datetime(2025, 1, 1))
        self.assertEqual(parse_month_string("Fevereiro 2023"), datetime.datetime(2023, 2, 1))
        self.assertEqual(parse_month_string("Dezembro 2099"), datetime.datetime(2099, 12, 1))

    def test_valid_month_year_en(self):
        """Test parsing valid 'Month Year' dates in English."""
        self.assertEqual(parse_month_string("January 2025"), datetime.datetime(2025, 1, 1))
        self.assertEqual(parse_month_string("February 2023"), datetime.datetime(2023, 2, 1))
        self.assertEqual(parse_month_string("December 2099"), datetime.datetime(2099, 12, 1))

    def test_invalid_date_values(self):
        """Test invalid date components (e.g., month 13, day 32)."""
        self.assertIsNone(parse_month_string("32/01/2025"))  # Invalid day
        self.assertIsNone(parse_month_string("01/13/2025"))  # Invalid month
        self.assertIsNone(parse_month_string("00/01/2025"))  # Invalid day 0

    def test_invalid_format_structure(self):
        """Test completely wrong formats."""
        self.assertIsNone(parse_month_string("2025/01/01")) # YYYY/MM/DD not supported by specific logic
        self.assertEqual(parse_month_string("01-01-2025"), datetime.datetime(2025, 1, 1))
        self.assertIsNone(parse_month_string("random string"))
        self.assertIsNone(parse_month_string("123"))

    def test_empty_string(self):
        """Test empty string input."""
        self.assertIsNone(parse_month_string(""))

    def test_non_alphanumeric_month_string(self):
        """Test parsing month strings with non-alphanumeric characters."""
        self.assertEqual(parse_month_string("January-2025"), datetime.datetime(2025, 1, 1))
        self.assertEqual(parse_month_string("February 2023!"), datetime.datetime(2023, 2, 1))
        self.assertEqual(parse_month_string("December_2099"), datetime.datetime(2099, 12, 1))
        self.assertEqual(parse_month_string("  MaRcH  2024  $$$"), datetime.datetime(2024, 3, 1))

    def test_sanitize_filename(self):
        """Test the sanitize_filename function with various inputs."""
        self.assertEqual(sanitize_filename("  My File Name.txt  "), "My_File_Name_txt")
        self.assertEqual(sanitize_filename("João da Silva"), "Joao_da_Silva")
        self.assertEqual(sanitize_filename("Test-File-123.pdf"), "Test_File_123_pdf")
        self.assertEqual(sanitize_filename("File with $pec!@l Ch@r$"), "File_with_pec_l_Ch_r")
        self.assertEqual(sanitize_filename("a" * 60), "a" * 50) # Test truncation
        self.assertEqual(sanitize_filename("File_with_hyphens-and_underscores"), "File_with_hyphens_and_underscores") # Hyphens should be removed, underscores kept

    def test_timestamp_input(self):
        """Test passing a datetime object directly."""
        dt = datetime.datetime(2025, 11, 28)
        expected_dt = datetime.datetime(2025, 11, 1)
        self.assertEqual(parse_month_string(dt), expected_dt)

if __name__ == '__main__':
    unittest.main()
