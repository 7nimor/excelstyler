import pytest
from datetime import datetime, date
import jdatetime
from excelstyler.utils import shamsi_date, convert_str_to_date


class TestShamsiDate:
    """Test cases for shamsi_date function."""
    
    def test_shamsi_date_with_datetime_object(self):
        """Test converting datetime object to Shamsi date."""
        test_date = datetime(2023, 3, 21)
        result = shamsi_date(test_date, in_value=False)
        assert result == "01-01-1402"
    
    def test_shamsi_date_with_date_object(self):
        """Test converting date object to Shamsi date."""
        test_date = date(2023, 3, 21)
        result = shamsi_date(test_date, in_value=False)
        assert result == "01-01-1402"
    
    def test_shamsi_date_in_value_true(self):
        """Test shamsi_date with in_value=True returns jdatetime object."""
        test_date = date(2023, 3, 21)
        result = shamsi_date(test_date, in_value=True)
        assert isinstance(result, jdatetime.date)
        assert result.year == 1402
        assert result.month == 1
        assert result.day == 1
    
    def test_shamsi_date_in_value_false(self):
        """Test shamsi_date with in_value=False returns string."""
        test_date = date(2023, 3, 21)
        result = shamsi_date(test_date, in_value=False)
        assert isinstance(result, str)
        assert result == "01-01-1402"
    
    def test_shamsi_date_none_input(self):
        """Test shamsi_date with None input raises ValueError."""
        with pytest.raises(ValueError, match="Date cannot be None"):
            shamsi_date(None)
    
    def test_shamsi_date_invalid_date(self):
        """Test shamsi_date with invalid date raises ValueError."""
        with pytest.raises(ValueError, match="Invalid date format"):
            shamsi_date("invalid_date")


class TestConvertStrToDate:
    """Test cases for convert_str_to_date function."""
    
    def test_convert_iso_with_milliseconds(self):
        """Test converting ISO string with milliseconds."""
        date_str = "2023-03-21T10:30:45.123Z"
        result = convert_str_to_date(date_str)
        assert result == date(2023, 3, 21)
    
    def test_convert_iso_without_milliseconds(self):
        """Test converting ISO string without milliseconds."""
        date_str = "2023-03-21T10:30:45Z"
        result = convert_str_to_date(date_str)
        assert result == date(2023, 3, 21)
    
    def test_convert_simple_date(self):
        """Test converting simple date string."""
        date_str = "2023-03-21"
        result = convert_str_to_date(date_str)
        assert result == date(2023, 3, 21)
    
    def test_convert_invalid_date(self):
        """Test converting invalid date string returns None."""
        date_str = "invalid-date"
        result = convert_str_to_date(date_str)
        assert result is None
    
    def test_convert_empty_string(self):
        """Test converting empty string returns None."""
        date_str = ""
        result = convert_str_to_date(date_str)
        assert result is None
    
    def test_convert_whitespace_string(self):
        """Test converting whitespace string returns None."""
        date_str = "   "
        result = convert_str_to_date(date_str)
        assert result is None
