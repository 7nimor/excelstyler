import pytest
from excelstyler.to_locale_string import to_locale_str


class TestToLocaleStr:
    """Test cases for to_locale_str function."""
    
    def test_to_locale_str_integer(self):
        """Test converting integer to locale string."""
        result = to_locale_str(1234567)
        assert result == "1,234,567"
    
    def test_to_locale_str_float(self):
        """Test converting float to locale string."""
        result = to_locale_str(1234567.89)
        assert result == "1,234,567"
    
    def test_to_locale_str_small_number(self):
        """Test converting small number to locale string."""
        result = to_locale_str(123)
        assert result == "123"
    
    def test_to_locale_str_zero(self):
        """Test converting zero to locale string."""
        result = to_locale_str(0)
        assert result == "0"
    
    def test_to_locale_str_negative_number(self):
        """Test converting negative number to locale string."""
        result = to_locale_str(-1234567)
        assert result == "-1,234,567"
    
    def test_to_locale_str_none_input(self):
        """Test to_locale_str with None input raises ValueError."""
        with pytest.raises(ValueError, match="Input cannot be None"):
            to_locale_str(None)
    
    def test_to_locale_str_string_input(self):
        """Test to_locale_str with string input raises ValueError."""
        with pytest.raises(ValueError, match="Cannot convert"):
            to_locale_str("not_a_number")
    
    def test_to_locale_str_list_input(self):
        """Test to_locale_str with list input raises ValueError."""
        with pytest.raises(ValueError, match="Cannot convert"):
            to_locale_str([1, 2, 3])
