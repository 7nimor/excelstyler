import pytest
import openpyxl
from openpyxl import Workbook
from excelstyler.helpers import excel_description


class TestExcelDescription:
    """Test cases for excel_description function."""
    
    def setup_method(self):
        """Set up test workbook and worksheet."""
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
    
    def test_excel_description_basic(self):
        """Test basic description creation."""
        excel_description(self.worksheet, 'A1', 'Test Description')
        
        assert self.worksheet['A1'].value == 'Test Description'
        assert self.worksheet['A1'].alignment.horizontal == 'center'
        assert self.worksheet['A1'].alignment.vertical == 'center'
    
    def test_excel_description_with_size(self):
        """Test description creation with custom font size."""
        excel_description(self.worksheet, 'A1', 'Test Description', size=14)
        
        assert self.worksheet['A1'].value == 'Test Description'
        assert self.worksheet['A1'].font.size == 14
    
    def test_excel_description_with_color(self):
        """Test description creation with red font color."""
        excel_description(self.worksheet, 'A1', 'Test Description', color='red')
        
        assert self.worksheet['A1'].value == 'Test Description'
        assert self.worksheet['A1'].font.color.index == '00C00000'
        assert self.worksheet['A1'].font.bold == True
    
    def test_excel_description_with_merge(self):
        """Test description creation with merged cells."""
        excel_description(self.worksheet, 'A1', 'Test Description', to_row='C1')
        
        assert self.worksheet['A1'].value == 'Test Description'
        # Check if cells are merged
        merged_ranges = list(self.worksheet.merged_cells.ranges)
        assert len(merged_ranges) == 1
        assert str(merged_ranges[0]) == 'A1:C1'
    
    def test_excel_description_with_custom_color(self):
        """Test description creation with custom background color."""
        excel_description(self.worksheet, 'A1', 'Test Description', my_color='FF0000')
        
        assert self.worksheet['A1'].value == 'Test Description'
        assert self.worksheet['A1'].fill.start_color.index == '00FF0000'
