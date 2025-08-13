from .styles import *


def excel_description(worksheet, from_row, description, size=None, color=None, my_color=None, to_row=None):
    """
    Write a description or label in an Excel worksheet, optionally merge cells and apply styling.

    Parameters:
    -----------
    worksheet : openpyxl.Worksheet
        The Excel worksheet where the description will be written.
    from_row : str
        The cell reference where the description starts (e.g., 'A1').
    description : str
        The text to write in the cell.
    size : int, optional
        Font size for the text.
    color : str, optional
        Use predefined color for the font (e.g., red_font).
    my_color : str, optional
        Custom background color for the cell (hex format).
    to_row : str, optional
        If provided, merge cells from `from_row` to `to_row`.

    Notes:
    ------
    - The text will be centered using Alignment_CELL.
    - If `to_row` is given, the range from `from_row` to `to_row` will be merged.
    - Either `color` or `my_color` can be used to apply font or background color respectively.

    Example:
    --------
    excel_description(worksheet, 'A1', 'Cold House Report', size=14, color='red', to_row='C1')
    """
    worksheet[from_row] = description
    worksheet[from_row].alignment = Alignment_CELL
    if size is not None:
        worksheet[from_row].font = Font(size=size)
    if color is not None:
        worksheet[from_row].font = red_font
    if my_color is not None:
        worksheet[from_row].font = PatternFill(start_color=my_color, fill_type="solid")

    if to_row is not None:
        merge_range = f'{from_row}:{to_row}'
        worksheet.merge_cells(merge_range)
