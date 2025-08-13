from .styles import *


def excel_description(worksheet, from_row, description, size=None, color=None, my_color=None, to_row=None):
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
