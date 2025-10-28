# excelstyler

`excelstyler` is a Python package that makes it easy to style and format Excel worksheets using [openpyxl](https://openpyxl.readthedocs.io). It provides a simple API for creating professional-looking Excel reports with Persian/Farsi language support.

## Features

- 🎨 **Easy Styling**: Pre-defined color schemes and styling options
- 🇮🇷 **Persian Support**: Built-in support for Persian dates and RTL text
- 📊 **Charts**: Create line and bar charts with simple function calls
- 🔧 **Flexible**: Customizable headers, values, and formatting
- 🧪 **Well Tested**: Comprehensive test suite with pytest
- 📝 **Well Documented**: Clear documentation and examples

## Installation

```bash
pip install excelstyler
```

## Quick Start

```python
from openpyxl import Workbook
from excelstyler.headers import create_header
from excelstyler.values import create_value
from excelstyler.utils import shamsi_date
from datetime import datetime

# Create workbook
workbook = Workbook()
worksheet = workbook.active
worksheet.sheet_view.rightToLeft = True  # For Persian/Farsi support

# Create header
headers = ["Name", "Age", "City", "Date"]
create_header(worksheet, headers, 1, 1, color='green', height=25, width=15)

# Add data
data = [
    ["John Doe", 30, "New York", shamsi_date(datetime.now())],
    ["Jane Smith", 25, "London", shamsi_date(datetime.now())],
    ["Ali Ahmad", 35, "Tehran", shamsi_date(datetime.now())]
]

for i, row_data in enumerate(data, start=2):
    create_value(worksheet, row_data, i, 1, border_style='thin')

# Save workbook
workbook.save("report.xlsx")
```

## 📚 Complete Tutorial

### 1. Basic Setup

First, let's create a simple Excel file with basic styling:

```python
from openpyxl import Workbook
from excelstyler.headers import create_header
from excelstyler.values import create_value

# Create a new workbook
workbook = Workbook()
worksheet = workbook.active

# Set RTL for Persian/Farsi support
worksheet.sheet_view.rightToLeft = True

# Create a simple header
headers = ["نام", "سن", "شهر"]
create_header(worksheet, headers, 1, 1, color='green')

# Add some data
data = [
    ["علی احمدی", 25, "تهران"],
    ["فاطمه محمدی", 30, "اصفهان"],
    ["حسن رضایی", 35, "شیراز"]
]

for i, row_data in enumerate(data, start=2):
    create_value(worksheet, row_data, i, 1)

workbook.save("simple_report.xlsx")
```

### 2. Advanced Styling

Let's create a more sophisticated report with various styling options:

```python
from openpyxl import Workbook
from excelstyler.headers import create_header_freez
from excelstyler.values import create_value
from excelstyler.helpers import excel_description
from excelstyler.utils import shamsi_date, to_locale_str
from datetime import datetime

workbook = Workbook()
worksheet = workbook.active
worksheet.sheet_view.rightToLeft = True

# Add title
excel_description(worksheet, 'A1', 'گزارش فروش ماهانه', size=16, to_row='E1')

# Create header with freeze panes
headers = ['ردیف', 'نام محصول', 'تعداد فروش', 'قیمت واحد', 'مجموع']
create_header_freez(
    worksheet, 
    headers, 
    start_col=1, 
    row=3, 
    header_row=4,
    height=30, 
    width=18,
    color='blue',
    border_style='medium'
)

# Sample data
products = [
    ['لپ‌تاپ', 50, 15000000],
    ['موبایل', 200, 8000000],
    ['تبلت', 100, 12000000],
    ['هدفون', 300, 2000000]
]

# Add data with alternating colors
for i, (name, quantity, price) in enumerate(products, start=4):
    total = quantity * price
    row_data = [
        i-3,  # Row number
        name,
        quantity,
        to_locale_str(price),
        to_locale_str(total)
    ]
    create_value(
        worksheet, 
        row_data, 
        i, 
        1, 
        border_style='thin',
        m=i,  # For alternating colors
        height=25
    )

# Add summary
total_sales = sum(q * p for _, q, p in products)
excel_description(
    worksheet, 
    'A9', 
    f'مجموع کل فروش: {to_locale_str(total_sales)} تومان', 
    size=14, 
    color='red',
    to_row='E9'
)

workbook.save("advanced_report.xlsx")
```

### 3. Working with Charts

Create reports with visual charts:

```python
from openpyxl import Workbook
from excelstyler.headers import create_header
from excelstyler.values import create_value
from excelstyler.chart import add_chart

workbook = Workbook()
worksheet = workbook.active

# Create header
headers = ['ماه', 'فروش (میلیون تومان)']
create_header(worksheet, headers, 1, 1, color='green')

# Add data
monthly_data = [
    ['فروردین', 120],
    ['اردیبهشت', 150],
    ['خرداد', 180],
    ['تیر', 200],
    ['مرداد', 220],
    ['شهریور', 190]
]

for i, (month, sales) in enumerate(monthly_data, start=2):
    create_value(worksheet, [month, sales], i, 1, border_style='thin')

# Add a line chart
add_chart(
    worksheet=worksheet,
    chart_type='line',
    data_columns=2,  # Sales column
    category_column=1,  # Month column
    start_row=2,
    end_row=7,
    chart_position="D2",
    chart_title="نمودار فروش ماهانه",
    x_axis_title="ماه",
    y_axis_title="فروش (میلیون تومان)",
    chart_width=20,
    chart_height=12
)

workbook.save("chart_report.xlsx")
```

### 4. Persian Date Handling

Working with Persian (Shamsi) dates:

```python
from openpyxl import Workbook
from excelstyler.headers import create_header
from excelstyler.values import create_value
from excelstyler.utils import shamsi_date, convert_str_to_date
from datetime import datetime, date

workbook = Workbook()
worksheet = workbook.active
worksheet.sheet_view.rightToLeft = True

# Create header
headers = ['تاریخ میلادی', 'تاریخ شمسی (متن)', 'تاریخ شمسی (شیء)']
create_header(worksheet, headers, 1, 1, color='orange')

# Sample dates
dates = [
    datetime(2023, 3, 21),  # Nowruz
    datetime(2023, 6, 21),  # Summer solstice
    datetime(2023, 9, 23),  # Autumn equinox
    datetime(2023, 12, 21)  # Winter solstice
]

for i, gregorian_date in enumerate(dates, start=2):
    # Convert to Persian date as string
    persian_str = shamsi_date(gregorian_date, in_value=False)
    
    # Convert to Persian date as object
    persian_obj = shamsi_date(gregorian_date, in_value=True)
    
    row_data = [
        gregorian_date.strftime('%Y-%m-%d'),
        persian_str,
        persian_obj
    ]
    
    create_value(worksheet, row_data, i, 1, border_style='thin')

workbook.save("persian_dates.xlsx")
```

### 5. Conditional Formatting

Highlight specific cells based on conditions:

```python
from openpyxl import Workbook
from excelstyler.headers import create_header
from excelstyler.values import create_value

workbook = Workbook()
worksheet = workbook.active
worksheet.sheet_view.rightToLeft = True

# Create header
headers = ['نام کارمند', 'امتیاز', 'وضعیت']
create_header(worksheet, headers, 1, 1, color='green')

# Employee data
employees = [
    ['علی احمدی', 95, 'عالی'],
    ['فاطمه محمدی', 75, 'خوب'],
    ['حسن رضایی', 45, 'ضعیف'],
    ['زهرا کریمی', 88, 'عالی'],
    ['محمد نوری', 60, 'متوسط']
]

for i, (name, score, status) in enumerate(employees, start=2):
    # Highlight low scores in red
    different_cell = 1 if score < 50 else None
    different_value = 45 if score < 50 else None
    
    create_value(
        worksheet, 
        [name, score, status], 
        i, 
        1, 
        border_style='thin',
        different_cell=different_cell,
        different_value=different_value
    )

workbook.save("conditional_formatting.xlsx")
```

### 6. Complete Business Report

A comprehensive example combining all features:

```python
from openpyxl import Workbook
from excelstyler.headers import create_header_freez
from excelstyler.values import create_value
from excelstyler.helpers import excel_description
from excelstyler.chart import add_chart
from excelstyler.utils import shamsi_date, to_locale_str
from datetime import datetime

def create_sales_report():
    """Create a comprehensive sales report"""
    
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.sheet_view.rightToLeft = True
    
    # Company header
    excel_description(
        worksheet, 
        'A1', 
        'شرکت فناوری پارس - گزارش فروش سه ماهه', 
        size=18, 
        to_row='F1'
    )
    
    # Report date
    excel_description(
        worksheet, 
        'A2', 
        f'تاریخ گزارش: {shamsi_date(datetime.now())}', 
        size=12, 
        to_row='F2'
    )
    
    # Data header with freeze
    headers = [
        'ردیف', 'نام محصول', 'دسته‌بندی', 'تعداد فروش', 
        'قیمت واحد', 'مجموع فروش', 'درصد از کل'
    ]
    create_header_freez(
        worksheet, 
        headers, 
        start_col=1, 
        row=4, 
        header_row=5,
        height=25, 
        width=15,
        color='blue',
        border_style='medium'
    )
    
    # Sample sales data
    sales_data = [
        ['لپ‌تاپ ایسوس', 'کامپیوتر', 25, 15000000],
        ['آیفون 14', 'موبایل', 50, 25000000],
        ['سامسونگ گلکسی', 'موبایل', 30, 20000000],
        ['مک‌بوک پرو', 'کامپیوتر', 15, 35000000],
        ['تبلت آیپد', 'تبلت', 40, 12000000],
        ['هدفون سونی', 'لوازم جانبی', 100, 3000000],
        ['کیبورد مکانیکال', 'لوازم جانبی', 80, 2000000],
        ['ماوس گیمینگ', 'لوازم جانبی', 60, 1500000]
    ]
    
    # Calculate totals
    total_sales = sum(q * p for _, _, q, p in sales_data)
    
    # Add data rows
    for i, (name, category, quantity, price) in enumerate(sales_data, start=5):
        sales_total = quantity * price
        percentage = (sales_total / total_sales) * 100
        
        row_data = [
            i-4,  # Row number
            name,
            category,
            quantity,
            to_locale_str(price),
            to_locale_str(sales_total),
            f"{percentage:.1f}%"
        ]
        
        create_value(
            worksheet, 
            row_data, 
            i, 
            1, 
            border_style='thin',
            m=i,  # Alternating colors
            height=22
        )
    
    # Add summary section
    summary_row = len(sales_data) + 6
    excel_description(
        worksheet, 
        f'A{summary_row}', 
        f'مجموع کل فروش: {to_locale_str(total_sales)} تومان', 
        size=14, 
        color='red',
        to_row=f'F{summary_row}'
    )
    
    # Add chart
    chart_start_row = summary_row + 2
    add_chart(
        worksheet=worksheet,
        chart_type='bar',
        data_columns=6,  # Total sales column
        category_column=2,  # Product name column
        start_row=5,
        end_row=5 + len(sales_data) - 1,
        chart_position=f"A{chart_start_row}",
        chart_title="نمودار فروش محصولات",
        x_axis_title="محصولات",
        y_axis_title="مبلغ فروش (تومان)",
        chart_width=25,
        chart_height=15
    )
    
    # Save the report
    filename = f"sales_report_{datetime.now().strftime('%Y%m%d')}.xlsx"
    workbook.save(filename)
    print(f"گزارش فروش با موفقیت ایجاد شد: {filename}")
    
    return filename

# Create the report
create_sales_report()
```

### 7. Django Integration

Using excelstyler in Django views:

```python
from django.http import HttpResponse
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO
from excelstyler.headers import create_header_freez
from excelstyler.values import create_value
from excelstyler.utils import shamsi_date
from datetime import datetime

def export_employee_report(request):
    """Export employee data as Excel file"""
    
    # Create workbook in memory
    output = BytesIO()
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.sheet_view.rightToLeft = True
    
    # Add title
    worksheet['A1'] = 'گزارش کارمندان'
    worksheet['A1'].font = Font(size=16, bold=True)
    
    # Create header
    headers = ['ردیف', 'نام', 'نام خانوادگی', 'کد ملی', 'تاریخ استخدام']
    create_header_freez(worksheet, headers, 1, 3, 4, color='green')
    
    # Sample employee data (replace with your actual data)
    employees = [
        ['علی', 'احمدی', '1234567890', datetime(2020, 1, 15)],
        ['فاطمه', 'محمدی', '0987654321', datetime(2021, 3, 20)],
        ['حسن', 'رضایی', '1122334455', datetime(2019, 6, 10)]
    ]
    
    # Add data
    for i, (first_name, last_name, national_id, hire_date) in enumerate(employees, start=4):
        row_data = [
            i-3,
            first_name,
            last_name,
            national_id,
            shamsi_date(hire_date)
        ]
        create_value(worksheet, row_data, i, 1, border_style='thin')
    
    # Save to BytesIO
workbook.save(output)
output.seek(0)

    # Create HTTP response
response = HttpResponse(
        output.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename="employee_report.xlsx"'
    
return response
```

### 8. Error Handling Best Practices

```python
from excelstyler.utils import shamsi_date, to_locale_str
from datetime import datetime, date

def safe_date_conversion(date_input):
    """Safely convert date with error handling"""
    try:
        if isinstance(date_input, str):
            # Try to parse string date
            parsed_date = datetime.strptime(date_input, '%Y-%m-%d')
            return shamsi_date(parsed_date)
        elif isinstance(date_input, (datetime, date)):
            return shamsi_date(date_input)
        else:
            return "تاریخ نامعتبر"
    except (ValueError, TypeError) as e:
        print(f"خطا در تبدیل تاریخ: {e}")
        return "تاریخ نامعتبر"

def safe_number_formatting(number):
    """Safely format number with error handling"""
    try:
        return to_locale_str(number)
    except (ValueError, TypeError) as e:
        print(f"خطا در فرمت عدد: {e}")
        return str(number)

# Usage example
dates = [datetime.now(), "2023-12-25", None, "invalid-date"]
numbers = [1234567, "500000", None, "not-a-number"]

for date_val in dates:
    result = safe_date_conversion(date_val)
    print(f"تاریخ: {date_val} -> {result}")

for num_val in numbers:
    result = safe_number_formatting(num_val)
    print(f"عدد: {num_val} -> {result}")
```



#Example
@api_view(["GET"])
@permission_classes([TokenHasReadWriteScope])
@csrf_exempt
def test_cold_house_excel(request):
    """
    A simplified example Excel report for Cold Houses.
    Excel output support Persian name.
    """

    # --- Excel Setup ---
    output = BytesIO()
    workbook = Workbook()
    worksheet = workbook.active
    workbook.remove(worksheet)
    worksheet = workbook.create_sheet("Cold House Info")
    worksheet.sheet_view.rightToLeft = True
    worksheet.insert_rows(1)

    # --- Header ---
    header = [
        'Row', 'Cold House Name', 'City', 'Address',
        'Total Weight', 'Allocated Weight', 'Remaining Weight',
        'Status', 'Broadcast', 'Relocate', 'Capacity'
    ]
    create_header_freez(worksheet, header, start_col=1, row=2, header_row=3, height=25, width=18)

    # --- Example Data ---
    # Here we use some mock data for testing
    example_data = [
        {
            'name': 'Cold House A',
            'city': 'Tehran',
            'address': 'Street 1',
            'total_input_weight': 1000,
            'total_allocated_weight': 700,
            'total_remain_weight': 300,
            'status': True,
            'broadcast': False,
            'relocate': True,
            'capacity': 1200
        },
        {
            'name': 'Cold House B',
            'city': 'Shiraz',
            'address': 'Street 2',
            'total_input_weight': 800,
            'total_allocated_weight': 500,
            'total_remain_weight': 300,
            'status': False,
            'broadcast': True,
            'relocate': False,
            'capacity': 1000
        }
    ]

    # --- Fill Data ---
    row_index = 3
    for i, house in enumerate(example_data, start=1):
        values = [
            i,
            house['name'],
            house['city'],
            house['address'],
            house['total_input_weight'],
            house['total_allocated_weight'],
            house['total_remain_weight'],
            'Active' if house['status'] else 'Inactive',
            'Yes' if house['broadcast'] else 'No',
            'Yes' if house['relocate'] else 'No',
            house['capacity']
        ]
        create_value(worksheet, values, start_col=row_index, row=1)
        row_index += 1


    # --- Save and Response ---
    workbook.save(output)
    output.seek(0)
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename="ColdHouseExample.xlsx"'
    response.write(output.getvalue())
    return response

```

### 9. Tips and Tricks

#### Color Customization
```python
from excelstyler.styles import PatternFill

# Custom colors
custom_red = PatternFill(start_color="FF6B6B", fill_type="solid")
custom_blue = PatternFill(start_color="4ECDC4", fill_type="solid")

# Use in create_value
create_value(worksheet, data, 1, 1, color=custom_red)
```

#### Working with Large Datasets
```python
def create_large_report(data_list):
    """Create report for large datasets efficiently"""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.sheet_view.rightToLeft = True
    
    # Create header once
    headers = ['ردیف', 'نام', 'مقدار', 'تاریخ']
    create_header_freez(worksheet, headers, 1, 1, 2, color='green')
    
    # Process data in chunks to avoid memory issues
    chunk_size = 1000
    for i in range(0, len(data_list), chunk_size):
        chunk = data_list[i:i + chunk_size]
        for j, row_data in enumerate(chunk, start=i + 2):
            create_value(worksheet, row_data, j, 1, m=j)
    
    return workbook
```

#### Dynamic Column Width
```python
from openpyxl.utils import get_column_letter

def auto_adjust_columns(worksheet, start_col, end_col):
    """Automatically adjust column widths based on content"""
    for col in range(start_col, end_col + 1):
        column_letter = get_column_letter(col)
        max_length = 0
        
        for row in worksheet[column_letter]:
            if row.value:
                max_length = max(max_length, len(str(row.value)))
        
        worksheet.column_dimensions[column_letter].width = max_length + 2
```

### 10. Common Use Cases

#### Financial Reports
```python
def create_financial_report():
    """Create a financial report with currency formatting"""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.sheet_view.rightToLeft = True
    
    # Header
    headers = ['دوره', 'درآمد', 'هزینه', 'سود/زیان']
    create_header(worksheet, headers, 1, 1, color='green')
    
    # Financial data
    financial_data = [
        ['Q1 2023', 1000000000, 800000000, 200000000],
        ['Q2 2023', 1200000000, 900000000, 300000000],
        ['Q3 2023', 1100000000, 850000000, 250000000],
        ['Q4 2023', 1300000000, 950000000, 350000000]
    ]
    
    for i, (period, income, expense, profit) in enumerate(financial_data, start=2):
        row_data = [
            period,
            f"{to_locale_str(income)} تومان",
            f"{to_locale_str(expense)} تومان",
            f"{to_locale_str(profit)} تومان"
        ]
        create_value(worksheet, row_data, i, 1, border_style='thin')
    
    return workbook
```

#### Inventory Management
```python
def create_inventory_report(products):
    """Create inventory report with stock alerts"""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.sheet_view.rightToLeft = True
    
    headers = ['کد محصول', 'نام محصول', 'موجودی', 'حداقل موجودی', 'وضعیت']
    create_header(worksheet, headers, 1, 1, color='blue')
    
    for i, product in enumerate(products, start=2):
        status = "کمبود" if product['stock'] < product['min_stock'] else "کافی"
        
        row_data = [
            product['code'],
            product['name'],
            product['stock'],
            product['min_stock'],
            status
        ]
        
        # Highlight low stock items
        different_cell = 4 if product['stock'] < product['min_stock'] else None
        different_value = product['min_stock'] if product['stock'] < product['min_stock'] else None
        
        create_value(
            worksheet, 
            row_data, 
            i, 
            1, 
            border_style='thin',
            different_cell=different_cell,
            different_value=different_value
        )
    
    return workbook
```

#### Student Grade Report
```python
def create_grade_report(students):
    """Create student grade report with performance indicators"""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.sheet_view.rightToLeft = True
    
    headers = ['نام دانشجو', 'نمره ریاضی', 'نمره فیزیک', 'نمره شیمی', 'میانگین', 'وضعیت']
    create_header(worksheet, headers, 1, 1, color='green')
    
    for i, student in enumerate(students, start=2):
        avg_score = (student['math'] + student['physics'] + student['chemistry']) / 3
        status = "قبول" if avg_score >= 12 else "مردود"
        
        row_data = [
            student['name'],
            student['math'],
            student['physics'],
            student['chemistry'],
            f"{avg_score:.1f}",
            status
        ]
        
        # Highlight failing students
        different_cell = 5 if avg_score < 12 else None
        different_value = avg_score if avg_score < 12 else None
        
        create_value(
            worksheet, 
            row_data, 
            i, 
            1, 
            border_style='thin',
            different_cell=different_cell,
            different_value=different_value
        )
    
    return workbook
```

## 🎨 Available Colors

The library provides these predefined colors:

| Color Name | Hex Code | Usage |
|------------|----------|-------|
| `green` | #00B050 | Success, positive values |
| `red` | #FCDFDC | Errors, negative values |
| `yellow` | #FFFF00 | Warnings, attention |
| `orange` | #FFC000 | Important information |
| `blue` | #538DD5 | Headers, primary info |
| `light_green` | #92D050 | Secondary success |
| `very_light_green` | #5AFC56 | Subtle success |
| `gray` | #B0B0B0 | Disabled, inactive |
| `cream` | #D8AA72 | Default header |
| `light_cream` | #E8C6A0 | Light header |
| `very_light_cream` | #FAF0E7 | Very light background |

## 🔧 Configuration Options

### Border Styles
- `thin` - Thin border
- `medium` - Medium border  
- `thick` - Thick border
- `dashed` - Dashed border
- `dotted` - Dotted border

### Chart Types
- `line` - Line chart
- `bar` - Bar chart

### Text Alignment
All headers and values are automatically center-aligned with text wrapping enabled.

## 🚀 Performance Tips

1. **Use freeze panes** for large datasets to improve navigation
2. **Process data in chunks** for very large datasets
3. **Use alternating colors** sparingly for better performance
4. **Set column widths** explicitly to avoid auto-calculation overhead
5. **Use `in_value=True`** for Persian dates when storing in Excel cells

## 🐛 Troubleshooting

### Common Issues

**Issue**: Persian text not displaying correctly
**Solution**: Always set `worksheet.sheet_view.rightToLeft = True`

**Issue**: Charts not appearing
**Solution**: Ensure data range is correct and data exists in specified cells

**Issue**: Colors not applying
**Solution**: Check color name spelling and ensure it's in the predefined list

**Issue**: Date conversion errors
**Solution**: Use try-catch blocks and validate input dates

### Debug Mode
```python
import logging
logging.basicConfig(level=logging.DEBUG)

# Your excelstyler code here
```

## API Reference

### Headers

#### `create_header(worksheet, data, start_col, row, **kwargs)`
Create a styled header row in an Excel worksheet.

**Parameters:**
- `worksheet`: The Excel worksheet object
- `data`: List of header titles
- `start_col`: Starting column index (1-based)
- `row`: Row index where header will be placed
- `height`: Row height (optional)
- `width`: Column width (optional)
- `color`: Background color ('green', 'red', 'blue', etc.)
- `text_color`: Font color (optional)
- `border_style`: Border style ('thin', 'medium', etc.)

#### `create_header_freez(worksheet, data, start_col, row, header_row, **kwargs)`
Create a header with freeze panes and auto-filter.

### Values

#### `create_value(worksheet, data, start_col, row, **kwargs)`
Write formatted values to Excel cells.

**Parameters:**
- `worksheet`: The Excel worksheet object
- `data`: List of values to write
- `start_col`: Starting row index
- `row`: Starting column index
- `border_style`: Border style (optional)
- `m`: For alternating row colors
- `color`: Cell background color
- `different_cell`: Index of cell to highlight
- `different_value`: Value to highlight

### Utilities

#### `shamsi_date(date, in_value=None)`
Convert Gregorian date to Persian (Shamsi) date.

#### `to_locale_str(number)`
Format number with thousands separators.

### Charts

#### `add_chart(worksheet, chart_type, data_columns, category_column, start_row, end_row, chart_position, chart_title, x_axis_title, y_axis_title, **kwargs)`
Add line or bar charts to Excel worksheets.

## Testing

Run the test suite:

```bash
pip install pytest
pytest
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Run the test suite
6. Submit a pull request

## 📖 Additional Resources

### Video Tutorials
- [Basic Excel Report Creation](https://example.com/basic-tutorial)
- [Advanced Styling Techniques](https://example.com/advanced-tutorial)
- [Persian Date Integration](https://example.com/persian-dates)

### Community Examples
- [GitHub Examples Repository](https://github.com/7nimor/excelstyler-examples)
- [Stack Overflow Tag](https://stackoverflow.com/questions/tagged/excelstyler)

### Related Projects
- [openpyxl](https://openpyxl.readthedocs.io/) - The underlying Excel library
- [jdatetime](https://github.com/slashmili/python-jalali) - Persian date library

## 🤝 Contributing

We welcome contributions! Here's how you can help:

### Reporting Issues
1. Check existing issues first
2. Provide detailed reproduction steps
3. Include Python and excelstyler versions
4. Attach sample code if possible

### Submitting Pull Requests
1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Add tests for your changes
4. Ensure all tests pass (`pytest`)
5. Commit your changes (`git commit -m 'Add amazing feature'`)
6. Push to the branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

### Development Setup
```bash
# Clone the repository
git clone https://github.com/7nimor/excelstyler.git
cd excelstyler

# Install in development mode
pip install -e .

# Install test dependencies
pip install -e ".[test]"

# Run tests
pytest

# Run linting
flake8 src/ tests/
```

## 📊 Changelog

See [CHANGELOG.md](CHANGELOG.md) for detailed version history.

## 🏆 Acknowledgments

- Thanks to the [openpyxl](https://openpyxl.readthedocs.io/) team for the excellent Excel library
- Thanks to the [jdatetime](https://github.com/slashmili/python-jalali) team for Persian date support
- Thanks to all contributors and users who help improve this library

## 📞 Support

- **Documentation**: [Read the docs](https://excelstyler.readthedocs.io/)
- **Issues**: [GitHub Issues](https://github.com/7nimor/excelstyler/issues)
- **Discussions**: [GitHub Discussions](https://github.com/7nimor/excelstyler/discussions)
- **Email**: 7nimor@gmail.com

## ⭐ Star History

[![Star History Chart](https://api.star-history.com/svg?repos=7nimor/excelstyler&type=Date)](https://star-history.com/#7nimor/excelstyler&Date)

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details.

---

<div align="center">

**Made with ❤️ for the Persian/Farsi developer community**

[![GitHub stars](https://img.shields.io/github/stars/7nimor/excelstyler?style=social)](https://github.com/7nimor/excelstyler)
[![GitHub forks](https://img.shields.io/github/forks/7nimor/excelstyler?style=social)](https://github.com/7nimor/excelstyler)
[![GitHub watchers](https://img.shields.io/github/watchers/7nimor/excelstyler?style=social)](https://github.com/7nimor/excelstyler)

</div>
