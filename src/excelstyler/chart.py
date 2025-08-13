from openpyxl.chart import LineChart, Reference, BarChart


def add_chart(
        worksheet,
        chart_type,
        data_columns,
        category_column,
        start_row,
        end_row,
        chart_position,
        chart_title,
        x_axis_title,
        y_axis_title,
        chart_width=25,  # عرض نمودار پیش‌فرض (واحد: cm)
        chart_height=15
):
    """
    افزودن نمودار به صفحه اکسل.

    ورودی:
        worksheet (openpyxl.Worksheet): صفحه اکسل.
        chart_type (str): نوع نمودار ("line" یا "bar").
        data_columns (list): لیستی از ستون‌های داده.
        category_column (int): ستون دسته‌بندی‌ها.
        start_row (int): ردیف شروع داده‌ها.
        end_row (int): ردیف پایان داده‌ها.
        chart_position (str): محل قرار گرفتن نمودار.
        chart_title (str): عنوان نمودار.
        x_axis_title (str): عنوان محور X.
        y_axis_title (str): عنوان محور Y.
        chart_width (float): عرض نمودار (واحد: cm).
        chart_height (float): ارتفاع نمودار (واحد: cm).
    """

    if chart_type == 'line':
        chart = LineChart()
        chart.style = 20
    elif chart_type == 'bar':
        chart = BarChart()
    else:
        raise ValueError("chart_type باید 'line' یا 'bar' باشد.")

    chart.title = chart_title
    chart.y_axis.title = y_axis_title
    chart.x_axis.title = x_axis_title
    chart.width = chart_width
    chart.height = chart_height

    categories = Reference(worksheet, min_col=category_column, min_row=start_row, max_row=end_row)
    data = Reference(worksheet, min_col=data_columns, min_row=start_row - 1, max_row=end_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    for series in chart.series:
        series.graphicalProperties.line.solidFill = "277358"
        series.graphicalProperties.line.width = 30000

    worksheet.add_chart(chart, chart_position)
    # example
    # add_chart(
    #     worksheet=worksheet,
    #     chart_type='line',
    #     data_columns=7,  # ستون وزن وارد شده
    #     category_column=2,  # ستون نام سردخانه‌ها
    #     start_row=7,
    #     end_row=l + 1,
    #     chart_position="A12",
    #     chart_title="نمودار تغییرات وزن در سردخانه‌ها",
    #     x_axis_title="سردخانه‌ها",
    #     y_axis_title="وزن (کیلوگرم)"
    # )
