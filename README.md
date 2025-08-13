# excelstyler

`excelstyler` is a Python package that makes it easy to style and format Excel worksheets using [openpyxl](https://openpyxl.readthedocs.io).

## Installation
```bash
pip install excelstyler
```

## Example
```python
from excelstyler.styles import GREEN_CELL
from excelstyler.utils import shamsi_date
from excelstyler.headers import create_header

import openpyxl

output = BytesIO()
workbook = Workbook()
worksheet = workbook.active
worksheet.sheet_view.rightToLeft = True   # if you iranian else False
worksheet.insert_rows(1)
create_header(ws, ["Name", "Score"], 1, 1, color='green')
workbook.save(output)
output.seek(0)

response = HttpResponse(
  content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
response[
  'Content-Disposition'] = f'attachment; filename="test file.xlsx"'.encode(
  'utf-8') # support from persian
response.write(output.getvalue())
return response
```

## License
MIT
