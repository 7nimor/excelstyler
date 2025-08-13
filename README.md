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

wb = openpyxl.Workbook()
ws = wb.active
create_header(ws, ["Name", "Score"], 1, 1, color='green')
wb.save("example.xlsx")
```

## License
MIT
