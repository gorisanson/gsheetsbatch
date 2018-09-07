# Google Sheets API batch wrapper
A batch wrapper for Google Sheets API v4 (Python)

Features:

* Uses cached JSONs for reading.
* Uses [batchUpdate](https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/batchUpdate) for writing.


So, you can do more works much faster under the same [quota limit](https://developers.google.com/sheets/api/limits)!

Execution of one [batuchUpdate](https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/batchUpdate)
request counts as one request in [quota](https://developers.google.com/sheets/api/limits),
even though its batch-request body can contain over 1000 requests!

## How to Start

1. [Turn on Google Sheets API.](https://developers.google.com/sheets/api/quickstart/python)

2. [Install Google Client Library](https://developers.google.com/sheets/api/quickstart/python) and Oauth2client Library:

```
pip install --upgrade google-api-python-client
pip install oauth2client
```

3. [Move ***client_secret.json*** file to
    the current working directory.](https://developers.google.com/sheets/api/quickstart/python)

4. Start using this wrapper:

```python
from gsheetsbatch.client import Client

# get client object which is responsible for communicating with Google Sheets API
client = Client()

# get spreadsheet object whose url is
# https://docs.google.com/spreadsheets/d/1Q5jSop27MzBdhirmFEw_FsNw7stNsRbiZPmjL6cwlc0/
spreadsheet = client.open_by_id(spreadsheet_id='1Q5jSop27MzBdhirmFEw_FsNw7stNsRbiZPmjL6cwlc0')

# get sheet object whose title is 'Sheet1'
sheet = spreadsheet.read_cache.get_sheet_by_title('Sheet1')

# read value of cell B3 (row:3, col:2)
value = sheet.read_cache.value_of_cell(row=3, col=2)
```

## How to Read

To read with this wrapper, it is essentail to undertand what **"read_cache"** mean.

Since it uses cached JSONs of spreadsheets and sheets for reading,
sometimes you should ***refresh cache*** manually.

```python
client = Client()
spreadsheet = client.open_by_id(spreadsheet_id='1Q5jSop27MzBdhirmFEw_FsNw7stNsRbiZPmjL6cwlc0')
sheet = spreadsheet.read_cache.get_sheet_by_title('Sheet1')

# read value of cell C2 (row:2, col:3)
# suppose that cell C2 has a string value 'Who?'
value = sheet.read_cache.value_of_cell(row=2, col=3)
print(value) # prints "Who?"

# some codes deposit and execute writing to cell C2
sheet.deposit_request.update_cells_values(row=2, col=3, values_list_list=[['Me!']],
                                          type='stringValue')
client.execute_all_deposited_requests()

# If you want to read the newly written value of cell C2
# you should refresh cache.
sheet.read_cache.refresh_cache()
new_value = sheet.read_cache.value_of_cell(row=2, col=3)
print(new_value) # prints "Me!"
```

## How to Write

To write with this wrapper, it is essential to understand what **"deposit_request"** mean.

Since it uses batchUpdate for writing, it ***deposites*** writing requests.

It ***executes*** deposited requests ***only when ordered explicitly***.

```python
client = Client()

# create new spreadsheet whose title is 'This_is_title_of_new_spreadsheet'
spreadsheet = client.create_spreadsheet(title='This_is_title_of_new_spreadsheet')

# get first sheet (zero-based index)
sheet = spreadsheet.read_cache.get_sheet_by_index(index=0)

# prepare values to write
values_list_list = [
    [1, 2, 3, 4],
    [5, 6, 7, 8],
    [9, 10, 11, 12]
]

# deposit a request to write the 2D list above
# to the 3x4 range "B2:D4" whose left-top cell is B2 (row:2, col:2)
sheet.deposit_request.update_cells_values(row=2, col=2, values_list_list=values_list_list,
                                          type='numberValue')

# deposit a request to draw solid borders around the 3x4 range "B2:D4".
sheet.deposit_request.update_borders_around(min_row=2, min_col=2, max_row=4, max_col=5,
                                                    style='SOLID')

# excecute all deposited requests!
client.execute_all_deposited_requests()
```

## More Examples

Please refer to ***[sample_code.py](https://github.com/gorisanson/gsheetsbatch/blob/master/sample_code.py)*** for more examples.

## License and copyright notice

Some portions of the code and comments and outlines taken from
**[gspread (MIT License Copyright (C) 2011-2018 Anton Burnashev)](https://github.com/burnash/gspread)**. 
