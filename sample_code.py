"""
sample_code.py
~~~~~~~~~~

sample code.

"client_secret.json" file of your OAuth2 credentials from Google Developers Console
should be in the same directory with this code.

If you don't know how to get your own "client_secret.json" file,
Please visit the following page.

https://developers.google.com/sheets/api/quickstart/python

"""

from gsheetsbatch.client import Client


def sample_read():
    """URL of sample spreadsheet for this sample function is

    https://docs.google.com/spreadsheets/d/1Q5jSop27MzBdhirmFEw_FsNw7stNsRbiZPmjL6cwlc0/
    
    """
    client = Client()
    ss = client.open_by_id(spreadsheet_id='1Q5jSop27MzBdhirmFEw_FsNw7stNsRbiZPmjL6cwlc0')
    sheet = ss.read_cache.get_sheet_by_title('Sheet1')
    print('number of rows of Sheet1:', sheet.read_cache.row_count)
    print('number of columns of Sheet1:', sheet.read_cache.column_count)
    print()

    print('Value of Cell B2:', sheet.read_cache.value_of_cell(2, 2))
    print('Value of Cell B3:', sheet.read_cache.value_of_cell(3, 2))
    print()

    b5_value = sheet.read_cache.value_of_cell(5, 2)
    b5_display_value = sheet.read_cache.display_value_of_cell(5, 2)
    print('Value of Cell B5:', b5_value)
    print('Display Value of Cell B5:', b5_display_value)
    print('Type of value of Cell B5:', type(b5_value))
    print('Type of display value of Cell B5:', type(b5_display_value))
    print()

    # To refresh cache, do refresh_cache().
    # (In this code, refresh_code() is not needed but inserted for showing how to use.)
    sheet.read_cache.refresh_cache()

    value = sheet.read_cache.value_of_cell(6, 2)
    print('Value, Type of Cell B6:', value, type(value))

    effective_value = sheet.read_cache.effective_value_of_cell(6, 2)
    print('Effective Value, Type of Cell B6:', effective_value, type(effective_value))

    display_value = sheet.read_cache.display_value_of_cell(6, 2)
    print('Display Value, Type of Cell B6:', display_value, type(display_value))
    print()

    print(sheet.read_cache.cells_including_text('Hi'))


def sample_write():
    client = Client()
    ss = client.create_spreadsheet(title='This_is_title_of_new_spreadsheet')
    sheet = ss.read_cache.get_sheet_by_index(0)    # get first sheet

    # deposit requests
    sheet.deposit_request.request_update_cell_value(row=2, col=2, values_list_list=[[1, 2, 3], [4, 5, 6], [7, 8, 9]],
                                                    type='numberValue')
    sheet.deposit_request.request_update_borders_around(min_row=2, min_col=2, max_row=4, max_col=4, style='SOLID')

    # execute deposited requests
    client.execute_all_deposited_requests()


if __name__ == '__main__':
    sample_read()
#   sample_write()
