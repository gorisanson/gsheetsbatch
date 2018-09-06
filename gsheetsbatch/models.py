"""
models
~~~~~~~~~~

This module contains Spreadsheet, Sheet class.

"""


def grid_range(sheet_id, min_row, min_col, max_row, max_col):
    """Returns GridRange json.

    min_row, min_col, max_row, max_col: int (None if unbound)

    :returns: GridRange json
    """
    if min_row is None:
        start_row_index = None
    else:
        start_row_index = min_row - 1

    if min_col is None:
        start_column_index = None
    else:
        start_column_index = min_col - 1

    grid_range_json = {
        'sheetId': sheet_id,
        'startRowIndex': start_row_index,
        'endRowIndex': max_row,
        'startColumnIndex': start_column_index,
        'endColumnIndex': max_col
    }
    return grid_range_json


class Spreadsheet:
    """The class that represents a spreadsheet.

    :param json: JSON representation that represents a spreadsheet
    """
    def __init__(self, client, json):
        self.client = client
        self.spreadsheet_id = json['spreadsheetId']
        self.json = json
        self.read_cache = Spreadsheet.JSONCacheReader(self)
        self.deposit_request = Spreadsheet.RequestDepositor(self)

    class JSONCacheReader:
        def __init__(self, spreadsheet):
            self._spreadsheet = spreadsheet

        @property
        def title(self):
            """Spreadsheet title (str)"""
            return self._spreadsheet.json['properties']['title']

        def refresh_cache(self):
            request = self._spreadsheet.client.service.spreadsheets().\
                get(spreadsheetId=self._spreadsheet.spreadsheet_id, includeGridData=True)
            self._spreadsheet.json = request.execute()

        def get_sheets(self):
            """Returns a list of Sheet objects of this Spreadsheet object.

            :returns: list of Sheet objects
            """
            sheets_json = self._spreadsheet.json['sheets']
            return [Sheet(self._spreadsheet, sheet_json) for sheet_json in sheets_json]

        def get_sheet_by_id(self, sheet_id):
            """Returns a Sheet object whose sheet id is sheet_id.

            :param sheet_id: int
            :returns: Sheet object
            """
            sheets_json = self._spreadsheet.json['sheets']
            for sheet_json in sheets_json:
                if sheet_json['properties']['sheetId'] == sheet_id:
                    return Sheet(self._spreadsheet, sheet_json)
            raise AssertionError("Couldn't find corresponding cached sheet json"
                                 "whose sheetId is " + sheet_id + ". "
                                 "Please consider to do refresh_cache before")

        def get_sheet_by_index(self, index):
            """Returns a Sheet object whose index is index.

            :param index: int
            :returns: Sheet object
            """
            sheets_json = self._spreadsheet.json['sheets']
            return Sheet(self._spreadsheet, sheets_json[index])

        def get_sheet_by_title(self, title):
            """Returns a Sheet object whose title is title.

            :param title: str
            :returns: Sheet object
            """
            sheets_json = self._spreadsheet.json['sheets']
            for sheet_json in sheets_json:
                if sheet_json['properties']['title'] == title:
                    return Sheet(self._spreadsheet, sheet_json)
            raise AssertionError("Couldn't find corresponding cached sheet json"
                                 "whose sheet title is" + title + ". "
                                 "Please consider to do refresh_cache before")

    class RequestDepositor:
        def __init__(self, spreadsheet):
            self._spreadsheet = spreadsheet
            self.sheet_ids_requested_to_add = []

        def update_spreadsheet_title(self, title):
            """Deposit a request to update title of spreadsheet to the title

            :param title: str
            """
            request = {
                'updateSpreadsheetProperties': {
                    'properties': {
                        'title': title
                    },
                    'fields': 'title'
                }
            }
            self._spreadsheet.client.requests_container.deposit(self._spreadsheet.spreadsheet_id, request)

        def add_new_sheet(self):
            """Deposit a request to add a new sheet.
            the created new sheet has sheet id which is the least natural number (including 0) not in current sheet ids.

            :returns new Sheet object requested to add
            (this object is only for deposit_request, not for read_cache,
            until deposited requests executed and cache is refreshed.)
            """
            new_sheet_id = self.sheet_ids_requested_to_add[-1] + 1 if self.sheet_ids_requested_to_add else 0
            sheets_json = self._spreadsheet.json['sheets']
            sheet_ids = [sheet_json['properties']['sheetId'] for sheet_json in sheets_json]
            while new_sheet_id in sheet_ids:  # sheet_ids 에 없는 제일 작은 0 이상의 정수로 new_sheet_id 를 고른다.
                new_sheet_id += 1

            request = {
                'addSheet': {
                    'properties': {
                        'sheetId': new_sheet_id
                    }
                }
            }
            self._spreadsheet.client.requests_container.deposit(self._spreadsheet.spreadsheet_id, request)
            self.sheet_ids_requested_to_add.append(new_sheet_id)
            new_sheet_json = {
                'properties': {
                    'sheetId': new_sheet_id
                }
            }
            return Sheet(self._spreadsheet, new_sheet_json)


class Sheet:
    """The class that represents a single sheet in a spreadsheet.

    :param spreadsheet: Spreadsheet Object including this sheet.
    :param json: JSON representation that represents a single sheet
    """
    def __init__(self, spreadsheet, json):
        self.client = spreadsheet.client
        self.parent_spreadsheet = spreadsheet
        self.sheet_id = json['properties']['sheetId']
        self.json = json
        self.read_cache = Sheet.JSONCacheReader(self)
        self.deposit_request = Sheet.RequestDepositor(self)

    class JSONCacheReader:
        def __init__(self, sheet):
            self._sheet = sheet

        def refresh_cache(self):
            get_spreadsheet_by_data_filter_request_body = {
                'dataFilters': [
                    {
                        'gridRange': grid_range(self._sheet.sheet_id, None, None, None, None)
                    }
                ],
                'includeGridData': True
            }
            request = self._sheet.client.service.spreadsheets().\
                getByDataFilter(spreadsheetId=self._sheet.parent_spreadsheet.spreadsheet_id,
                                body=get_spreadsheet_by_data_filter_request_body)
            spreadsheet_json = request.execute()
            self._sheet.json = spreadsheet_json["sheets"][0]

        def value_of_cell(self, row, col):
            """Returns user entered value of the cell (row, col).
            For cell with formulas, returned type is str.
            """
            try:
                d = self._sheet.json['data'][0]['rowData'][row-1]['values'][col-1]['userEnteredValue']
            except (KeyError, IndexError):
                d = {}

            value = None
            for key in d:
                value = d[key]
            return value

        def effective_value_of_cell(self, row, col):
            """Returns effective value of the cell (row, col).
            For cells with formulas, effective value is the calculated value.
            """
            try:
                d = self._sheet.json['data'][0]['rowData'][row-1]['values'][col-1]['effectiveValue']
            except (KeyError, IndexError):
                d = {}

            value = None
            for key in d:
                value = d[key]
            return value

        def display_value_of_cell(self, row, col):
            """Returns display value of the cell (row, col).
            Returned type is always str.

            :returns str
            """
            try:
                display_value = self._sheet.json['data'][0]['rowData'][row - 1]['values'][col - 1]['formattedValue']
            except (KeyError, IndexError):
                display_value = ''

            return display_value

        @property
        def title(self):
            """Title of this sheet (str)."""
            return self._sheet.json['properties']['title']

        @property
        def index(self):
            """Zero-based index of this sheet (int)."""
            return self._sheet.json['properties']['index']

        @property
        def row_count(self):
            """Number of rows (int)."""
            return self._sheet.json['properties']['gridProperties']['rowCount']

        @property
        def column_count(self):
            """Number of columns (int)."""
            return self._sheet.json['properties']['gridProperties']['columnCount']

        @property
        def is_hidden(self):
            """Is hidden sheet? (boolean)"""
            return self._sheet.json['properties']['hidden']

        @property
        def merged_ranges(self):
            """Returns list of 4-tuples that represents merged ranges.
            If there's no merged ranges, return empty list.

            returns: list of 4-tuple of int (min_row, min_col, max_row, max_col)
            """
            merges = []
            sheet = self._sheet.json
            for merge in sheet.get('merges'):
                merges.append((merge['startRowIndex'] + 1, merge['startColumnIndex'] + 1,
                               merge['endRowIndex'], merge['endColumnIndex']))
            return merges

        def merged_range_of_cell(self, row, col):
            """Returns 4-tuple that represents merged range of the cell.
            If the cell is not merged, returns the range of cell (row, column, row, column).

            row: int
            column: int
            returns: 4-tuple of int (min_row, min_col, max_row, max_col).
            """
            merge = (row, col, row, col)
            merges = self.merged_ranges
            for _merge in merges:
                min_row, min_col, max_row, max_col = _merge
                if min_row <= row <= max_row and min_col <= col <= max_col:
                    merge = _merge
                    break
            return merge

        @property
        def conditional_formats(self):
            """Returns list of ConditionalFormatRule objects.

            returns: ConditionalFormatRule json
            """
            sheet = self._sheet.json
            conditional_formats = sheet['conditionalFormats']
            return conditional_formats

        @property
        def protected_ranges(self):
            """Returns list of ProtectedRange objects.

            returns: ProtectedRange json
            """
            sheet = self._sheet.json
            protected_ranges = sheet['protectedRanges']
            return protected_ranges

        def data_validation_rule_of_cell(self, row, col):
            """Returns data validation rule of the cell.
            If there's no data validation rule, returns None

            row: int
            col: int

            returns: data validation rule json or None
            """
            sheet = self._sheet.json
            data_validation_rule = sheet['data'][0]['rowData'][row - 1]['values'][col - 1].get('dataValidation')
            return data_validation_rule

        def cells_including_text(self, s):
            """Returns list of tuples of positions of cells contain str s.

            s: str
            ws: Worksheet object

            returns: list
            """
            t = []
            for row in range(1, self.row_count + 1):
                for col in range(1, self.column_count + 1):
                    if s in self.display_value_of_cell(row, col):
                        t.append((row, col))
            return t

        def border_style_of_cell(self, row, col, side):
            """Returns style of border on the side.

            :param side: str (one of 'top', 'bottom', 'left', 'right')

            :return: str (one of None, 'DOTTED', 'DASHED', 'SOLID', 'SOLID_MEDIUM', 'SOLID_THICK', 'DOUBLE')
            """
            if side not in ('top', 'bottom', 'left', 'right'):
                raise ValueError("side should be one of 'top', 'bottom', 'left', 'right', but '{}'.".format(side))
            sheet = self._sheet.json
            try:
                border_style = sheet['data'][0]['rowData'][row - 1]['values'][col - 1] \
                ['effectiveFormat']['borders'][side]['style']
            except KeyError:
                border_style = None
            return border_style

    class RequestDepositor:
        def __init__(self, sheet):
            self._sheet = sheet

        def insert_row_or_column(self, dimension, start_index, end_index, inherit_from_before=True):
            """Deposit a request to insert rows or columns.

            dimension: str ('COLUMNS' or 'ROWS')
            start_index: int
            end_index: int
            inherit_from_before: boolean ('true' if inherit from before, 'false' if inherit from after)

            :returns: None
            """
            if dimension != 'COLUMNS' and dimension != 'ROWS':
                raise ValueError("dimension must be 'COLUMNS'or 'ROWS'.")
            request = {
                'insertDimension': {
                    'range': {
                        'sheetId': self._sheet.sheet_id,
                        'dimension': dimension,
                        'startIndex': start_index - 1,
                        'endIndex': end_index
                    },
                    'inheritFromBefore': inherit_from_before
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def size_row_or_column(self, dimension, start_index, end_index, pixel_size):
            """Deposit a request to fix a column width or row height.

            dimension: str ('COLUMNS' or 'ROWS')
            start_index: int (None if unbound)
            end_index: int (None if unbound)
            pixelSize: int

            :returns: None
            """
            if dimension != 'COLUMNS' and dimension != 'ROWS':
                raise ValueError("dimension must be 'COLUMNS'or 'ROWS'.")
            request = {
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': self._sheet.sheet_id,
                        'dimension': dimension,
                        # google api는 zero based index를 사용하고, min쪽은 inclusive,
                        # max 쪽은 excluive를 사용하기에 아래와 같이 조정해서 대입해야한다.
                        'startIndex': start_index - 1,
                        'endIndex': end_index
                    },
                    'properties': {
                        'pixelSize': pixel_size
                    },
                    'fields': 'pixelSize'
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def hide_grid_lines(self, b):
            """Deposit a request to fix the hideGridlines option.

            sheet_id: int
            b: boolean
            returns; None
            """
            request = {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': self._sheet.sheet_id,
                        'gridProperties': {
                            'hideGridlines': b
                        }
                    },
                    'fields': 'gridProperties.hideGridlines'
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def update_borders(self, min_row, min_col, max_row, max_col, side, style='SOLID', color=None):
            """Deposit a request to update borders.

            sheet_id: int
            min_row: int
            min_col: int
            max_row: int
            max_col: int
            side: str (top, bottom, left, right, innerHorizontal, innerVertical)
            style: str (DOTTED, DASHED, SOLID, SOLID_MEDIUM, SOLID_THICK, NONE, DOUBLE)
            color: color object (RGBA color)

            returns; None
            """
            if side not in ('top', 'bottom', 'left', 'right', 'innerHorizontal', 'innerVertical'):
                raise ValueError("side must be a str which is one of 'top', 'bottom',"
                                 "'left', 'right', 'innerHorizontal', 'innerVertical'")
            if style not in ('DOTTED', 'DASHED', 'SOLID', 'SOLID_MEDIUM', 'SOLID_THICK', 'NONE', 'DOUBLE'):
                raise ValueError("style must be a str wich is one of 'DOTTED', 'DASHED', 'SOLID',"
                                 "'SOLID_MEDIUM', 'SOLID_THICK', 'NONE', 'DOUBLE'")

            border_object = {'style': style}
            if color is not None:
                border_object['color'] = color

            request = {
                'updateBorders': {
                    'range': grid_range(self._sheet.sheet_id, min_row, min_col, max_row, max_col),
                    side: border_object
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def update_borders_around(self, min_row, min_col, max_row, max_col, style='SOLID', color=None):
            """Deposit requests_container to update borders around.

            sheet_id: int
            min_row: int
            min_col: int
            max_row: int
            max_col: int
            style: str (DOTTED, DASHED, SOLID, SOLID_MEDIUM, SOLID_THICK, NONE, DOUBLE)
            color: color object (RGBA color)

            returns; None
            """
            self.update_borders(min_row, min_col, max_row, max_col, 'top', style, color)
            self.update_borders(min_row, min_col, max_row, max_col, 'right', style, color)
            self.update_borders(min_row, min_col, max_row, max_col, 'bottom', style, color)
            self.update_borders(min_row, min_col, max_row, max_col, 'left', style, color)

        def update_cells_default_format(self, min_row, min_col, max_row, max_col,
                                                horizontal_alignment='LEFT', vertical_alignment='MIDDLE',
                                                font_family='Malgun Gothic', font_size=15):
            """Deposit requests to update default formats of cells in the range.

            min_row, min_col, max_row, max_col: int
            horizontal_alignment: str (one of 'LEFT', 'CENTER', 'RIGHT')
            vertical_alignment: str (one of 'TOP', 'MIDDLE', 'BOTTOM')
            font_family: str
            font_size: int

            returns: None
            """
            request = {
                'repeatCell': {
                    'range': grid_range(self._sheet.sheet_id, min_row, min_col, max_row, max_col),
                    'cell': {
                        'userEnteredFormat': {
                            'horizontalAlignment': horizontal_alignment,
                            'verticalAlignment': vertical_alignment,
                            'textFormat': {
                                'fontFamily': font_family,
                                'fontSize': font_size
                            }
                        }
                    },
                    'fields': 'userEnteredFormat.horizontalAlignment, userEnteredFormat.verticalAlignment,'
                              'userEnteredFormat.textFormat.fontFamily, userEnteredFormat.textFormat.fontSize'
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def update_cells_text_format(self, min_row, min_col, max_row, max_col, **text_format):
            """Deposit requests to update text formats of cells in the range.

            min_row, min_col, max_row, max_col: int
            text_format: key: value as follows. (optional)
                foregroundColor: { object(Color)}, fontFamily: string, fontSize: number, bold: boolean,
                italic: boolean, strikethrough: boolean, underline: boolean,
            returns: None
            """
            if not text_format:
                raise ValueError("text_format should be specified")

            fields = ''
            first_key_flag = True
            for key in text_format:
                if first_key_flag:
                    fields += 'userEnteredFormat.textFormat.' + key
                    first_key_flag = False
                fields += ', userEnteredFormat.textFormat.' + key

            request = {
                'repeatCell': {
                    'range': grid_range(self._sheet.sheet_id, min_row, min_col, max_row, max_col),
                    'cell': {
                        'userEnteredFormat': {
                            'textFormat': text_format
                        }
                    },
                    'fields': fields
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def update_cell_note(self, row, col, text):
            """Deposit request to update note of the cell.

            row, column: int
            text: str
            returns: None
            """
            request = {
                "repeatCell": {
                    "range": grid_range(self._sheet.sheet_id, row, col, row, col),
                    "cell": {
                        "note": text
                    },
                    "fields": "note"
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def update_cells_alignment(self, min_row, min_col, max_row, max_col,
                                           horizontal_alignment, vertical_alignment='MIDDLE'):
            """Deposit requests to update default formats of cells in the range.

            sheet_id:  int
            min_row, min_col, max_row, max_col: int
            horizontal_alignment: str (one of 'LEFT', 'CENTER', 'RIGHT')
            vertical_alignment: str (one of 'TOP', 'MIDDLE'(default), 'BOTTOM')

            returns: None
            """
            if horizontal_alignment not in ('LEFT', 'CENTER', 'RIGHT'):
                raise ValueError("horizontal_alignment must be a str which is one of 'LEFT', 'CENTER', 'RIGHT'")
            if vertical_alignment not in ('TOP', 'MIDDLE', 'BOTTOM'):
                raise ValueError("vertical_alignment should be a str which is one of 'TOP', 'MIDDLE', 'BOTTOM'")

            request = {
                'repeatCell': {
                    'range': grid_range(self._sheet.sheet_id, min_row, min_col, max_row, max_col),
                    'cell': {
                        'userEnteredFormat': {
                            'horizontalAlignment': horizontal_alignment,
                            'verticalAlignment': vertical_alignment,
                        }
                    },
                    'fields': 'userEnteredFormat.horizontalAlignment, userEnteredFormat.verticalAlignment,'
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def update_cells_background_color(self, min_row, min_col, max_row, max_col, color):
            """Deposit requests to update cells background color.


            sheet_id: int
            min_row, min_col, max_row, max_col: int
            color: color object (in RGBA color space)

            returns:
            """
            request = {
                'repeatCell': {
                    'range': grid_range(self._sheet.sheet_id, min_row, min_col, max_row, max_col),
                    'cell': {
                        'userEnteredFormat': {
                            'backgroundColor': color
                        }
                    },
                    'fields': 'userEnteredFormat.backgroundColor'
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def update_cells_foreground_color(self, min_row, min_col, max_row, max_col, color):
            """Deposit requests_container to update cells foreground color.
            foreground color 란 글자 색을 말한다.


            sheet_id: int
            min_row, min_col, max_row, max_col: int
            color: color object (RGBAcolor)

            returns:
            """
            request = {
                'repeatCell': {
                    'range': grid_range(self._sheet.sheet_id, min_row, min_col, max_row, max_col),
                    'cell': {
                        'userEnteredFormat': {
                            'textFormat': {
                                'foregroundColor': color
                            }
                        }
                    },
                    'fields': 'userEnteredFormat.textFormat.foregroundColor'
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def update_cells_values(self, row, col, values_list_list, type, **text_format):
            """Deposit requests to update cells which starts with (row, col).
            values_list_list must be a list of lists.
            values_list_list[i][j] corresponds to value of (row+i-1, col+j-1) cell.

            :param row: int
            :param col: int
            :param values_list_list: list of lists
            :param type: str ('numberValue', 'stringValue', 'boolValue', 'formulaValue', 'errorValue')
            :param text_format: key: value as follows. (optional)
                                foregroundColor: {object(Color)}, fontFamily: string, fontSize: number, bold: boolean,
                                italic: boolean, strikethrough: boolean, underline: boolean

            :returns: None
            """
            if type not in ('numberValue', 'stringValue', 'boolValue', 'formulaValue', 'errorValue'):
                raise ValueError("type must be str which is one of: 'numberValue', 'stringValue',"
                                 "'boolValue', 'formulaValue', 'errorValue'.")

            for key in text_format:
                if key not in (
                'foregroundColor', 'fontFamily', 'fontSize', 'bold', 'italic', 'strikethrough', 'underline'):
                    raise ValueError('text_format is not in a valid format.')

            fields = 'userEnteredValue'
            for key in text_format:
                fields += ', userEnteredFormat.textFormat.' + key

            rows_data = []
            for row_values in values_list_list:
                _row_data = []
                for cell_value in row_values:
                    cell_data = {
                        'userEnteredValue': {
                            type: cell_value
                        },
                        'userEnteredFormat': {
                            'textFormat': text_format
                        }
                    }
                    _row_data.append(cell_data)
                rows_data.append({'values': _row_data})

            request = {
                'updateCells': {
                    'rows': rows_data,
                    'fields': fields,
                    'start': {
                        'sheetId': self._sheet.sheet_id,
                        'rowIndex': row - 1,
                        'columnIndex': col - 1
                    }
                }
            }

            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def update_cells_text_format_runs(self, min_row, min_col, max_row, max_col, start_index,
                                                  end_index=None,
                                                  **text_format):
            """Deposit requests to update cells' text format runs from start_index to end_index.
            If end_index is not specified or None, this run continues to the end.

            min_row, min_col, max_row, max_col: int
            start_index, end_index: int (start_index is inclusive, end_index is exclusive)
            text_format: key: value as follows. (absent values inherit the cell's format)
                        foregroundColor: {object(Color)}, fontFamily: string, fontSize: number, bold: boolean,
                        italic: boolean, strikethrough: boolean, underline: boolean,

            returns: None
            """
            for key in text_format:
                if key not in (
                'foregroundColor', 'fontFamily', 'fontSize', 'bold', 'italic', 'strikethrough', 'underline'):
                    raise ValueError('text_format is not in a valid format.')

            text_format_runs = [
                {'format': text_format,
                 'startIndex': start_index}
            ]

            if end_index is not None:
                if start_index > end_index:
                    raise IndexError("start_index is bigger then end_index."
                                     "start_index must be smaller or equal to end_index. ")
                text_format_runs.append(
                    {'format': {},
                     'startIndex': end_index}
                )

            request = {
                'repeatCell': {
                    'range': grid_range(self._sheet.sheet_id, min_row, min_col, max_row, max_col),
                    'cell': {
                        'textFormatRuns': text_format_runs
                    },
                    'fields': 'textFormatRuns'
                }
            }

            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def merge_cells(self, min_row, min_col, max_row, max_col, merge_type='MERGE_ALL'):
            """Deposit a request to merge cells.

            sheet_id:
            min_row:
            min_col:
            max_row:
            max_col:
            merge_type: it is optional. str. (one of: 'MERGE_ALL', 'MERGE_COLUMNS', 'MERGE_ROWS')

            returns: None
            """
            if merge_type not in ('MERGE_ALL', 'MERGE_COLUMNS', 'MERGE_ROWS'):
                raise ValueError("merge_type(optional) must be a str which is one of"
                                 "'MERGE_ALL', 'MERGE_COLUMNS', 'MERGE_ROWS'")
            request = {
                'mergeCells': {
                    'range': grid_range(self._sheet.sheet_id, min_row, min_col, max_row, max_col),
                    'mergeType': merge_type
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def unmerge_cells(self, min_row, min_col, max_row, max_col):
            """Deposit a request to unmerge cells.

            sheet_id:
            min_row:
            min_col:
            max_row:
            max_col:

            returns: None
            """
            request = {
                'unmergeCells': {
                    'range': grid_range(self._sheet.sheet_id, min_row, min_col, max_row, max_col)
                },
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def update_cells_data_validation_rule(self, min_row, min_col, max_row, max_col, rule):
            """Deposit a request to update cells data validation rule.

            sheet_id: int
            min_row, min_col, max_row, max_col: int
            rule: DataValidationRule object

            returns: None
            """
            request = {
                'repeatCell': {
                    'range': grid_range(self._sheet.sheet_id, min_row, min_col, max_row, max_col),
                    'cell': {
                        'dataValidation': rule
                    },
                    'fields': 'dataValidation'
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def update_sheet_title(self, title):
            """Deposit a request to update sheet title.

            title: string
            returns: None
            """
            request = {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': self._sheet.sheet_id,
                        'title': title
                    },
                    'fields': 'title'
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def add_conditional_format_rule(self, min_row, min_col, max_row, max_col, type, rule, index=0):
            """Deposit a request to add conditional format rule

            min_row, min_col, max_row, max_col: int
            type: str (one of 'booleanRule', 'gradientRule')
            rule: booleanRule Object or gradientRule Object corresponding to type
            index: int (the zero-based index where the rule should be inserted) 이 룰들에 번호가 매겨지나보다.
            returns: None
            """
            if type not in ('booleanRule', 'gradientRule'):
                raise ValueError("type should be a str which is one of 'booleanRule', 'gradientRule'.")
            request = {
                'addConditionalFormatRule': {
                    'rule': {
                        'ranges': [
                            grid_range(self._sheet.sheet_id, min_row, min_col, max_row, max_col)
                        ],
                        type: rule
                    },
                    'index': index
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def delete_conditional_format_rule(self, index):
            """Deposit a request to delete conditional format rule indicated by zero-based index.
            All subsequent rules' indexes are decremented.

            index: int
            returns: None
            """
            request = {
                'deleteConditionalFormatRule': {
                    'index': index,
                    'sheetId': self._sheet.sheet_id
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def update_cells_number_format(self, min_row, min_col, max_row, max_col, type, pattern=None):
            """Deposit a request to add conditional format rule

            min_row, min_col, max_row, max_col: int
            type: str (one of AUTOMATIC, TEXT, NUMBER, PERCENT, CURRENCY, DATE, TIME, DATE_TIME, SCIENTIFIC)
            pattern: str (Pattern string)
            returns: None
            """
            if type not in ('AUTOMATIC', 'TEXT', 'NUMBER', 'PERCENT', 'CURRENCY', 'DATE', 'TIME', 'DATE_TIME',
                            'SCIENTIFIC'):
                raise ValueError("type must be one of TEXT, NUMBER, PERCENT, CURRENCY, DATE, TIME, DATE_TIME,"
                                 "SCIENTIFIC. But type={}".format(type))
            if type == 'AUTOMATIC':
                request = {
                    'repeatCell': {
                        'range': grid_range(self._sheet.sheet_id, min_row, min_col, max_row, max_col),
                        'cell': {
                            'userEnteredFormat': {
                                'numberFormat': {}
                            }

                        },
                        'fields': 'userEnteredFormat.numberFormat'
                    }
                }
            else:
                request = {
                    'repeatCell': {
                        'range': grid_range(self._sheet.sheet_id, min_row, min_col, max_row, max_col),
                        'cell': {
                            'userEnteredFormat': {
                                'numberFormat': {
                                    'type': type,
                                    'pattern': pattern
                                }
                            }

                        },
                        'fields': 'userEnteredFormat.numberFormat'
                    }
                }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def add_sheet_protection(self):
            assert self._sheet.client.google_account_email
            request = {
                'addProtectedRange': {
                    'protectedRange': {
                        'range': {
                            'sheetId': self._sheet.sheet_id,
                            'startRowIndex': None,
                            'endRowIndex': None,
                            'startColumnIndex': None,
                            'endColumnIndex': None
                        },
                        # 이 editors 객체가 꼭 있어야한다.
                        'editors': {
                            'users': [self._sheet.client.google_account_email],
                        }
                    }
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def add_sheet_protection_except_unprotected_range(self, min_row, min_col, max_row, max_col):
            """Deposit a request to add whole sheet protection except the specified range.

            min_row, min_col, max_row, max_col: int
            returns: None
            """
            assert self._sheet.client.google_account_email
            request = {
                'addProtectedRange': {
                    'protectedRange': {
                        'range': {
                            'sheetId': self._sheet.sheet_id,
                            'startRowIndex': None,
                            'endRowIndex': None,
                            'startColumnIndex': None,
                            'endColumnIndex': None
                        },
                        'unprotectedRanges': [
                            grid_range(self._sheet.sheet_id, min_row, min_col, max_row, max_col)
                        ],
                        # 이 editors 객체가 꼭 있어야한다.
                        'editors': {
                            'users': [self._sheet.client.google_account_email],
                        }
                    }
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def add_sheet_protection_except_unprotected_ranges(self, unprotected_ranges):
            """Deposit a request to add whole sheet protection except the specified ranges.

            unprotected_ranges: list of dicts {'min_row': min_row, 'min_col': min_col,
                                               'max_row': max_row, 'max_col': max_col}

            returns: None
            """
            assert self._sheet.client.google_account_email
            _unprotected_ranges = [grid_range(self._sheet.sheet_id, **x) for x in unprotected_ranges]
            request = {
                'addProtectedRange': {
                    'protectedRange': {
                        'range': {
                            'sheetId': self._sheet.sheet_id,
                            'startRowIndex': None,
                            'endRowIndex': None,
                            'startColumnIndex': None,
                            'endColumnIndex': None
                        },
                        'unprotectedRanges': _unprotected_ranges,
                        # 이 editors 객체가 꼭 있어야한다.
                        'editors': {
                            'users': [self._sheet.client.google_account_email],
                        }
                    }
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def delete_sheet_protection(self, protected_range_id):
            """Deposit a request to delete sheet protection whose id is protected_range_id.

            protected_range_id: int
            returns: None
            """
            request = {
                'deleteProtectedRange': {
                    'protectedRangeId': protected_range_id
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def update_cells_wrap_strategy(self, min_row, min_col, max_row, max_col, wrap_strategy):
            """
            Request to update cells WrapStrategy.

            min_row, min_col, max_row, max_col: int
            wrap_strategy: str (one of OVERFLOW_CELL, LEGACY_WRAP, CLIP, WRAP)

            returns: None
            """
            request = {
                'repeatCell': {
                    'range': grid_range(self._sheet.sheet_id, min_row, min_col, max_row, max_col),
                    'cell': {
                        'userEnteredFormat': {
                            'wrapStrategy': wrap_strategy
                        }
                    },
                    'fields': 'userEnteredFormat.wrapStrategy'
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def delete_dimension(self, dimension, min_index, max_index):
            """Request to delete specified dimension(행 전체 또는 열 전체)

            dimension: str (one of ROWS, COLUMNS)
            min_index, max_index: int
            returns None
            """
            request = {
                'deleteDimension': {
                    'range': {
                        'sheetId': self._sheet.sheet_id,
                        'dimension': dimension,
                        'startIndex': min_index - 1,
                        'endIndex': max_index
                    }
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def update_frozen_row_count(self, count):
            """Request to update frozen row count (행 고정)

            count: int
            returns: None
            """
            request = {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': self._sheet.sheet_id,
                        'gridProperties': {
                            'frozenRowCount': count
                        }
                    },
                    'fields': 'gridProperties.frozenRowCount'
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)

        def update_sheet_hidden(self, is_hidden):
            """Request to update sheet's hidden state.

            is_hidden: boolean
            returns: None
            """
            request = {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': self._sheet.sheet_id,
                        'hidden': is_hidden

                    },
                    'fields': 'hidden'
                }
            }
            self._sheet.client.requests_container.deposit(self._sheet.parent_spreadsheet.spreadsheet_id, request)




