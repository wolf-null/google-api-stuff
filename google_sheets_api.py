import os.path

import googleapiclient.discovery
from googleapiclient.discovery import build

from google.oauth2 import service_account

from typing import List, Union
import re


class GoogleSheetsApiInterface:
    """
    One may use this to read, write and paint cells in Google Sheets via Service Account.

    WARNING: It won't work with other authorization mechanisms (will require some modifications in connect() method

    To use this, you'll need an account in Google Cloud Platform.

    1. Go to IAM & Admin ( somewhere here: https://console.cloud.google.com/iam-admin/serviceaccounts )
    2. Then Service Accounts (at the left sidebar) and create a service account.
    3. Then press '...' at the most-right column (actions)
    4. it will rise a popup menu. Click manage keys -> Add key
    5. After adding a key download it

    PYTHON REQUIREMENTS:
        - Python 3.xx
        pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

    SOURCES:
        - Official docs: https://developers.google.com/sheets/api/quickstart/python
        - Some stackoverflow shamanism: https://stackoverflow.com/questions/59727373/update-cell-background-color-for-google-spreadsheet-using-python

    NOTES:
        - Spreadsheet ID is some sort of '1lSFT-Nqka-cXR3BtGxTrOPgA5ZAaafzd' thing one can find in the URL when the google sheet is opened in a browser
        - A1 Notation have form <sheetname>!
    """

    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']  # Or use '...spreadsheets.readonly' if you wanna readonly

    class BadQuery(Exception):
        pass

    # Some things for converting A1-notation to indices
    _regex_numbers = re.compile('[0-9]+')
    _regex_letters = re.compile('[a-zA-Z]*')
    _regex_a1 = re.compile('^[a-zA-Z]*[0-9]*$')

    def __init__(self):
        self._credentials = None  # type: service_account.Credentials
        self._service = None  # type: googleapiclient.discovery.Resource
        self._spreadsheet_interface = None
        self._sheet_ids = dict()
        self._current_spreadsheet_id = None

    # ---------------------------------------- CONNECTION / LOADING ----------------------------------------------------

    def connect(self, service_account_json_key_path:str=''):
        """
        service_account_json_key_path is a path to OAuth2 json token. Used to authorize.
        You can get it from Google Cloud (see above)
        """
        if not os.path.exists(service_account_json_key_path):
            raise FileNotFoundError('[google_sheets_api.py]: JSON file {} not found'.format(service_account_json_key_path))

        self._credentials = service_account.Credentials.from_service_account_file(service_account_json_key_path, scopes=GoogleSheetsApiInterface.SCOPES)  # type: service_account.Credentials
        self._service = build('sheets', 'v4', credentials=self._credentials)
        self._spreadsheet_interface = self._service.spreadsheets()

        self._current_spreadsheet_id = None

    def retrieve_spreadsheet_info(self, spreadsheet_id:str):
        return self._spreadsheet_interface.get(spreadsheetId=spreadsheet_id).execute()

    def retrieve_sheet_ids(self, spreadsheet_id:str) -> dict:
        """ Retrieve and parse """
        spreadsheet_info = self.retrieve_spreadsheet_info(spreadsheet_id=spreadsheet_id)
        sheet_ids_mapping = dict()
        if 'sheets' in spreadsheet_info:
            for sheet_info in spreadsheet_info['sheets']:
                sheet_props = sheet_info['properties']
                sheet_id = sheet_props['sheetId']
                sheet_title = sheet_props['title']
                sheet_ids_mapping[sheet_title] = sheet_id
        return sheet_ids_mapping

    def select_spreadsheet_id(self, spreadsheet_id:str):
        """ Returns sheet id. Requires set_spreadsheet_id() to be executet before """

        if self._spreadsheet_interface is None:
            raise Exception('[google_sheets_api.py|set_spreadsheet]: no connection or failed. Use connect() before any get-set activities')

        self._current_spreadsheet_id = spreadsheet_id
        self._sheet_ids = self.retrieve_sheet_ids(spreadsheet_id=spreadsheet_id)

    def get_sheet_id(self, sheet_name:str) -> int:
        """ Returns sheet id. Requires set_spreadsheet_id() to be executet before """
        if self._current_spreadsheet_id is None:
            raise Exception('[google_sheets_api.py|set_spreadsheet]: current spreadsheet is not set. Use set_spreadsheet_id(). It is not the same as sheetId. One can find it from the google document URL in browser URL line')

        if sheet_name in self._sheet_ids:
            return self._sheet_ids[sheet_name]
        else:
            raise GoogleSheetsApiInterface.BadQuery('[google_sheets_api.py|set_spreadsheet]: {} not found in sheets: {}'.format(sheet_name, self._sheet_ids.keys()))

    # ---------------------------------------------- GET/SET TABLE VALUES ----------------------------------------------
    @staticmethod
    def ensure_rectangular_shape(a1_cell_range:str, values:List[List[Union[str, int]]], default_vals='') -> List[List[Union[str, int]]] :
        start_row, start_column, end_row, end_column, sheet_name = GoogleSheetsApiInterface.a1_notation_to_grid_range(a1_range=a1_cell_range)

        if len(values) <= end_row - start_row:
            values += [list() for add_row in range( max((end_row-start_row + 1) - len(values), 0))]
        for row in range(end_row-start_row+1):
            if len(values[row]) <= end_column - start_column:
                values[row] += [default_vals for add_col in range( max((end_column - start_column + 1) - len(values[row]), 0))]

        return values

    def get_cell_values(self, a1_cell_range:str= 'A1:A1') -> List[List[str]]:
        """"""
        if self._current_spreadsheet_id is None:
            raise Exception('[google_sheets_api.py|set_spreadsheet]: current spreadsheet is not set. Use set_spreadsheet_id(). It is not the same as sheetId. One can find it from the google document URL in browser URL line')

        result = self._spreadsheet_interface.values().get(spreadsheetId=self._current_spreadsheet_id, range=a1_cell_range).execute()
        values = result.get('values', [])

        values = GoogleSheetsApiInterface.ensure_rectangular_shape(a1_cell_range=a1_cell_range, values=values)

        return values

    def set_cell_values(self, a1_cell_range:str= '', values:List[List[str]]=list()):
        if self._current_spreadsheet_id is None:
            raise Exception('[google_sheets_api.py|set_spreadsheet]: current spreadsheet is not set. Use set_spreadsheet_id(). It is not the same as sheetId. One can find it from the google document URL in browser URL line')

        spreadsheet_id = self._current_spreadsheet_id

        query = {
            "values": values,
        }

        result = self._spreadsheet_interface.values().update(
            spreadsheetId=spreadsheet_id,
            range=a1_cell_range,
            valueInputOption='USER_ENTERED',
            body=query)

        return result.execute()

    def set_cell_format(self, a1_cell_range:str, cell_format:dict=None):
        """Applies cell_format attributes (see official docs) to each cell in specified __a1_cell_range__"""

        if self._current_spreadsheet_id is None:
            raise Exception('[google_sheets_api.py|set_spreadsheet]: current spreadsheet is not set. Use set_spreadsheet_id(). It is not the same as sheetId. One can find it from the google document URL in browser URL line')

        # Convert A1 notation to cell indices (Feiled to feed A1 to the query, so need conversion)
        start_row, start_column, end_row, end_column, sheet_name = GoogleSheetsApiInterface.a1_notation_to_grid_range(a1_range=a1_cell_range)
        sheet_id = self.get_sheet_id(sheet_name)

        # Replicate cell style to the whole range
        style_per_row = list()
        for row in range(start_row-1, end_row):
            style_per_column = [cell_format for column in range(start_column-1, end_column)]
            style_per_row.append(
                {"values": style_per_column}
            )

        # Construct the query
        """
        BASED ON SOURCE: https://stackoverflow.com/questions/59727373/update-cell-background-color-for-google-spreadsheet-using-python
        """
        body = {
            "requests": [
                {
                    "updateCells": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": start_row-1, # Indexing starts with 0
                            "startColumnIndex": start_column-1,
                            "endRowIndex": end_row,   # Indexing ends are non-inclusive
                            "endColumnIndex": end_column
                        },
                        "rows": style_per_row,
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                }
            ]
        }

        # Execute the query
        result = self._spreadsheet_interface.batchUpdate(spreadsheetId=self._current_spreadsheet_id, body=body).execute()
        return result

    def set_cell_background_color(self, a1_cell_range: str, rgb_normalized: List[float] = None):
        """Sets one and the same color for within __a1_cell_range__"""

        if isinstance(rgb_normalized, list):
            if len(rgb_normalized) != 3:
                raise GoogleSheetsApiInterface.BadQuery(
                    'rgb_normalized should be list or tuple of three floats 0...1; RGBa is not supported')
            if not all([0 <= x <= 1 for x in rgb_normalized]):
                raise GoogleSheetsApiInterface.BadQuery(
                    'rgb_normalized should be list or tuple of three floats between 0 and 1')
        else:
            raise GoogleSheetsApiInterface.BadQuery(
                'rgb_normalized should be list or tuple of three floats 0...1; Color names is not supported')

        cell_format = {
            "userEnteredFormat":
                {
                    "backgroundColor":
                        {
                            "red": rgb_normalized[0],
                            "green": rgb_normalized[1],
                            "blue": rgb_normalized[2]
                        }
                }
        }

        self.set_cell_format(a1_cell_range=a1_cell_range, cell_format=cell_format)

    def apply_to_each_cell(self, a1_cell_range, func) -> None:
        """
        Apply a function __func__ to each cell.
        Function shall return int or str and receives int or str as a single argument
        """

        # TODO: Pass row/column to func() if it will receive

        cells = api.get_cell_values(a1_cell_range=a1_cell_range)

        def apply_to_whole_row(row):
            return list(map(func, row))

        altered_cells = list(map(apply_to_whole_row, cells))

        api.set_cell_values(a1_cell_range=a1_cell_range, values=altered_cells)

    # ----------------------------------- A1 to CellRange indices conversion -------------------------------------------

    @staticmethod
    def a1_notation_to_grid_coordinate(a1:str):
        if not GoogleSheetsApiInterface._regex_a1.match(a1):
            raise Exception("[google_sheets_api.py]: {} is not an A1 single cell item".format(a1))
        column_str = GoogleSheetsApiInterface._regex_letters.match(a1).group().lower()
        column_str = list(reversed(column_str))
        charges = [(ord(column_str[i]) - ord('a')+1)*(26**i) for i in range(len(column_str))]
        column_idx = sum(charges)

        row_str = GoogleSheetsApiInterface._regex_numbers.findall(a1)[0]
        row_idx = int(row_str)

        return row_idx, column_idx

    @staticmethod
    def a1_notation_to_grid_range(a1_range:str):
        if a1_range.count(':') != 1:
            raise Exception("[google_sheets_api.py]: {} is not an A1 cell range".format(a1_range))

        if a1_range.count('!') == 0:
            sheet_range = a1_range
            sheet_name = None
        elif a1_range.count('!') == 1:
            sheet_name, sheet_range = a1_range.split('!')
        else:
            raise Exception("[google_sheets_api.py]: {} is not an A1 cell range".format(a1_range))

        cell_start_str, cell_end_str = sheet_range.split(':')
        start_row, start_column = GoogleSheetsApiInterface.a1_notation_to_grid_coordinate(cell_start_str)
        end_row, end_column = GoogleSheetsApiInterface.a1_notation_to_grid_coordinate(cell_end_str)
        if start_row > end_row or start_column > end_column:
            raise Exception("[google_sheets_api.py]: One or more of start coordinates are larger than end coordinates for start/end: {}/{}".format(cell_end_str, cell_end_str))
        return start_row, start_column, end_row, end_column, sheet_name

    @staticmethod
    def a1_notation_to_grid_coordinate__testing():
        assert GoogleSheetsApiInterface.a1_notation_to_grid_coordinate('A1') == (1, 1), "Should be A1 is (1, 1)"
        assert GoogleSheetsApiInterface.a1_notation_to_grid_coordinate('A0001') == (1, 1), "Should be A0001 is (1, 1)"
        assert GoogleSheetsApiInterface.a1_notation_to_grid_coordinate('A1000') == (1000, 1), "Should be A1 is (1000, 1)"
        assert GoogleSheetsApiInterface.a1_notation_to_grid_coordinate('Z1') == (1, 26), "Should be A1 is (1, 26)"
        assert GoogleSheetsApiInterface.a1_notation_to_grid_coordinate('Z1000') == (1000, 26), "Should be A1 is (1000, 26)"
        assert GoogleSheetsApiInterface.a1_notation_to_grid_coordinate('AA1') == (1, 27), "Should be AA1 is (1, 27)"
        assert GoogleSheetsApiInterface.a1_notation_to_grid_coordinate('AZ1') == (1, 52), "Should be AA1 is (1, 53=27+26)"
        assert GoogleSheetsApiInterface.a1_notation_to_grid_coordinate('ADG666') == (666, 787), "Should be (666, 787)"


if __name__ == '__main__':
    """
    Some examples: 
    """

    api = GoogleSheetsApiInterface()

    path_to_your_json_key = "You'll have to get your own Google Cloud Service Account and request a key. See class's description above."

    if not os.path.exists(path_to_your_json_key):
        print("[google_sheets_api.py|main]: FILE DOESN'T EXIST: {}".format(path_to_your_json_key))
        exit(0)

    api.connect(path_to_your_json_key)

    # 1. Write down the Google Sheets document ID. Find in in the URL in the browser when opened.

    google_sheet_id = '1lSFT-NqTr-  ABRAKADABRA  -khefkaXsdfsdaf'
    api.select_spreadsheet_id(google_sheet_id)


    # 2. Get some cells !
    cells = api.get_cell_values(a1_cell_range='testing!A1:D2')

    # 3. Increase by 1 (if possible) ;)
    def increment_cell(cell_val):
        try:
            return int(cell_val) + 1
        except Exception:
            return cell_val

    def increment_whole_row(row):
        return list(map(increment_cell, row))

    cells = list(map(increment_whole_row, cells))

    # 4. Set the result back
    api.set_cell_values('testing!A1:D2', cells)

    # 5. Paint those roses... dirty yellow?
    api.set_cell_background_color(a1_cell_range='testing!A1:D2', rgb_normalized=[0.9, 0.9, 0.7])


    # -----  An example of apply() method usage: filling cells with numbers

    # 1. Some generator for the generator god!
    def counter_generator():
        iteration = 0
        while True:
            yield iteration
            iteration += 1

    counter = counter_generator()

    # 2. Apply
    api.apply_to_each_cell('testing!A3:D6', lambda x: next(counter))

    # Paint this shit with yet another weird color
    results = api.set_cell_background_color(a1_cell_range='testing!A3:D6', rgb_normalized=[0.8, 0.9, 0.8])

    # ...

    # PROFIT

