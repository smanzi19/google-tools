import os
import pandas as pd


SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Sample sheet
SAMPLE_SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
SAMPLE_RANGE_NAME = 'Class Data!A2:E'

class GoogleSheetsWB:
    def __init__(self, wb_dict):
        self.wb = wb_dict

    def AddWbProperties(self, wb_properties):
        self.wb['properties'].update(wb_properties)

    def AddSheets(self, sheets_to_add):
        sheets_to_add = self.wb.get('sheets', []) + [sheets_to_add]
        self.wb['sheets'] = sheets_to_add

    def UploadWb(self, service):
        sheet_file = service.spreadsheets().create(body=self.wb).execute()
        self.wb = sheet_file
        self.spreadsheetId = sheet_file['spreadsheetId']
        return sheet_file

    def _process_pandas(self, df, append=False):
        df = df.copy()
        vals = df.values.tolist()
        if not append:
            vals = [df.columns.tolist()] + vals
        return vals

    def UploadData(self,
                   sheet_name,
                   cell_range_key,
                   values,
                   majorDimension,
                   service,
                   valueInputOption='USER_ENTERED',
                   append=False):
        if isinstance(values, pd.DataFrame):
            values = self._process_pandas(values, append)
        sheet_name += '!'
        value_body = {'majorDimension':majorDimension, 'values':values}
        update_dict = {'spreadsheetId':self.spreadsheetId,
                       'valueInputOption':valueInputOption,
                       'range':sheet_name + cell_range_key,
                       'body':value_body}
        if not append:
            service.spreadsheets().values().update(**update_dict).execute()
        else:
            service.spreadsheets().values().append(**update_dict).execute()
