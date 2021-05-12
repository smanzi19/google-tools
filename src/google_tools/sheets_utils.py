import os
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Sample sheet
SAMPLE_SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
SAMPLE_RANGE_NAME = 'Class Data!A2:E'

class GoogleSheetsWB:
    def __init__(self, wb_dict):
        self.wb = wb_dict

    def add_wb_properties(self, wb_properties):
        self.wb['properties'].update(wb_properties)

    def add_sheets(self, sheets_to_add):
        sheets_to_add = self.wb.get('sheets', []) + [sheets_to_add]
        self.wb['sheets'] = sheets_to_add

    def upload_wb(self, service):
        sheet_file = service.spreadsheets().create(body=self.wb).execute()
        self.wb = sheet_file
        self.spreadsheetId = sheet_file['spreadsheetId']
        return sheet_file

    def upload_data(self,
                    sheet_name,
                    cell_range_key,
                    values,
                    majorDimension,
                    service,
                    valueInputOption='USER_ENTERED',
                    append=False):
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
