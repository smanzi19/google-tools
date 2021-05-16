import pandas as pd

class GoogleSheetsWB:

    def __init__(self, service, wb_dict, spreadsheetId=None):
        self.wb = wb_dict
        self.service = service
        self.spreadsheetId = spreadsheetId

    @classmethod
    def FromExisting(cls, service, google_driver, wb_name):
        google_driver._UpdateDriver()
        spreadsheetId = google_driver.gdrive_files[wb_name]
        wb_dict = service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
        return cls(service, wb_dict, spreadsheetId)

    def _UpdateSpreadsheet(self):
        self.wb = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId).execute()

    def AddWbProperties(self, wb_properties):
        self.wb['properties'].update(wb_properties)

    def AddSheets(self, sheet_properties):
        assert 'title' in sheet_properties['properties'], 'You must name the sheet you want to upload'
        update_dict = {'requests':[{'addSheet':sheet_properties}]}
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=update_dict).execute()
        self._UpdateSpreadsheet()

    def UploadWb(self):
        sheet_file = self.service.spreadsheets().create(body=self.wb).execute()
        self.spreadsheetId = sheet_file['spreadsheetId']
        self._UpdateSpreadsheet()
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
            self.service.spreadsheets().values().update(**update_dict).execute()
        else:
            self.service.spreadsheets().values().append(**update_dict).execute()
