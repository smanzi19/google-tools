import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Sample sheet
SAMPLE_SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
SAMPLE_RANGE_NAME = 'Class Data!A2:E'
BUILD_PARAMS = {'serviceName':'sheets',
                'version':'v4'}
CREDENTIALS_PATH = os.environ['GOOGLE_APPLICATIONS_CREDENTIALS_PATH']



def Create_Service(build_params):
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # Search for token creds. If scope permissions change, token.json should be deleted so it will be reissued.
    if os.path.exists(f'{CREDENTIALS_PATH}/token.json'):
        creds = Credentials.from_authorized_user_file(F'{CREDENTIALS_PATH}/token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                f'{CREDENTIALS_PATH}/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open(f'{CREDENTIALS_PATH}/token.json', 'w') as token:
            token.write(creds.to_json())
    build_params['credentials'] = creds
    service = build(**build_params)
    return service


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
