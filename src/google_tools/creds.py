import os
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive']
CREDENTIALS_PATH = os.environ['GOOGLE_APPLICATIONS_CREDENTIALS_PATH']

def Create_Creds():
    creds = None
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
    return creds

creds = Create_Creds()
