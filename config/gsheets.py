import gspread
import os

from oauth2client.service_account import ServiceAccountCredentials

scope = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
]

cred_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 
            'client_secret.json')

creds = ServiceAccountCredentials.from_json_keyfile_name(cred_path, scope)
client = gspread.authorize(creds)