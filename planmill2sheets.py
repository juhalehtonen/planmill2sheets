#!/usr/bin/env python
# coding: utf-8

# setup python-getenv
import os
from os.path import join, dirname
from dotenv import load_dotenv

dotenv_path = join(dirname(__file__), '.env')
load_dotenv(dotenv_path)


# If modifying these scopes, delete the file token.pickle
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Path to the CSV file uploaded to Sheets
CSV_FILE_PATH = os.getenv('CSV_FILE_PATH')

# Google Sheets Spreadsheet ID where data gets written
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')

# PlanMill configuration
PLANMILL_CLIENT_ID = os.getenv('PLANMILL_CLIENT_ID')
PLANMILL_CLIENT_SECRET = os.getenv('PLANMILL_CLIENT_SECRET')
PLANMILL_REDIRECT_URI = os.getenv('PLANMILL_REDIRECT_URI')
PLANMILL_API_URL = os.getenv('PLANMILL_API_URL')
PLANMILL_AUTH_URL = os.getenv('PLANMILL_AUTH_URL')
PLANMILL_TOKEN_URL = os.getenv('PLANMILL_TOKEN_URL')
PLANMILL_GRANT_TYPE = 'Authorization code'


# Let's authenticate with PlanMill API and fetch the latest Opportunities data
from oauthlib.oauth2 import BackendApplicationClient
from requests_oauthlib import OAuth2Session
client = BackendApplicationClient(client_id=PLANMILL_CLIENT_ID)
oauth = OAuth2Session(client=client)
token = oauth.fetch_token(token_url=PLANMILL_TOKEN_URL, client_id=PLANMILL_CLIENT_ID, client_secret=PLANMILL_CLIENT_SECRET)
# print(token)

# Fetch Opportunities from PlanMill API
json_response = oauth.get(PLANMILL_API_URL)

# Let's import our required libraries
import pandas as pd


# Let's read our JSON response into a Pandas DataFrame...
df = pd.read_json(json_response.content)

# ..and then convert that to a more-easy-to-import-elsewhere CSV file.
# `index=False` removes the by-default first column of indexes.
df.to_csv(CSV_FILE_PATH, index=False, encoding='utf-8')


# Import requirements for using Google Spreadsheet API V4
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# Helper functions to operate the Google Sheets API
def find_sheet_id(service):
    # ugly, but works
    sheets_with_properties = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID, fields='sheets.properties') .execute().get('sheets')
    return sheets_with_properties[0]['properties']['sheetId']

def push_csv_to_gsheet(csv_path, sheet_id):
    with open(csv_path, 'r') as csv_file:
        csvContents = csv_file.read()
    body = {
        'requests': [{
            'pasteData': {
                "coordinate": {
                    "sheetId": sheet_id,
                    "rowIndex": "0",  # adapt this if you need different positioning
                    "columnIndex": "0", # adapt this if you need different positioning
                },
                "data": csvContents,
                "type": 'PASTE_NORMAL',
                "delimiter": ',',
            }
        }]
    }
    return body


# Run Google Sheets API authentication and push CSV file contents to Spreadsheet
def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    # Build the API service object
    service = build('sheets', 'v4', credentials=creds)

    # Build body for request
    body = push_csv_to_gsheet(
        csv_path=CSV_FILE_PATH,
        sheet_id=find_sheet_id(service)
    )

    # Finally send the request to Google Sheets
    request = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body)
    response = request.execute()
    print(response)
    return response

if __name__ == '__main__':
    main()
