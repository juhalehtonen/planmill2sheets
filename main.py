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

# Google Sheets Spreadsheet ID where data gets written
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')

# PlanMill configuration
PLANMILL_CLIENT_ID = os.getenv('PLANMILL_CLIENT_ID')
PLANMILL_CLIENT_SECRET = os.getenv('PLANMILL_CLIENT_SECRET')
PLANMILL_REDIRECT_URI = os.getenv('PLANMILL_REDIRECT_URI')
PLANMILL_API_ENDPOINT = os.getenv('PLANMILL_API_ENDPOINT')
PLANMILL_AUTH_URL = os.getenv('PLANMILL_AUTH_URL')
PLANMILL_TOKEN_URL = os.getenv('PLANMILL_TOKEN_URL')
PLANMILL_GRANT_TYPE = 'Authorization code'

# OfficeVibe configuration
OFFICEVIBE_API_KEY = os.getenv('OFFICEVIBE_API_KEY')

# Freshdesk configuration
FRESHDESK_API_KEY = os.getenv('FRESHDESK_API_KEY')
FRESHDESK_DOMAIN = os.getenv('FRESHDESK_DOMAIN')

# Let's authenticate with PlanMill API and fetch the latest Opportunities data
import requests
from oauthlib.oauth2 import BackendApplicationClient
from requests_oauthlib import OAuth2Session
import pandas as pd
import json
from pandas.io.json import json_normalize
from flatten_json import flatten

# To generate date strings
import datetime

def get_officevibe_data():
    endpoint = "https://app.officevibe.com/api/v2/engagement"
    # Generate a list of dates for this year
    year_today = datetime.date.today().strftime("%Y")
    date_time_index = pd.date_range(start='1/1/'+year_today, periods=12, freq='M')
    dti_to_datetimes = date_time_index.to_pydatetime()
    dt_strings = list(map(lambda x: x.strftime("%Y-%m-%d"), dti_to_datetimes))

    params = 	{
        "groupNames": [],
        "dates": dt_strings
    }
    headers = {"Authorization": "Bearer " + OFFICEVIBE_API_KEY}
    json_response = requests.get(endpoint, params=params, headers=headers).json()

    # Get data
    d = json_response['data']['weeklyReports']

    # Normalize json
    jn = json_normalize(data=d, record_path='metricsValues', errors='ignore', meta=['date'])

    # ..and then convert that to a more-easy-to-import-elsewhere CSV file.
    # `index=False` removes the by-default first column of indexes.
    csv_string = jn.to_csv(None, index=False, encoding='utf-8')

    return csv_string

import io

# Get FreshDesk data
def get_freshdesk_data(api_path):
    api_key = FRESHDESK_API_KEY
    domain = FRESHDESK_DOMAIN
    password = "x"
    json_response = requests.get("https://"+ domain +".freshdesk.com/api/v2/"+api_path, auth = (api_key, password))

    # Let's read our JSON response into a Pandas DataFrame...
    df = pd.read_json(json_response.content)

    # ..and then convert that to a more-easy-to-import-elsewhere CSV file.
    # `index=False` removes the by-default first column of indexes.
    csv_string = df.to_csv(None, index=False, encoding='utf-8')

    return csv_string

# Get CSV data return it yes
def get_planmill_data(api_path):
    client = BackendApplicationClient(client_id=PLANMILL_CLIENT_ID)
    oauth = OAuth2Session(client=client)
    token = oauth.fetch_token(token_url=PLANMILL_TOKEN_URL, client_id=PLANMILL_CLIENT_ID, client_secret=PLANMILL_CLIENT_SECRET)

    # Special treatment for nested Reports
    if 'Actual' in api_path:
        # Fetch Opportunities from PlanMill API
        headers = {'Accept': 'text/csv'}
        csv_response = oauth.get(PLANMILL_API_ENDPOINT + api_path, headers=headers)

        print('processing actual utilization report..')
        df = pd.read_csv(
            io.BytesIO(csv_response.content),
            encoding='utf8',
            decimal=".",
            sep="\t",
            parse_dates=["Period"],
            names=["Person", "Period", "Actual capacity", "Reported", "Billable", "Non-billable", "Actual utilization", "Absences"]
        )
        df['Period'] = df['Period'].dt.strftime('%Y%m%d')

        # There has to be a better way..
        df['Actual capacity'] = df['Actual capacity'].str.replace(',','.')
        df['Reported'] = df['Reported'].str.replace(',','.')
        df['Billable'] = df['Billable'].str.replace(',','.')
        df['Non-billable'] = df['Non-billable'].str.replace(',','.')
        df['Actual utilization'] = df['Actual utilization'].str.replace(',','.')
        df['Actual utilization'] = df['Actual utilization'].str.replace(' %','')
        df['Absences'] = df['Absences'].str.replace(',','.')

        csv_string = df.to_csv(None, index=False, encoding='utf-8')
        return csv_string

    # Special treatment for nested Reports
    if 'Revenues' in api_path:
        # Fetch Opportunities from PlanMill API
        headers = {'Accept': 'text/csv'}
        csv_response = oauth.get(PLANMILL_API_ENDPOINT + api_path, headers=headers)

        print('processing revenue report..')
        df = pd.read_csv(
            io.BytesIO(csv_response.content),
            encoding='utf8',
            decimal=".",
            sep="\t",
            parse_dates=["Date", "Invoice date", "Year/Month"],
            names=["Year/Month", "Date", "Customer", "Project", "Revenue item", "Sales order / item", "Product", "Project manager", "Billing rule", "à price", "Forecast", "Actual", "Invoiced", "Invoice number", "Invoice date"]
        )
        df['Date'] = df['Date'].dt.strftime('%Y%m%d')
        df['Invoice date'] = df['Invoice date'].dt.strftime('%Y%m%d')
        df['Year/Month'] = df['Year/Month'].dt.strftime('%Y%m')

        csv_string = df.to_csv(None, index=False, encoding='utf-8')
        return csv_string

    # Special treatment for nested Reports
    if 'Time' in api_path:
        # Fetch Opportunities from PlanMill API
        headers = {'Accept': 'text/csv'}
        csv_response = oauth.get(PLANMILL_API_ENDPOINT + api_path, headers=headers)

        print('processing time balance report..')
        df = pd.read_csv(
            io.BytesIO(csv_response.content),
            encoding='utf8',
            decimal=".",
            sep="\t",
            parse_dates=["Start", "Finish"],
            names=["Team", "Person", "Start", "Finish", "Last month", "Balance", "Balance adjust", "Balance maximum", "Capacity", "Normal time", "Overtime & on-call"]
        )
        df['Start'] = df['Start'].dt.strftime('%Y%m%d')
        df['Finish'] = df['Finish'].dt.strftime('%Y%m%d')

        # There has to be a better way..
        df['Balance'] = df['Balance'].str.replace(',','.')
        df['Last month'] = df['Last month'].str.replace(',','.')
        df['Balance adjust'] = df['Balance adjust'].str.replace(',','.')
        df['Capacity'] = df['Capacity'].str.replace(',','.')
        df['Normal time'] = df['Normal time'].str.replace(',','.')
        df['Overtime & on-call'] = df['Overtime & on-call'].str.replace(',','.')

        csv_string = df.to_csv(None, index=False, encoding='utf-8')
        return csv_string

    # Fetch Opportunities from PlanMill API
    headers = {'Accept': 'application/json'}
    json_response = oauth.get(PLANMILL_API_ENDPOINT + api_path, headers=headers)

    # Let's read our JSON response into a Pandas DataFrame...
    df = pd.read_json(json_response.content)

    # ..and then convert that to a more-easy-to-import-elsewhere CSV file.
    # `index=False` removes the by-default first column of indexes.
    csv_string = df.to_csv(None, index=False, encoding='utf-8')

    return csv_string

# Import requirements for using Google Spreadsheet API V4
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Helper functions to operate the Google Sheets API
def find_sheet_id(service, index):
    print('index is {}'.format(index))
    # ugly, but works
    sheets_with_properties = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID, fields='sheets.properties') .execute().get('sheets')
    return sheets_with_properties[index]['properties']['sheetId']

def build_gsheet_body(csv_data, sheet_id):
    body = {
        'requests': [{
            'pasteData': {
                "coordinate": {
                    "sheetId": sheet_id,
                    "rowIndex": "0",  # adapt this if you need different positioning
                    "columnIndex": "0", # adapt this if you need different positioning
                },
                "data": csv_data,
                "type": 'PASTE_NORMAL',
                "delimiter": ',',
            }
        }]
    }
    return body

# Build google credentials
def build_google_creds():
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

    return creds

# Helper to get last day of month. any_day is is a datetime.date
# e.g. last_day_of_month(datetime.date(2012, 13, 1))
def last_day_of_month(any_day):
    next_month = any_day.replace(day=28) + datetime.timedelta(days=4)  # this will never fail
    return next_month - datetime.timedelta(days=next_month.day)

# Run Google Sheets API authentication and push CSV file contents to Spreadsheet
def main():
    creds = build_google_creds()

    # Build the API service object
    service = build('sheets', 'v4', credentials=creds)

    # Date strings for various query parameters
    date_today = datetime.date.today()
    # datestring_last_of_this_month = last_day_of_month(date_today).strftime("%Y-%m-%d")
    datestring_today = date_today.strftime("%Y-%m-%d")
    datestring_this_year = date_today.strftime("%Y")

    # Get CSV data from PlanMill and OfficeVibe
    csv_data_opportunities = get_planmill_data(api_path='opportunities?rowcount=3000')
    csv_data_projects = get_planmill_data(api_path='projects?rowcount=3000')
    csv_data_salesorders = get_planmill_data(api_path='salesorders?rowcount=3000')
    csv_data_revenue = get_planmill_data(api_path='reports/Revenues%20summary%20by%20month?param1=-1&param2=-1&param3='+datestring_this_year+'&param4='+datestring_this_year+'-01-01T00%3A00%3A00.000%2B0200&param5='+datestring_this_year+'-12-31T00%3A00%3A00.000%2B0200&param6=-1&param7=-1&param8=-1&param9=-1&param10=-1&param11=-1&param12=-1&param13=-1&rowcount=3000')
    csv_data_utilization = get_planmill_data(api_path='reports/Actual%20billable%20utilization%20rate%20analysis%20by%20person?param1=23&param3=-1&exportType=detailed&rowcount=3000')
    csv_data_timebalance = get_planmill_data(api_path='reports/Time%20balance%20by%20person?param3='+datestring_today+'T00%3A00%3A00.000%2B0200&exportType=detailed&rowcount=3000')
    csv_data_officevibe = get_officevibe_data()
    csv_data_freshdesk = get_freshdesk_data("tickets")

    # Create an ordered list of PlanMill data. The order is important, because
    # it will correspond to the sheets in the spreadsheet.
    csv_data = [
        csv_data_opportunities,
        csv_data_projects,
        csv_data_salesorders,
        csv_data_revenue,
        csv_data_utilization,
        csv_data_timebalance,
        csv_data_officevibe,
        csv_data_freshdesk
    ]

    # Loop over all desired API responses
    for num, data in enumerate(csv_data, start=0):
        # Get the sheet id
        sheet_id = find_sheet_id(service, num)

        # Build body for request
        body = build_gsheet_body(
            csv_data=data,
            sheet_id=sheet_id
        )

        #print(sheet_id)

        # Finally send the request to Google Sheets
        request = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body)
        response = request.execute()
        print('done with {}'.format(num))

if __name__ == '__main__':
    main()
