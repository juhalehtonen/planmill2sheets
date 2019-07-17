# PlanMill2Sheets

This tool fetches a JSON response from the PlanMill Opportunities API, converts it to a CSV file, and uploads it to a configured Google Spreadsheet.

Note that this tool writes over all previous data in the configured Spreadsheet, so some caution should be taken when running this.

## Setup

1. Copy `.env.example` to `.env` and fill in the values.
2. Install Python dependencies with `pipenv`

## Usage

Setup a new app that uses Google Sheets API and download the credentials JSON file into `credentials.json`. You can do this [here](https://console.developers.google.com/).

Use OAuth and set the scopes to allow read/write access to spreadsheets.

Then run `pipenv run python3 planmill2sheets.py` and you're done. You need to authorize yourself to Google once before.
