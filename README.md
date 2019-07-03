# PlanMill2Sheets

This tool fetches a JSON response from the PlanMill Opportunities API, converts it to a CSV file, and uploads it to a configured Google Spreadsheet.

Note that this tool writes over all previous data in the configured Spreadsheet, so some caution should be taken when running this.

## Setup

1. Copy `.env.example` to `.env` and fill in the values.
2. Install Python dependencies

## Usage

Simply run `python planmill2sheets.py` and you're done.
