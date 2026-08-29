# Polylogger → Google Sheets

> Google Apps Script to pull Polylogger logs and stats into Google Sheets.

## Setup

1. Open your Google Sheet → **Extensions → Apps Script**
2. Copy the script into the project.
3. Add your Polylogger username as the `POLYLOGGER_USERNAME` Script Property.
4. Reload the sheet and use **Polylogger → Pull my logs + stats**.

## Sheets

The script creates:

* `polylogger_raw_logs` -  raw activity logs
* `polylogger_stats` - activity statistics

