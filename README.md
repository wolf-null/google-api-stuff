# google-api-stuff
*Yet another Google Api wrapper*

## google_sheets_api.py ##

One may use this to read, write and paint cells in Google Sheets via **Service Account**.

WARNING: It won't work with other authorization mechanisms (will require some modifications in connect() method

### SOURCES ###
I used mostly these things to figure out WTF is happening and HOWTO do such a simple thing like editing cells
  - Official docs: https://developers.google.com/sheets/api/quickstart/python
  - Some stackoverflow shamanism: https://stackoverflow.com/questions/59727373/update-cell-background-color-for-google-spreadsheet-using-python

### Description ###

To use this, you'll need an account in Google Cloud Platform AND to create a service account bu following approximately these steps:

1. Go to IAM & Admin ( *~somewhere here: https://console.cloud.google.com/iam-admin/serviceaccounts* )
2. Then -> Service Accounts (at the left sidebar) and create a service account.
3. Then press vertical three dots '...' at the most-right column (*actions*)
4. it will rise a popup menu. Click manage keys -> Add key
5. After adding a key download it

### PYTHON REQUIREMENTS ###
- Python 3.xx
- pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib



### NOTES ###
  - Spreadsheet ID is some sort of '1lSFT-NqkaABRACADABRAgA5ZAaafzd' thing one can find in the URL when the google sheet is opened in a browser
  - A1 Notation (*e.g. A1:D20*) have form sheet_name!start_cell:end_cell where <start_cell> consists of letters.... oh, c'mon, y'all know this stuff
  



