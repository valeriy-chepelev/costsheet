# Requirements
* connect.ini with Yandex Tracker read access credentials:

        [DEFAULT]
        token = token_credential_string
        org = org_id_numbers

# Dependencies
* yandex_tracker_client
* python-dateutil
* pandas
* openpyxl
* alive-progress
* colorama

# Notes

Command to build exe with pyinstaller should include '--collect-data grapheme'