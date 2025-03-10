# Requirements
* connect.ini with Yandex Tracker read access credentials:

        [DEFAULT]
        token = token_credential_string
        org = org_id_numbers

# Dependencies
* yandex_tracker_client
* pandas
* openpyxl
* alive-progress
* colorama

# Notes

Command to build .exe with pyinstaller:
    
    pyinstaller costtrack.py --onefile --collect-data grapheme