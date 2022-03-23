from __future__ import print_function

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account


SERVICE_ACCOUNT_FILE = 'keys.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = None
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

SPREADSHEET_ID = '1g4Qn9gbpx810jHTvgEvGkAN7OOURNgzU04c-ydqPwUo'


def fill_column_data(sheet, from_range, to_range):
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=from_range).execute()
    values = result.get('values', [])
    sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=to_range,
                          valueInputOption="USER_ENTERED", body={"values": values}).execute()

    return len(values)


def copy_paste_cells(sheet, source, destination):
    request_body = {
        "requests": [
            {
                "copyPaste": {
                    "source": source,
                    "destination": destination,
                    "pasteType": "PASTE_NORMAL",
                    "pasteOrientation": "NORMAL"
                }
            }
        ]
    }
    sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID,
                      body=request_body).execute()


def add_borders(sheet, rows):
    request_body = {"requests": [
        {
            "updateBorders": {
                    "range": {
                        "sheetId": 812605565,
                        "startRowIndex": 0,
                        "endRowIndex": rows,
                        "startColumnIndex": 0,
                        "endColumnIndex": 18
                    },
                    "top": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                    "bottom": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                    "right": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                    "left": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                    "innerHorizontal": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                    }
        }
    ]}
    sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID,
                      body=request_body).execute()


def fill_serial_nos(sheet, rows):
    request_body = {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": 812605565,
                        "startRowIndex": 1,
                        "endRowIndex": rows,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "cell": {
                        "userEnteredValue": {
                            "formulaValue": "=ROW()-1"
                        },
                    },
                    "fields": "userEnteredValue"
                }
            }
        ]}
    sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID,
                      body=request_body).execute()


def main():
    try:
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()

        no_of_stocks = fill_column_data(
            sheet, from_range="Stocks_near_52_week_high!C3:C", to_range="21st Mar!C2")

        no_of_rows = no_of_stocks + 1

        # add_borders(sheet, rows=no_rows)

        fill_column_data(
            sheet, from_range="Stocks_near_52_week_high!F3:F", to_range="21st Mar!F2")
        fill_column_data(
            sheet, from_range="Stocks_near_52_week_high!E3:E", to_range="21st Mar!G2")

        copy_paste_cells(sheet, source={
            "sheetId": 812605565,
            "startRowIndex": 1,
            "endRowIndex": 2,
            "startColumnIndex": 3,
            "endColumnIndex": 5
        },
            destination={
            "sheetId": 812605565,
            "startRowIndex": 2,
            "endRowIndex": no_of_rows,
            "startColumnIndex": 3,
            "endColumnIndex": 5
        })

        copy_paste_cells(sheet, source={
            "sheetId": 812605565,
            "startRowIndex": 1,
            "endRowIndex": 2,
            "startColumnIndex": 7,
            "endColumnIndex": 18
        },
            destination={
            "sheetId": 812605565,
            "startRowIndex": 2,
            "endRowIndex": no_of_rows,
            "startColumnIndex": 7,
            "endColumnIndex": 18
        })

        fill_serial_nos(sheet, rows=no_of_rows)

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()
