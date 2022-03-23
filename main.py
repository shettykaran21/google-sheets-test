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


def main():
    try:
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()

        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                    range="Stocks_near_52_week_high!C3:C").execute()
        symbols = result.get('values', [])
        sheet.values().update(spreadsheetId=SPREADSHEET_ID, range="21st Mar!C2",
                              valueInputOption="USER_ENTERED", body={"values": symbols}).execute()

        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                    range="Stocks_near_52_week_high!F3:F").execute()
        price_on_appearance = result.get('values', [])
        sheet.values().update(spreadsheetId=SPREADSHEET_ID, range="21st Mar!F2",
                              valueInputOption="USER_ENTERED", body={"values": price_on_appearance}).execute()

        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                    range="Stocks_near_52_week_high!E3:E").execute()
        change = result.get('values', [])
        sheet.values().update(spreadsheetId=SPREADSHEET_ID, range="21st Mar!G2",
                              valueInputOption="USER_ENTERED", body={"values": change}).execute()

        # request_body = {
        #     "requests": [
        #         {
        #             # 'autoFill': {
        #             #     "useAlternateSeries": False,
        #             #     "sourceAndDestination": {
        #             #         "source": {
        #             #             "sheetId": 812605565,
        #             #             "startRowIndex": 1,
        #             #             "endRowIndex": 2,
        #             #             "startColumnIndex": 2,
        #             #             "endColumnIndex": 5
        #             #         },
        #             #         "dimension": "ROWS"
        #             #     }

        #             # }
        #             "copyPaste": {
        #                 "source": {
        #                     "sheetId": 812605565,
        #                     "startRowIndex": 1,
        #                     "endRowIndex": 2,
        #                     "startColumnIndex": 3,
        #                     "endColumnIndex": 5
        #                 },
        #                 "destination": {
        #                     "sheetId": 812605565,
        #                     "startRowIndex": 1,
        #                     "endRowIndex": 2,
        #                     "startColumnIndex": 3,
        #                     "endColumnIndex": 5
        #                 },
        #                 "pasteType": "PASTE_FORMULA"
        #             }
        #         }
        #     ]
        # }
        # response = sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID,
        #                              body=request_body).execute()
        # print(response)

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()
