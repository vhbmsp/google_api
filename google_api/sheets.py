# -*- coding: utf-8 -*-

from google.oauth2 import service_account
from googleapiclient import discovery
import pathlib

path = pathlib.Path(__file__).parent.absolute()

SHEETS_API_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SHEETS_API_CREDENTIALS = str(path) + "/../google_api_credentials.json"


class Sheets:
    credentials = None
    sheet_service = None

    # ------------------------------------------------------------------------
    # Get Credentials
    def get_sheets_credentials(self):

        credentials = service_account.Credentials

        self.credentials = credentials.from_service_account_file(
            SHEETS_API_CREDENTIALS, scopes=SHEETS_API_SCOPES
        )

    # ------------------------------------------------------------------------
    # Get Sheets Api Service
    def get_sheets_service(self):

        if self.credentials is None:
            self.get_sheets_credentials()

        service = discovery.build(
           'sheets',
           'v4',
           credentials=self.credentials,
           cache_discovery=False
        )
        self.sheet_service = service.spreadsheets()

    # ------------------------------------------------------------------------
    # Read Sheet Values from range
    def get_sheet_values(self, sheet_id, sheet_name, sheet_range):

        if self.sheet_service is None:
            self.get_sheets_service()

        result = self.sheet_service.values().get(
            spreadsheetId=sheet_id,
            range=sheet_name + "!" + sheet_range
        ).execute()
        rows = result.get('values', [])

        return rows

    # ------------------------------------------------------------------------
    # Update Sheet values to range
    def update_sheet_values(self, sheet_id, sheet_name, sheet_range, rows):

        request_body = {
            'values': rows
        }

        self.sheet_service.values().update(
            spreadsheetId=sheet_id,
            range=sheet_name + '!' + sheet_range,
            valueInputOption='USER_ENTERED', body=request_body
        ).execute()

    # ------------------------------------------------------------------------
    # Add sheet_name to book
    def add_sheet(self, sheet_id, sheet_name):
        try:

            # Remove sheet if it already exists
            self.remove_sheet(sheet_id, sheet_name)

            request_body = {
                'requests': [{
                    'addSheet': {
                        'properties': {
                            'title': sheet_name,
                            'tabColor': {
                                'red': 0.44,
                                'green': 0.99,
                                'blue': 0.50
                            }
                        }
                    }
                }]
            }

            self.sheet_service.batchUpdate(
                spreadsheetId=sheet_id,
                body=request_body
            ).execute()

        except Exception as e:
            print(e)

    # ------------------------------------------------------------------------
    # Remove sheet_name from book
    def remove_sheet(self, sheet_id, sheet_name):

        # get book metadata
        sheet_metadata = self.sheet_service.get(
            spreadsheetId=sheet_id
        ).execute()

        sheet_data = sheet_metadata.get('sheets', '')

        page_names = [sh['properties']['title'] for sh in sheet_data]
        page_ids = [sh['properties']['sheetId'] for sh in sheet_data]

        if sheet_name in page_names:
            index = page_names.index(sheet_name)
            page_id = page_ids[index]

            # Remove sheet from book
            request_body = {
                'requests': [{
                    'deleteSheet': {
                        'sheetId': page_id
                    }
                }]
            }

            self.sheet_service.batchUpdate(
                spreadsheetId=sheet_id,
                body=request_body
            ).execute()
