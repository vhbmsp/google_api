# -*- coding: utf-8 -*-

from google.oauth2 import service_account
from googleapiclient import discovery
from googleapiclient.http import MediaFileUpload
import mimetypes
import pathlib

path = pathlib.Path(__file__).parent.absolute()

DRIVE_API_SCOPES = ["https://www.googleapis.com/auth/drive"]
DRIVE_API_CREDENTIALS = str(path) + "/../google_api_credentials.json"


class Drive:
    credentials = None
    drive_service = None

    # ------------------------------------------------------------------------
    # Get Credentials
    def get_drive_credentials(self):

        credentials = service_account.Credentials

        self.credentials = credentials.from_service_account_file(
            DRIVE_API_CREDENTIALS, scopes=DRIVE_API_SCOPES
        )

    # ------------------------------------------------------------------------
    # Get Drive Api Service
    def get_drive_service(self):

        if self.credentials is None:
            self.get_drive_credentials()

        drive = discovery.build(
            "drive",
            "v3",
            credentials=self.credentials,
            cache_discovery=False
        )

        self.drive_service = drive.files()

    # ------------------------------------------------------------------------
    # Get list of files from folder
    def get_files_list(self, folder_id, search_filename=None):

        if self.drive_service is None:
            self.get_drive_service()

        search_query = "trashed = false " \
                       "and '{}' in parents ".format(folder_id)

        if search_filename:
            search_query += " and name contains '{}' ".format(search_filename)

        result = self.drive_service.list(
           q=search_query,
           spaces="drive",
           orderBy="name",
           pageSize=100,
           supportsAllDrives=True,
           includeItemsFromAllDrives=True,
           fields="nextPageToken, files(id, name)",
        ).execute()

        rows = result.get("files", [])

        return rows

    # ------------------------------------------------------------------------
    # create file
    def _drive_create_file(self, folder_id, name, google_docs_type, filename=""):

        mimetypes.init()
        media_body = None

        if self.drive_service is None:
            self.get_drive_service()

        if filename:
            mime_data = mimetypes.guess_type(filename)
            media_body = MediaFileUpload(filename, mimetype=mime_data[0])

        file_metadata = {
            'name': name,
            'parents': [folder_id],
            'mimeType': google_docs_type,
        }

        drive_result = self.drive_service.create(
            body=file_metadata,
            fields='id',
            media_body=media_body,
            supportsAllDrives=True
        ).execute()

        file_id = drive_result["id"]
        return file_id

    # create empty spreadsheet or upload
    def create_sheet(self, folder_id, name, filename=""):

        return self._drive_create_file(
            folder_id,
            name,
            "application/vnd.google-apps.spreadsheet",
            filename
        )

    # create file or upload
    def create_file(self, folder_id, name, filename=""):

        return self._drive_create_file(
            folder_id,
            name,
            "",
            filename
        )
