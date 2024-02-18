import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import os


def get_URLs_ProductIDs(spreadsheet_ids):
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    credentials = None

    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
            with open("token.json", "r") as token:
                token.write(credentials.to_json())
    # pylint: disable=maybe-no-member
    try:
        service = build("sheets", "v4", credentials=credentials)
        result = (
            service.spreadsheets()
            .values()
            .batchGet(spreadsheetId=spreadsheet_ids, ranges="A1:ZZ")
            .execute()
        )
        ranges = result.get("valueRanges", [])
        print(f"{len(ranges)} ranges retrieved")
        result = ranges[0].get("values", [])
        ProductIDs = result[0]
        del ProductIDs[0]
        del result[0]

        URLs = [i[0] for i in result]
        ProductIDs = [i for i in ProductIDs if len(i) > 0]
        return URLs, ProductIDs

    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


# URLs, ProductIDs = get_URLs_ProductIDs("1ONJjNMj4TPJJdNsphrpXP_HSD2flZXAsr05jDxE-PHc")

# print("URLs", URLs)

# print("ProductIDs", ProductIDs)
