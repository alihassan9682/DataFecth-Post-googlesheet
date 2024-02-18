from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import os
from datetime import datetime

A = ["https://www.emag.ro/search/?ref=effective_search", 2, 5, 1, 6, 8, 1]
l = [A, A, A, A]


def append_values(spreadsheet_id, range_name, value_input_option, values):
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
            with open("token.json", "w") as token:
                token.write(credentials.to_json())

    try:
        service = build("sheets", "v4", credentials=credentials)

        # Splitting values list into two parts: first cell in Column A and rest from Column O
        value_cell_A = [values[i][:1] for i in range(len(values))]
        value_rest_O = [values[i][1:] for i in range(len(values))]

        # Define ranges for appending data
        range_start = 4
        range_end = range_start + len(values[0])
        range_name_A = f"{range_name}!A{range_start}"
        range_name_O = f"{range_name}!O{range_start}:V{range_end}"

      

        current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Append the date and time to the merged cell in Column A
        merge_range = f"{range_name}!A{range_start}"
        result_datetime = (
            service.spreadsheets()
            .values()
            .append(
                spreadsheetId=spreadsheet_id,
                range=merge_range,
                valueInputOption=value_input_option,
                body={
                    "values": [
                        [f"Ranks for the desired products at {current_datetime}"]
                    ]
                },
            )
            .execute()
        )

        print(
            f"{result_datetime.get('updates').get('updatedCells')} cells appended with current date and time."
        )

        # Append the first cell value to Column A
        result_A = (
            service.spreadsheets()
            .values()
            .append(
                spreadsheetId=spreadsheet_id,
                range=range_name_A,
                valueInputOption=value_input_option,
                body={"values": value_cell_A},
            )
            .execute()
        )
        print(
            f"{result_A.get('updates').get('updatedCells')} cells appended to Column A."
        )

        # Append the rest of the values to Column O
        result_O = (
            service.spreadsheets()
            .values()
            .append(
                spreadsheetId=spreadsheet_id,
                range=range_name_O,
                valueInputOption=value_input_option,
                body={"values": value_rest_O},
            )
            .execute()
        )
        print(
            f"{result_O.get('updates').get('updatedCells')} cells appended to Column O."
        )

        # Highlight the entire row
        format_request = {
            "repeatCell": {
                "range": {
                    "sheetId": 0,
                    "startRowIndex": range_start - 1,
                    "endRowIndex": range_start,
                    "startColumnIndex": 0,
                    "endColumnIndex": 0,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.57,
                            "green": 0.988,
                            "blue": 0.749,
                        }  # Change background color to light green
                    }
                },
                "fields": "userEnteredFormat.backgroundColor",
            }
        }
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": [format_request]}
        ).execute()

        return result_A, result_O

    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


append_values(
    "1R5odbc5fOhQ3hoNBIj2WmFRk4502DCncw9FBA8BuiLs",
    "Sheet1",
    "USER_ENTERED",
    l,
)
