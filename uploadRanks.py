from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import os
from datetime import datetime


def formatRanks(ranks):
    result = [
        ranks[0],
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        *ranks[1:],
    ]
    return result


def uploadRanks(spreadsheet_id, range_name, value_input_option, values):
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

        # Retrieve existing data to calculate the range
        result = (
            service.spreadsheets()
            .values()
            .get(spreadsheetId=spreadsheet_id, range=range_name, majorDimension="ROWS")
            .execute()
        )

        existing_data = result.get("values", [])
        num_existing_rows = len(existing_data)

        # Calculate the new range based on the number of existing rows
        range_start = 23
        range_end = range_start + len(values)
        range_name_A = f"{range_name}!A{range_start}"

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
                insertDataOption="INSERT_ROWS",
            )
            .execute()
        )

        print(
            f"{result_datetime.get('updates').get('updatedCells')} cells appended with current date and time."
        )

        value_cell_A = [formatRanks(i) for i in values]

        result_A = (
            service.spreadsheets()
            .values()
            .append(
                spreadsheetId=spreadsheet_id,
                range=range_name_A,
                valueInputOption=value_input_option,
                body={"values": value_cell_A},
                insertDataOption="INSERT_ROWS",
            )
            .execute()
        )
        print(
            f"{result_A.get('updates').get('updatedCells')} cells appended to Column A."
        )

        # Highlight the entire row for the date-time entry
        format_request_datetime_row = {
            "repeatCell": {
                "range": {
                    "sheetId": 0,
                    "startRowIndex": range_start - 1,
                    "endRowIndex": range_start,
                    "startColumnIndex": 0,
                    "endColumnIndex": len(formatRanks(A)),  # Adjust column end index
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
            spreadsheetId=spreadsheet_id,
            body={"requests": [format_request_datetime_row]},
        ).execute()

        return result_A
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error



# values = [
#     ["https://www.emag.ro/search/?ref=effective_search", 2, 5, 1, 6, 8, 189],
#     ["https://www.emag.ro/search/?ref=effective_search", 2, 5, 1, 6, 8, 189],
#     ["https://www.emag.ro/search/?ref=effective_search", 2, 5, 1, 6, 8, 189],
#     ["https://www.emag.ro/search/?ref=effective_search", 2, 5, 1, 6, 8, 189],
# ]

# uploadRanks(
#     "1R5odbc5fOhQ3hoNBIj2WmFRk4502DCncw9FBA8BuiLs",
#     "product y",
#     "USER_ENTERED",
#     values,
# )
