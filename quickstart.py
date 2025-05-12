import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import google.auth

import pandas as pd
import string

# If modifying these scopes, delete the file token.json.
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
SAMPLE_RANGE_NAME = "A1:F2"


def create(creds, title):
    """
    Creates the Sheet the user has access to.
    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
    """
    # creds, _ = google.auth.default()
    # pylint: disable=maybe-no-member
    try:
        service = build("sheets", "v4", credentials=creds)
        spreadsheet = {"properties": {"title": title}}
        spreadsheet = (
            service.spreadsheets()
            .create(body=spreadsheet, fields="spreadsheetId")
            .execute()
        )
        print(f"Spreadsheet ID: {(spreadsheet.get('spreadsheetId'))}")
        return spreadsheet.get("spreadsheetId")
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


def get_values(creds, spreadsheet_id, range_name):
    """
    Creates the batch_update the user has access to.
    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
    """
    # pylint: disable=maybe-no-member
    try:
        service = build("sheets", "v4", credentials=creds)

        result = (
            service.spreadsheets()
            .values()
            .get(spreadsheetId=spreadsheet_id, range=SAMPLE_RANGE_NAME)
            .execute()
        )
        rows = result.get("values", [])
        print(f"{len(rows)} rows retrieved")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


def update_values(
    spreadsheet_id,
    creds,
    range_name,
    value_input_option,
    values,
):
    """
    Creates the batch_update the user has access to.
    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
    """
    # creds, _ = google.auth.default()
    # pylint: disable=maybe-no-member
    try:
        service = build("sheets", "v4", credentials=creds)
        # values = [
        #     [
        #         # Cell values ...
        #     ],
        #     # Additional rows ...
        # ]
        body = {"values": values}
        result = (
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption=value_input_option,
                body=body,
            )
            .execute()
        )
        print(f"{result.get('updatedCells')} cells updated.")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.

    https://developers.google.com/workspace/drive/api/quickstart/python
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=8080)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    # reading
    values = get_values(creds, SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME)

    if not values:
        print("No data found.")
        return

    for row in values['values']:
        print(row)

    # SPREADSHEET_ID = "1p3Vu3p1w4Ehwlfqlw94cJeUk91ZrWVizYLHeDAbaeW4"
    SPREADSHEET_ID = "1j42_wob0jwSyPgNNsOwZp7v1HxL9QZBS91Bh5bY_mxo"
    # creating
    # create(creds, "Test Sheet 2")

    # update_values(
    #     "1j42_wob0jwSyPgNNsOwZp7v1HxL9QZBS91Bh5bY_mxo",
    #     creds,
    #     "A1:C2",
    #     "USER_ENTERED",
    #     [["A", "B"], ["C", "D"]],
    # )

    df = pd.read_csv('sales_data_sample.csv', encoding='latin1')
    df = df.replace({pd.NA: None}) # run into issue with NA values in csv
    # print(df.columns.tolist())

    rows, columns = df.shape

    alpha = string.ascii_uppercase[columns - 1]  # return Y

    update_values(
        SPREADSHEET_ID,
        creds,
        f"A1:{alpha}" + str(rows + 1),  # including the header of the csv
        "USER_ENTERED",
        [df.columns.tolist()] + df.values.tolist(),
    )


if __name__ == "__main__":
    main()
