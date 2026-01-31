import os
import sys
from datetime import datetime

from googleapiclient.discovery import build

# auth.py 모듈 경로 추가
sys.path.append(os.path.expanduser("~/Documents/github_cloud/module_auth"))
import auth


def fetch_sheet_values(spreadsheet_id, sheet_name):
    creds = auth.get_credentials()
    service = build("sheets", "v4", credentials=creds)
    range_a1 = f"{sheet_name}"

    result = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=range_a1,
            majorDimension="ROWS",
        )
        .execute()
    )

    return result.get("values", [])


def fetch_sheet_names(spreadsheet_id):
    creds = auth.get_credentials()
    service = build("sheets", "v4", credentials=creds)
    result = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = result.get("sheets", [])
    sheet_names = [sheet.get("properties", {}).get("title", "") for sheet in sheets]
    return [name for name in sheet_names if name]


def find_next_row_index(spreadsheet_id, sheet_name):
    creds = auth.get_credentials()
    service = build("sheets", "v4", credentials=creds)
    result = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}",
            majorDimension="ROWS",
        )
        .execute()
    )
    values = result.get("values", [])
    last_filled_index = 0
    for idx, row in enumerate(values, start=1):
        if any(str(cell).strip() != "" for cell in row):
            last_filled_index = idx
    return last_filled_index + 1


def write_name_rrn(spreadsheet_id, sheet_name, name, rrn):
    row_index = find_next_row_index(spreadsheet_id, sheet_name)
    today = datetime.now().strftime("%y%m%d")
    creds = auth.get_credentials()
    service = build("sheets", "v4", credentials=creds)
    range_a1 = f"{sheet_name}!C{row_index}:J{row_index}"
    body = {"values": [[today, "", "", name, "", "", "", rrn]]}
    return (
        service.spreadsheets()
        .values()
        .update(
            spreadsheetId=spreadsheet_id,
            range=range_a1,
            valueInputOption="RAW",
            body=body,
        )
        .execute()
    )
