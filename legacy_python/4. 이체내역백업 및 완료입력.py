from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime
import sys
import time
import os

API_KEY_MODULE_DIR = os.path.join(os.path.expanduser('~'), 'Documents', 'github_cloud', 'module_auth')
if API_KEY_MODULE_DIR not in sys.path:
    sys.path.insert(0, API_KEY_MODULE_DIR)

from auth import get_credentials

SPREADSHEET_ID = '1CK2UXTy7HKjBe2T0ovm5hfzAAKZxZAR_ev3cbTPOMPs'

def update_sheets():
    try:
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)

        sheet_metadata = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        sheets = sheet_metadata.get('sheets', '')

        sheet_map = {sheet['properties']['title']: sheet['properties']['sheetId'] for sheet in sheets}

        for sheet_name, sheet_id in sheet_map.items():
            # '완료'가 포함된 시트명은 제외
            if '완료' not in sheet_name:
                process_sheet(service, sheet_id, sheet_name)
            else:
                print(f"'{sheet_name}' 시트는 '완료'가 포함되어 제외됩니다.")

    except HttpError as error:
        print(f'오류가 발생했습니다: {error}')
        if error.resp.status == 502:
            print("서버 오류가 발생했습니다. 30초 후에 다시 시도합니다.")
            time.sleep(30)
            update_sheets()  # 재시도

def process_sheet(service, sheet_id, sheet_name):
    print(f"Processing sheet: {sheet_name} (ID: {sheet_id})")
    
    range_name = f'{sheet_name}!A1:P1000'  # P열까지만 범위 지정
    
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
    values = result.get('values', [])

    if not values:
        print(f'{sheet_name}에서 데이터를 찾을 수 없습니다.')
        return

    print(f"{sheet_name}의 총 행 수: {len(values)}")

    requests = []
    for row_index, row in enumerate(values, start=1):
        if len(row) >= 16:  # P열이 존재하는 경우
            status = row[15].strip() if len(row) > 15 else ''
            if status.lower() in ['입금요청', '입금오류']:
                if status.lower() == '입금요청':
                    requests.append(update_row_color(sheet_id, row_index, '진한 회색 1'))
                    requests.append(update_cell_value(sheet_id, row_index, 15, get_deposit_complete_text()))
                    print(f"Row {row_index}: 입금요청 처리")
                elif status.lower() == '입금오류':
                    requests.append(update_row_color(sheet_id, row_index, '주황색'))
                    print(f"Row {row_index}: 입금오류 처리")

    # 원본 시트 업데이트
    if requests:
        body = {'requests': requests}
        response = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
        print(f'{sheet_name} 업데이트 완료 ({len(requests) // 2} 행 처리됨)')
    else:
        print(f'{sheet_name}에서 업데이트할 내용이 없습니다.')

def update_row_color(sheet_id, row_index, color):
    return {
        'updateCells': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': row_index - 1,
                'endRowIndex': row_index,
                'startColumnIndex': 1,  # B열부터 시작
                'endColumnIndex': 16  # P열까지
            },
            'rows': [{
                'values': [{'userEnteredFormat': {'backgroundColor': get_color(color)}}] * 15
            }],
            'fields': 'userEnteredFormat.backgroundColor'
        }
    }

def update_cell_value(sheet_id, row_index, column_index, new_value):
    return {
        'updateCells': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': row_index - 1,
                'endRowIndex': row_index,
                'startColumnIndex': column_index,
                'endColumnIndex': column_index + 1
            },
            'rows': [{
                'values': [{
                    'userEnteredValue': {'stringValue': new_value}
                }]
            }],
            'fields': 'userEnteredValue.stringValue'
        }
    }

def get_color(color_name):
    colors = {
        '진한 회색 1': {'red': 0.8, 'green': 0.8, 'blue': 0.8},  # 더 밝은 회색
        '주황색': {'red': 1.0, 'green': 0.6, 'blue': 0.0}
    }
    return colors.get(color_name, {'red': 1, 'green': 1, 'blue': 1})

def get_deposit_complete_text():
    today = datetime.now().strftime('%y%m%d')
    return f'입금완료_{today}'

if __name__ == '__main__':
    update_sheets()