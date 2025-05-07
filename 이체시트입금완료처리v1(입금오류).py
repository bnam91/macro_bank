from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from auth import get_credentials
from datetime import datetime
import time

SPREADSHEET_ID = '1CK2UXTy7HKjBe2T0ovm5hfzAAKZxZAR_ev3cbTPOMPs'
BACKUP_SPREADSHEET_ID = '12Ivr6aKhl585Y6889TVnbrSY-Qgevu9qri0cI4-y84s'
BACKUP_SHEET_NAME = 'BackupData'  # 백업 스프레드시트의 시트 이름을 지정

def update_sheets():
    try:
        # 쉼표로 구분된 이름들을 입력받아 리스트로 변환
        names_input = input("입금오류인 인원들을 쉼표(,)로 구분하여 입력하세요: ")
        target_names = [name.strip() for name in names_input.split(',')]
        
        print(f"처리할 인원: {', '.join(target_names)}")
        
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)

        sheet_metadata = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        sheets = sheet_metadata.get('sheets', '')

        sheet_map = {sheet['properties']['title']: sheet['properties']['sheetId'] for sheet in sheets}

        # 먼저 각 이름이 어떤 시트에 있는지 찾기
        name_locations = find_name_locations(service, sheet_map, target_names)
        
        # 각 이름별로 위치 정보 출력 및 확인 후 처리
        for name, locations in name_locations.items():
            if locations:
                print(f"\n'{name}'님은 다음 시트에서 발견되었습니다:")
                for idx, loc in enumerate(locations, 1):
                    print(f"{idx}. {loc['sheet_name']} (행: {loc['row_index']})")
                
                if len(locations) > 1:
                    exclude_input = input(f"'{name}'님 처리에서 제외할 시트 이름을 쉼표(,)로 구분하여 입력하세요 (모두 처리: y, 모두 건너뛰기: n): ")
                    
                    if exclude_input.lower() == 'n':
                        print(f"'{name}'님 처리를 건너뛰었습니다.")
                        continue
                    elif exclude_input.lower() == 'y':
                        # 모든 시트 처리
                        process_name(service, name, locations)
                    else:
                        # 특정 시트 제외
                        exclude_sheets = [sheet.strip() for sheet in exclude_input.split(',')]
                        filtered_locations = [loc for loc in locations if loc['sheet_name'] not in exclude_sheets]
                        
                        if filtered_locations:
                            process_name(service, name, filtered_locations)
                            print(f"'{name}'님 처리 완료 (제외된 시트: {', '.join(exclude_sheets)})")
                        else:
                            print(f"모든 시트가 제외되어 '{name}'님은 처리되지 않았습니다.")
                else:
                    # 시트가 1개인 경우 바로 처리
                    process_name(service, name, locations)
                    print(f"'{name}'님을 자동으로 입금오류 처리했습니다.")
            else:
                print(f"\n'{name}'님은 '입금요청' 상태로 어떤 시트에서도 발견되지 않았습니다.")

    except HttpError as error:
        print(f'오류가 발생했습니다: {error}')
        if error.resp.status == 502:
            print("서버 오류가 발생했습니다. 30초 후에 다시 시도합니다.")
            time.sleep(30)
            update_sheets()

def find_name_locations(service, sheet_map, target_names):
    name_locations = {name: [] for name in target_names}
    
    for sheet_name, sheet_id in sheet_map.items():
        print(f"시트 '{sheet_name}' 검사 중...")
        range_name = f'{sheet_name}!A1:P1000'
        
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
        values = result.get('values', [])

        if not values:
            print(f'{sheet_name}에서 데이터를 찾을 수 없습니다.')
            continue

        for row_index, row in enumerate(values, start=1):
            if len(row) >= 16:  # P열이 존재하는 경우
                status = row[15].strip() if len(row) > 15 else ''
                name = row[5].strip() if len(row) > 5 else ''  # F열의 이름
                
                if status.lower() == '입금요청' and name in target_names:
                    name_locations[name].append({
                        'sheet_id': sheet_id,
                        'sheet_name': sheet_name,
                        'row_index': row_index,
                        'row_data': row
                    })
    
    return name_locations

def process_name(service, name, locations):
    requests = []
    backup_data = []
    
    for loc in locations:
        backup_data.append(loc['row_data'])
        sheet_id = loc['sheet_id']
        row_index = loc['row_index']
        sheet_name = loc['sheet_name']
        
        requests.append(update_row_color(sheet_id, row_index, '주황색'))
        requests.append(update_cell_value(sheet_id, row_index, 15, '입금오류'))
        print(f"시트 '{sheet_name}', Row {row_index}: {name}님 입금오류 처리")
    
    # 백업 먼저 수행
    if backup_data:
        backup_sheet(service, backup_data, locations[0]['sheet_name'])
    
    # 원본 시트 업데이트
    if requests:
        body = {'requests': requests}
        response = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
        print(f'{name}님 업데이트 완료 ({len(requests) // 2} 행 처리됨)')
    else:
        print(f'{name}님의 업데이트할 내용이 없습니다.')

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

def backup_sheet(service, backup_data, original_sheet_name):
    try:
        # 백업 스프레드시트의 시트 확인 및 생성
        sheet_metadata = service.spreadsheets().get(spreadsheetId=BACKUP_SPREADSHEET_ID).execute()
        sheets = sheet_metadata.get('sheets', '')
        sheet_names = [sheet['properties']['title'] for sheet in sheets]
        
        if BACKUP_SHEET_NAME not in sheet_names:
            add_sheet_request = {
                'addSheet': {
                    'properties': {
                        'title': BACKUP_SHEET_NAME
                    }
                }
            }
            service.spreadsheets().batchUpdate(
                spreadsheetId=BACKUP_SPREADSHEET_ID,
                body={'requests': [add_sheet_request]}
            ).execute()
            print(f"'{BACKUP_SHEET_NAME}' 시트가 생성되었습니다.")

        # 백업 스프레드시트의 마지막 행 찾기
        result = service.spreadsheets().values().get(
            spreadsheetId=BACKUP_SPREADSHEET_ID, range=f'{BACKUP_SHEET_NAME}!A:A').execute()
        values = result.get('values', [])
        next_row = len(values) + 1

        # 백업 데이터에 원본 시트 이름과 타임스탬프 추가
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        backup_data = [[original_sheet_name, timestamp] + row for row in backup_data]

        # 백업 데이터 추가
        range_name = f'{BACKUP_SHEET_NAME}!A{next_row}'
        body = {'values': backup_data}
        result = service.spreadsheets().values().append(
            spreadsheetId=BACKUP_SPREADSHEET_ID, range=range_name,
            valueInputOption='USER_ENTERED', body=body).execute()
        
        print(f'{len(backup_data)} 행이 백업되었습니다.')
    except HttpError as error:
        print(f'백업 중 오류 발생: {error}')
        print("백업에 실패했지만, 원본 시트 업데이트는 계속 진행합니다.")

if __name__ == '__main__':
    update_sheets()