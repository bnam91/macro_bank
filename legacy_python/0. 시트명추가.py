import sys
from pathlib import Path

# 외부 auth 모듈 경로 추가
AUTH_MODULE_DIR = Path("/Users/a1/Documents/github_cloud/module_auth")
if str(AUTH_MODULE_DIR) not in sys.path:
    sys.path.append(str(AUTH_MODULE_DIR))

from auth import get_credentials
from googleapiclient.discovery import build
import logging

def get_sheet_names(spreadsheet_id):
    """스프레드시트의 모든 시트명을 가져오는 함수"""
    logging.info(f"스프레드시트에서 시트명 가져오기 시작")
    try:
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)
        
        # 스프레드시트 정보 가져오기
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        
        # 시트 정보에서 시트명만 추출
        sheets = spreadsheet.get('sheets', [])
        sheet_names = [sheet['properties']['title'] for sheet in sheets]
        
        # '완료'가 포함된 시트명 제외
        filtered_sheet_names = [name for name in sheet_names if '완료' not in name]
        
        logging.info(f"시트명 가져오기 성공: {len(filtered_sheet_names)} 시트")
        return filtered_sheet_names
    
    except Exception as e:
        logging.error(f"시트명 가져오기 실패: {str(e)}")
        raise

def create_query_formula(spreadsheet_id, sheet_names):
    """시트명을 이용해 QUERY 식 생성"""
    formula = "=QUERY({"
    
    for i, name in enumerate(sheet_names):
        # 시트명에 공백이나 특수문자가 있는 경우 작은따옴표로 감싸기
        if ' ' in name or '(' in name or ')' in name:
            formatted_name = f"'{name}'"
        else:
            formatted_name = name
            
        import_line = f'IMPORTRANGE("https://docs.google.com/spreadsheets/d/{spreadsheet_id}", "{formatted_name}!A1:P1000")'
        
        # 마지막 줄이 아니면 세미콜론 추가
        if i < len(sheet_names) - 1:
            import_line += ";"
        
        formula += import_line
    
    formula += '}, "select * where Col16 = \'입금요청\'", 0)'
    return formula

def write_to_spreadsheet(target_spreadsheet_id, formula):
    """스프레드시트의 A2 셀에 생성된 식 입력"""
    logging.info(f"스프레드시트 {target_spreadsheet_id}에 식 입력 시작")
    try:
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)
        
        # A2 셀에 쓰기
        body = {
            'values': [[formula]]
        }
        
        result = service.spreadsheets().values().update(
            spreadsheetId=target_spreadsheet_id,
            range='A2',
            valueInputOption='USER_ENTERED',  # 'USER_ENTERED'로 설정하여 수식으로 인식되도록 함
            body=body
        ).execute()
        
        logging.info(f"식 입력 성공: {result.get('updatedCells')} 셀 업데이트됨")
        return result
    
    except Exception as e:
        logging.error(f"식 입력 실패: {str(e)}")
        raise

def main():
    # 소스 스프레드시트 ID (시트명을 가져올 스프레드시트)
    SOURCE_SPREADSHEET_ID = '1CK2UXTy7HKjBe2T0ovm5hfzAAKZxZAR_ev3cbTPOMPs'
    
    # 타겟 스프레드시트 ID (식을 쓸 스프레드시트)
    TARGET_SPREADSHEET_ID = '1NOP5_s0gNUCWaGIgMo5WZmtqBbok_5a4XdpNVwu8n5c'
    
    # 로깅 설정
    logging.basicConfig(level=logging.INFO)
    
    # 시트명 가져오기
    sheet_names = get_sheet_names(SOURCE_SPREADSHEET_ID)
    
    # QUERY 식 생성
    formula = create_query_formula(SOURCE_SPREADSHEET_ID, sheet_names)
    
    # 생성된 식 출력 (확인용)
    print("생성된 QUERY 식:")
    print(formula)
    
    # 타겟 스프레드시트에 식 쓰기
    write_to_spreadsheet(TARGET_SPREADSHEET_ID, formula)
    print(f"QUERY 식이 스프레드시트 {TARGET_SPREADSHEET_ID}의 A2 셀에 성공적으로 입력되었습니다.")

if __name__ == "__main__":
    main()