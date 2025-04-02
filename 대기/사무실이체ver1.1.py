import subprocess
import pyautogui
import time as tm
import cv2
import numpy as np
import mss
import os   
import pandas as pd
import re          
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC



# 작업 디렉토리 설정                    
os.chdir("C:/Users/USER/Desktop/이체매크로")

print("Current working directory:", os.getcwd())    

# 이미지 파일 경로 설정
COORDS = {
    "localdisktab": "localdisk.png",
    "outdisktab": "seagate_bt.png",
    "commonlogin": "localdisk_loginbt.png",
    "send_tab": "send_tab.png",
    "multisend": "multisend_bt.png",
    "password4_input": "password4_input.png",
}

monitor_1 = {'top': 0, 'left': 1920, 'width': 1920, 'height': 1080}
monitor_2 = {'top': 0, 'left': 0, 'width': 1920, 'height': 1080}

def capture_screen(region):
    with mss.mss() as sct:   
        screenshot = sct.grab(region)
        img = np.array(screenshot)
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
        return img

def match_template(screen, template_path):
    template = cv2.imread(template_path, cv2.IMREAD_COLOR)
    if template is None:
        raise ValueError(f"이미지를 찾을 수 없습니다: {template_path}")

    result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(result)
    if max_val >= 0.95:  # 일치율이 95% 이상인 경우
        return (max_loc[0] + template.shape[1] // 2, max_loc[1] + template.shape[0] // 2)
    return None

def locate_image_on_monitor(image_path, monitor):
    screen = capture_screen(monitor)
    location = match_template(screen, image_path)
    if location:
        location = (location[0] + monitor['left'], location[1] + monitor['top']) 
    return location

def locate_image_on_monitors(image_path):
    location = locate_image_on_monitor(image_path, monitor_1)
    if not location:
        location = locate_image_on_monitor(image_path, monitor_2)       
    return location

def locate_and_click(image_path):
    print(f"{image_path}을(를) 찾고 클릭합니다.")
    location = locate_image_on_monitors(image_path)
    if location:
        print(f"{image_path}의 위치를 찾았습니다: {location}, 클릭합니다.")
        pyautogui.click(location)
        tm.sleep(1)
    else:
        print(f"{image_path}의 위치를 찾을 수 없습니다.")

def scroll_down():
    pyautogui.scroll(-10000)


# Selenium 설정 및 웹페이지 열기
url = 'https://www.kebhana.com/common/login.do'
options = Options()
options.add_experimental_option("detach", True)
options.add_argument("--disable-blink-features=AutomationControlled")

driver = webdriver.Chrome(options=options)
driver.get(url)

# WebDriverWait를 사용하여 요소가 나타날 때까지 대기 (최대 30초)
wait = WebDriverWait(driver, 30)
cert_login_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#certLogin")))
print(cert_login_element.text)

# 클릭 시도와 재시도 로직
max_attempts = 5  # 최대 시도 횟수
for attempt in range(max_attempts):
    try:
        cert_login_element.click()
        print("공동인증서 로그인 버튼 클릭 성공")
        break  # 성공하면 반복문 종료
    except Exception as e:
        if attempt < max_attempts - 1:  # 마지막 시도가 아니면
            print(f"클릭 실패, 1초 후 재시도합니다. ({attempt+1}/{max_attempts})")
            tm.sleep(1)  # 1초 대기 후 재시도
        else:
            print(f"최대 시도 횟수 초과: {e}")
            raise  # 재시도 모두 실패 시 예외 발생

tm.sleep(10)  # 다음 단계를 위한 대기시간

# 이미지가 나타날 때까지 기다리는 함수
def wait_for_image(image_path, max_wait=10):
    print(f"{image_path}이(가) 나타날 때까지 최대 {max_wait}초 대기합니다.")
    start_time = tm.time()
    while tm.time() - start_time < max_wait:
        location = locate_image_on_monitors(image_path)
        if location:
            print(f"{image_path}의 위치를 찾았습니다: {location}")
            return location
        tm.sleep(1)  # 1초마다 확인
    print(f"{image_path}를 {max_wait}초 내에 찾지 못했습니다.")
    return None

# PyAutoGUI를 이용한 인증서 로그인
def login_with_certificate():
    # localdisktab이 나타날 때까지 기다린 후 클릭
    location = wait_for_image(COORDS["localdisktab"])
    if location:
        pyautogui.click(location)
        print(f"{COORDS['localdisktab']}을(를) 클릭했습니다.")
    else:
        print(f"{COORDS['localdisktab']}를 찾지 못했습니다.")
    
    # outdisktab이 나타날 때까지 기다린 후 클릭 (최대 3초)
    location = wait_for_image(COORDS["outdisktab"], max_wait=3)
    if location:
        pyautogui.click(location)
        print(f"{COORDS['outdisktab']}을(를) 클릭했습니다.")
    else:
        print(f"{COORDS['outdisktab']}를 찾지 못했습니다.")
        # 찾지 못한 경우에도 계속 진행
        
    tm.sleep(3)  # 이 대기 시간은 유지 (다음 단계를 위한 대기)
    
    for _ in range(6):   # 탭 6번
        pyautogui.press('tab')
        tm.sleep(0.15)
    
    # 비밀번호 입력 (한글자씩 입력)
    password = "@gusqls120"
    for char in password:
        pyautogui.press(char)
        tm.sleep(0.07)
    
    locate_and_click(COORDS["commonlogin"])    # 공동로그인 버튼 클릭

login_with_certificate()
tm.sleep(8)  # 로그인 완료 후 8초 대기

# 프레임 전환
frame_id = "hanaMainframe"  # 프레임 ID를 여기에 입력하세요
driver.switch_to.frame(frame_id)

# 이체 탭 클릭
locate_and_click(COORDS["send_tab"])
tm.sleep(1)

# 다계좌 이체 버튼 클릭
for _ in range(6):   # 탭 3번
    pyautogui.press('tab')

pyautogui.press('enter')   
tm.sleep(3)

def scroll_down(amount):
    pyautogui.scroll(-amount)  # 스크롤 다운

def scroll_up(amount):
    pyautogui.scroll(amount)  # 스크롤 업

# 예시: 전체 페이지를 스크롤 다운하고 절반만 다시 스크롤 업
scroll_amount = 1000  # 페이지의 길이에 맞게 조정
scroll_down(scroll_amount)
tm.sleep(1)  # 잠시 대기
scroll_up(scroll_amount // 2)

# 은행 선택을 위한 맵핑
bank_options = {
    "하나은행": "081",
    "경남은행": "039",
    "광주은행": "034",
    "국민은행": "004",
    "기업은행": "003",
    "농협": "011",
    "iM뱅크(대구)": "031",
    "도이치뱅크": "055",
    "부산은행": "032",
    "산업은행": "002",
    "저축은행": "050",
    "새마을금고": "045",
    "수협은행": "007",
    "신협": "048",
    "신한은행": "088",
    "우리은행": "020",
    "우체국": "071",
    "전북은행": "037",
    "제주은행": "035",
    "카카오뱅크": "090",
    "케이뱅크": "089",
    "한국씨티은행": "027", 
    "BOA": "060",
    "HSBC": "054",
    "JP모간": "057",
    "SC제일은행": "023",
    "하나증권": "270",
    "교보증권": "261",
    "대신증권": "267",
    "미래에셋증권": "238",
    "DB금융투자": "279",        
    "유안타증권": "209",
    "메리츠증권": "287",
    "부국증권": "290",
    "삼성증권": "240",
    "신영증권": "291",
    "신한투자증권": "278",
    "NH투자증권": "247",
    "유진증권": "280",
    "키움증권": "264",
    "하이투자증권": "262",
    "한국투자": "243",
    "한화투자증권": "269",
    "KB증권": "218",
    "LS증권": "265",
    "현대차증권": "263",
    "케이프증권": "292",
    "SK증권": "266",
    "산림조합": "064",
    "중국공상은행": "062",
    "중국은행": "063",
    "중국건설은행": "067",
    "BNP파리바은행": "061",
    "한국포스증권": "294",
    "다올투자증권": "227",                      
    "BNK투자증권": "224",
    "카카오페이증권": "288",
    "IBK투자증권": "225",
    "토스증권": "271",
    "토스뱅크": "092",
    "상상인증권": "221"
}

# 은행명 표준화 함수
def standardize_bank_name(bank_name):
    bank_name = bank_name.lower().strip()
    if "sc제일" in bank_name or "제일" in bank_name:
        return "SC제일은행"
    elif "제일은행" in bank_name:
        return "SC제일은행"
    elif "하나" in bank_name:
        return "하나은행"
    elif "경남" in bank_name:
        return "경남은행"
    elif "광주" in bank_name:
        return "광주은행"
    elif "국민" in bank_name:
        return "국민은행"
    elif "기업" in bank_name:
        return "기업은행"
    elif "농협은행" in bank_name:
        return "농협"
    elif "NH농협" in bank_name:
        return "농협"
    elif "nh농협" in bank_name:
        return "농협"
    elif "대구" in bank_name:          
        return "iM뱅크(대구)"
    elif "im뱅크" in bank_name:
        return "iM뱅크(대구)"
    elif "대구은행" in bank_name:
        return "iM뱅크(대구)"
    elif "부산" in bank_name:
        return "부산은행"
    elif "새마을" in bank_name:
        return "새마을금고"
    elif "수협" in bank_name:
        return "수협은행"
    elif "신한" in bank_name:
        return "신한은행"
    elif "우리" in bank_name:
        return "우리은행"
    elif "전북" in bank_name:
        return "전북은행"
    elif "제주" in bank_name:
        return "제주은행"
    elif "카카오" in bank_name:
        return "카카오뱅크"
    elif "카뱅" in bank_name:
        return "카카오뱅크"
    elif "씨티" in bank_name:
        return "한국씨티은행"
    elif "토스" in bank_name:
        return "토스뱅크"
    elif "미래에셋대우" in bank_name:
        return "미래에셋증권"
       
    
    # 필요한 경우 추가 은행명 표준화
    return bank_name



# 전처리 함수
def preprocess_account_info(account_info):
    # NaN 값이나 문자열이 아닌 경우 처리
    if pd.isna(account_info) or not isinstance(account_info, str):
        return "", ""
        
    # 은행명과 계좌번호 분리
    match = re.match(r'(\D+)\s*([\d\s\-]+)', str(account_info))
    if match:
        bank_name = standardize_bank_name(match.group(1).strip())
        account_number = re.sub(r'[\s\-]', '', match.group(2))
    else:
        bank_name = ""
        account_number = ""
    
    return bank_name, account_number

# 엑셀 파일 로드
excel_path = "C:/Users/USER/Desktop/이체매크로/이체정보.xlsx"  # 엑셀 파일 경로를 지정하세요
df = pd.read_excel(excel_path)

# 전처리된 데이터를 저장할 리스트
processed_data = []

for index, row in df.iterrows():
    product_name = row.iloc[3]  # 제품명
    customer_name = row.iloc[4]  # 이름
    account_info = row.iloc[7]  # 은행+계좌번호
    amount = row.iloc[9]  # 금액

    # 계좌정보 전처리
    bank_name, account_number = preprocess_account_info(account_info)

    # 항목 1 : 은행
    # 항목 2 : 계좌번호             
    # 항목 3 : 이름.제품명
    # 항목 4 : 제품명
    name_product = f"{customer_name}{product_name}"
    processed_data.append((bank_name, account_number, name_product, product_name, amount))

# 전처리된 데이터 출력
for data in processed_data:
    print(f"은행: {data[0]}, 계좌번호: {data[1]}, 이름.제품명: {data[2]}, 제품명: {data[3]}, 금액: {data[4]}")

def type_text(element, text):
    element.clear()
    element.send_keys(text)

def type_number(text):
    for char in str(text):
        pyautogui.press(char)
        tm.sleep(0.005)

def input_transfer_info(data, index):
    bank, account_number, name_product, product_name, amount = data
    
    try:
        # 은행 선택
        bank_element = driver.find_element(By.ID, f"rcvBnkCd{index}")
        option_value = bank_options.get(bank, "")
        if option_value:
            option = bank_element.find_element(By.CSS_SELECTOR, f"option[value='{option_value}']")
            option.click()
            print(f"{name_product} 은행 선택 성공: {bank}")
        else:
            print(f"{name_product} 은행을 찾을 수 없습니다: {bank}")
        
        # 계좌번호 입력
        account_element = driver.find_element(By.ID, f"rcvAcctNo{index}")
        account_element.click()
        tm.sleep(0.1)  # 첫 번째 문자 입력 전에 짧은 지연을 추가
        type_number(account_number)
        print(f"{name_product} 계좌번호 입력 성공: {account_number}")
        
        # 금액 입력
        amount_element = driver.find_element(By.ID, f"trnsAmt{index}")
        amount_element.click()
        tm.sleep(0.1)  # 첫 번째 문자 입력 전에 짧은 지연을 추가
        type_number(int(amount))
        print(f"{name_product} 금액 입력 성공: {amount}")
        
        # 이름.제품명 입력
        name_product_element = driver.find_element(By.ID, f"wdrwPsbkMarkCtt{index}")
        type_text(name_product_element, name_product)
        print(f"{name_product} 이름.제품명 입력 성공: {name_product}")
        
        # 제품명 입력
        product_name_element = driver.find_element(By.ID, f"rcvPsbkMarkCtt{index}")
        type_text(product_name_element, product_name)
        print(f"{name_product} 제품명 입력 성공: {product_name}")

    except Exception as e:
        print(f"오류 발생 {name_product}: {e}")

# # paymAcctPw에 '5800' 입력
# def enter_password():
#     try:
#         password_element = driver.find_element(By.ID, "paymAcctPw")
#         password_element.click()
#         tm.sleep(0.1)  # 첫 번째 문자 입력 전에 짧은 지연을 추가
#         type_number("5800")
#         print("비밀번호 입력 성공")
#     except Exception as e:
#         print(f"비밀번호 입력 오류: {e}")

# 비밀번호 입력
# enter_password()

# 최대 10개의 항목 입력
for index, data in enumerate(processed_data):
    if index >= 10:
        break
    input_transfer_info(data, index)
    tm.sleep(0.5)  # 각 세트 완료 후 잠시 대기

# 나머지 코드는 필요에 따라 추가        

## 1.2ver 완성