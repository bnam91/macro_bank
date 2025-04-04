import webbrowser
import requests
from bs4 import BeautifulSoup
import time

def open_url_in_browser(url):
    # URL을 기본 웹 브라우저에서 엽니다.
    webbrowser.open(url)

def parse_html_and_save_to_txt(url, filename="output.txt"):
    # URL의 HTML을 가져옵니다.
    response = requests.get(url)
    response.raise_for_status()  # 요청이 성공했는지 확인

    # 인코딩 설정
    response.encoding = response.apparent_encoding

    # HTML 소스 코드를 파일에 저장합니다.
    with open(filename, 'w', encoding='utf-8') as file:
        file.write(response.text)
    
    print(f"HTML 소스 코드가 {filename}에 저장되었습니다.")

if __name__ == "__main__":
    # 사용자로부터 URL 입력 받기
    url = input("이동할 URL을 입력하세요: ")
    open_url_in_browser(url)
    
    # 3초 카운트다운
    for i in range(3, 0, -1):
        print(f"{i}초 남았습니다...")
        time.sleep(1)
    
    parse_html_and_save_to_txt(url)
