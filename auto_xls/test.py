from selenium import webdriver  # selenium 프레임 워크에서 webdriver 가져오기
url = 'https://tumblbug.com/'        # 접속할 웹 사이트 주소
d_location=str(input("크롬 드라이버의 위치를 입력해주세요(ex: C:\chromedriver.exe):"))
driver = webdriver.Chrome(d_location)
driver.get(url)  # 저장한 url 주소로 이동