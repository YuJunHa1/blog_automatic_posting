from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from fake_useragent import UserAgent
from openpyxl import Workbook, load_workbook
import time
import random
import os
import pyperclip

def chrome_setting():
    chrome_version = random.randint(118, 122)# 고정된 모바일 버전의 User-Agent 패턴 (숫자만 랜덤하게 변경)
    user_agent = f"Mozilla/5.0 (iPhone; CPU iPhone OS 17_3 like Mac OS X) AppleWebKit/537.36 (KHTML, like Gecko) Version/17.0 Mobile/{chrome_version}.0.0 Safari/537.36"
    options = webdriver.ChromeOptions()
    options.add_argument("accept-language=ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7") #한국어(ko-KR)**를 우선적으로 사용하고, 그다음으로 **영어(en-US, en)**를 선호한다는 것을 나타냅니다.
    options.add_argument("--disable-blink-features=AutomationControlled") #크롬의 셀레니움 자동화 탐지 비활성화
    options.add_argument(f"user-agent={user_agent}") #가짜 User-Agent를 생성하여 쿠팡이 봇을 감지하는 것을 방지
    options.add_experimental_option("excludeSwitches", ["enable-automation"]) #브라우저가 자동화된 환경에서 실행되고 있다는 표시를 숨김
    options.add_experimental_option("useAutomationExtension", False) #브라우저가 자동화된 환경에서 실행되고 있다는 표시를 숨김
    options.add_experimental_option("prefs", {"profile.managed_default_content_setting.images": 2}) #브라우저의 이미지 로딩 차단, 속도 개선
    options.add_experimental_option("detach", True)#창이 닫히지 않고 유지
    driver = webdriver.Chrome(options=options)# 크롬 드라이버 실행
    # 크롤링 방지 탐지 우회
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": "Object.defineProperty(navigator, 'webdriver', { get: () => undefined })"
    })
    return driver



# 계정 정보 및 검색어 설정
ID = "qetu5702@gmail.com"
PW = "dbwnsgk7575*"
search_word = "등산복"
login_page = 'https://login.coupang.com/login/login.pang?rtnUrl=https%3A%2F%2Fpartners.coupang.com%2Fapi%2Fv1%2Fpostlogin'

def login(driver):
    #아이디 넣기
    driver.get(login_page)
    login_box = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "login-email-input")))
    login_box.click()
    pyperclip.copy(ID)
    login_box.send_keys(Keys.CONTROL, "v")
    #비밀번호 넣기
    pw_box = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "login-password-input")))
    pw_box.click()
    pyperclip.copy(PW)
    pw_box.send_keys(Keys.CONTROL, "v")
    #로그인 버튼 누르기, 실패시 계속 실행행
    while(True):
        driver.find_element(By.CSS_SELECTOR, ".login__button--submit").click()
        WebDriverWait(driver, 10).until(EC.url_contains("postlogin"))
        time.sleep(1)
        if not driver.find_elements(By.CSS_SELECTOR, ".login__button--submit"):
            break
    
def search(driver):
    search_box = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".ant-input.ant-input-lg")))
    search_box.click()
    pyperclip.copy(search_word)
    search_box.send_keys(Keys.CONTROL, "v")
    driver.find_element(By.CSS_SELECTOR, ".ant-btn.search-button.ant-btn-primary").click()
    #WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".ant-btn.hover-btn.btn-generate-link")))


# 반복문의 반복 횟수, 상위 n개의 상품 가져오기기
n = 10

def get_link(driver, n): #n개의 상품 링크 가져오기
    links = [] # 상품 링크를 담을 리스트
    action = ActionChains(driver)
    for i in range(n):
        #Xpath를 이용해 이번 차례 탐색할 상품 div위로 마우스 이동
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f"(//div[@class='product-item'])[{i+1}]")))
        action.move_to_element(driver.find_element(By.XPATH, f"(//div[@class='product-item'])[{i+1}]")).perform()
        #마우스가 이동해 링크생성 버튼이 나타나면 해당 버튼을 클릭
        link_button = driver.find_element(By.XPATH, f"(//button[contains(@class, 'btn-generate-link')])[{i+1}]")
        link_button.click()
        #링크 복사 후 리스트에 저장
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".ant-btn.lg.shorten-url-controls-main"))).click()
        clipboard_text = pyperclip.paste()
        links.append(clipboard_text)
        driver.back()
    print("#####################################################")
    print(links)
    print("#####################################################")
    print("끝!!!!")
    return links


#링크 타고 들어가서 대표 이미지와 베스트 리뷰 가져오기
def get_img_review(driver, links):
    title_imgs = []
    best_reviews = []
    for link in links:
        driver.get(link)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, '.rds-button.rds-button--lg.rds-button--fill-blue-normal.rds-button--block.main-product_action__XSlcT'))).click()
        best_review = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".ProductReview_review__article__reviews__text__article__FveJz"))
        )
        print(best_review.text)
        best_reviews.append(best_review.text)
    return best_reviews

    



#링크, 이미지 경로, 리뷰 엑셀에 저장하기
def save_xl(links, best_reviews):
    # 엑셀 파일 경로
    file_path = "C:\\blog_automate\\blog_automate.xlsx"

    if os.path.exists(file_path):     # 파일이 존재하는지 확인
        wb = load_workbook(file_path) # 파일이 존재하면 열기
    else:
        wb = Workbook() # 파일이 존재하지 않으면 새로 생성

    
    ws = wb.active
    #1행은 속성 값으로 채우기기
    ws.cell(row=1, column=1, value="item_link")
    ws.cell(row=1, column=2, value="img_path")
    ws.cell(row=1, column=3, value="best_review")

    #2행부터 각 속성에 맞는 값 채우기기
    for i in range(2, ws.max_row+1):
        ws.cell(row=i, column=1, value=links[i-2])
        #ws.cell(row=i, column=2, value=title_imgs[i-2])
        ws.cell(row=i, column=3, value=best_reviews[i-2])
    #저장
    wb.save("C:\\blog_automate\\blog_automate.xlsx")


if __name__ == "__main__":
    driver = chrome_setting()
    login(driver)
    search(driver)
    links = get_link(driver, 10)
    best_reviews = get_img_review(driver, links)
    save_xl(links, best_reviews)
    print("스크립트 종료")
