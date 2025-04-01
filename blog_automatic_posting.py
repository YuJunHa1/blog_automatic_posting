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
        time.sleep(1)
        clipboard_text = pyperclip.paste()
        if(clipboard_text[:4] == "http"):
            links.append(clipboard_text)
        driver.back()
    return links



#링크 타고 들어가서 대표 이미지와 베스트 리뷰 가져오기
def get_img_review_title(driver, links):
    img_srcs = []
    best_reviews = []
    titles = []
    for link in links:
        driver.get(link)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, '.rds-button.rds-button--lg.rds-button--fill-blue-normal.rds-button--block.main-product_action__XSlcT'))).click()
        #베스트 리뷰 가져오기기
        img_src = get_img(driver)
        img_srcs.append(img_src)
        best_review = get_review(driver)
        best_reviews.append(best_review)
        title = get_title(driver)
        titles.append(title)
        print(img_srcs)
        print(best_reviews)
        print(titles)
    return img_srcs, best_reviews, titles

def get_img(driver):
    
    try:
        img_src = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".rds-img img"))
        )
        return img_src.get_attribute("src")
    except Exception as e:
            print(f"오류 발생: {e}")
            return None

def get_review(driver):
    best_review = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, ".ProductReview_review__article__reviews__text__article__FveJz"))
    )
    return best_review.text

def get_title(driver):
    title = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, ".ProductInfo_title__fLscZ"))
    )
    return title.text

    



#링크, 이미지 경로, 리뷰 엑셀에 저장하기
def save_xl(links, img_srcs, best_reviews, titles):
    print("save_xl실행!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!1")
    # 엑셀 파일 경로
    file_path = "C:\\blog_automatic_posting\\blog_automatic_posting.xlsx"

    if os.path.exists(file_path):     # 파일이 존재하는지 확인
        wb = load_workbook(file_path) # 파일이 존재하면 열기
    else:
        wb = Workbook() # 파일이 존재하지 않으면 새로 생성

    
    ws = wb.active
    #1행은 속성 값으로 채우기기
    ws.cell(row=1, column=1, value="item_link")
    ws.cell(row=1, column=2, value="img_path")
    ws.cell(row=1, column=3, value="best_review")
    ws.cell(row=1, column=4, value="title")

    #2행부터 각 속성에 맞는 값 채우기
    for i in range(len(links)):
        ws.cell(row=i+2, column=1, value=links[i])
        ws.cell(row=i+2, column=2, value=img_srcs[i])
        ws.cell(row=i+2, column=3, value=best_reviews[i])
        ws.cell(row=i+2, column=4, value=titles[i])

    #중복값 제거
    seen = set()
    rows_to_delete = []

    # 2번째 행부터 데이터 확인 (1번째 행은 헤더)
    for row in range(4, ws.max_row + 1):
        titles = ws.cell(row=row, column=1).value  # item_link 컬럼 값 가져오기

        if titles in seen:  # 이미 존재하면 삭제 리스트에 추가
            rows_to_delete.append(row)
        else:
            seen.add(titles)  # 처음 등장하는 값은 저장

    # 중복 행 삭제 (뒤에서부터 삭제해야 인덱스가 틀어지지 않음)
    for row in reversed(rows_to_delete):
        ws.delete_rows(row)

    #저장
    wb.save(file_path)
    wb.close()


if __name__ == "__main__":
    driver = chrome_setting()
    login(driver)
    search(driver)
    links = get_link(driver, 10)
    img_srcs, best_reviews, titles = get_img_review_title(driver, links)
    save_xl(links, img_srcs, best_reviews, titles)
    print("스크립트 종료")
