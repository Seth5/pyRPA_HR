from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import requests
from bs4 import BeautifulSoup
import time
import math
import csv
from openpyxl import Workbook

interval = 2 # 페이지 전환시 2초 대기

# 크롬 브라우저 열기
browser = webdriver.Chrome()
# browser.maximize_window()

# 더존 그룹웨어로 이동
browser.get("http://54.180.21.232/gw/userMain.do")

# ID 및 PASSWD 입력
browser.find_element_by_xpath("//*[@id='userId']").send_keys("seth.oh")
elem = browser.find_element_by_xpath("//*[@id='userPw']")
elem.send_keys("toangel2@")
elem.send_keys(Keys.ENTER)
# 페이지 로딩 대기
time.sleep(interval)

# 마스터 버튼 클릭
browser.find_element_by_xpath("/html/body/div[2]/div/div/div[2]/div[4]/a").click()
# 페이지 로딩 대기
time.sleep(interval)

# 전자결재 버튼
browser.find_element_by_xpath("//*[@id='topMenu1700000000']/div/div[2]").click()
# 페이지 로딩 대기
time.sleep(interval)

# 결재문서관리 클릭
browser.find_element_by_xpath("//*[@id='4dep']").click()
# 페이지 로딩 대기
time.sleep(interval)

# 결재문서목록 클릭
browser.find_element_by_xpath("//*[@id='1705020000_anchor']").click()
# 페이지 로딩 대기
time.sleep(interval)

# iframe 선택
browser.switch_to.frame("_content")

# iframe 전체 이름 확인하기
# iframes = browser.find_elements_by_css_selector("iframe")
# for iframe in iframes:
#         print(iframe.get_attribute("name"))

# 상세검색 클릭
browser.find_element_by_xpath("/html/body/div[1]/div[2]/span").click()
# 페이지 로딩 대기
time.sleep(interval)

# 문서분류 입력
elem = browser.find_element_by_xpath("//*[@id='txt_form_nm']")
elem.send_keys("연장")
elem.send_keys(Keys.ENTER)
# 페이지 로딩 대기
time.sleep(interval)

req = browser.page_source
soup = BeautifulSoup(browser.page_source, "lxml")
# # docs = soup.find_all("td", "pl5 text_ho")
# data_rows = soup.find("div", attrs={"class":"grid-wrap"}).find("tbody").find_all("tr")
# # print(data_rows) # ("td", attrs={"class":"number"})
# for row in data_rows:
#     columns = row.find_all("td")
#     if len(columns) <= 1:  # 의미없는 데이터 건너뛰기
#         continue
#     data = [column.get_text().strip().replace("\n","").replace("\xa0","").replace("\t","") for column in columns]
#     print(data)

page_info = soup.find("div", attrs={"class":"page_info"}).find_all("span")
# print(page_info[1].get_text()) # 총 문서
num_of_docs = int(page_info[1].get_text())
max_loop = math.ceil(num_of_docs / 10) # 총 문서수가 10을 초과하면 다중 루프가 필요
# print(max_loop)

filename = "연장-20210127.csv"
f = open(filename, "w", encoding="utf-8-sig", newline="")
writer = csv.writer(f)

wb = Workbook() # 새 워크북 생성
ws = wb.active # 활성화된 쉬트를 가져옴
# ws.tile = "TestSheet" # 쉬트의 이름 변경

for loop in range(1, max_loop + 1):
    if num_of_docs <= 10:
        n = num_of_docs + 1
    else: # 10 보다 큰 경우는 2번 이상의 루프를 돈다.
        if loop == max_loop: # 마지막 루프
            n = (num_of_docs % 10) + 1 # 36 % 10 = 6(나머지)
        else:
            n = 10 + 1

    for i in range(1,n):
        id = "//*[@id='grid']/div[1]/div[2]/table/tbody/tr[" + str(i) + "]/td[4]"
        # print(id)
        # 제목 컬럼에서 n번째 항목 클릭
        browser.find_element_by_xpath(id).click()
        time.sleep(interval)
        allhandles = browser.window_handles
        browser.switch_to.window(allhandles[1])

        req = browser.page_source
        soup = BeautifulSoup(browser.page_source, "lxml")

        # tables = soup.find_all("table")
        # # print(len(tables))
        # data_rows = tables[3].find("tbody").find_all("tr")
        tables = soup.find_all("tbody")
        data_rows = tables[3].find_all("tr")
        for row in data_rows:
            # columns = row.find_all("td")
            columns = row.select("td")
            if len(columns) <= 1:  # 의미없는 데이터 건너뛰기
                continue
            data = [column.get_text().strip().replace("\n","").replace("\xa0","").replace("\t","").replace("\u200b","") for column in columns]
            # data = [columns[1].get_text().strip()]
            if len(data) > 5:  # 의미없는 데이터 건너뛰기
                continue
            if data[0] == "" or data[0] == "수신 및 참조" or data[0] == "시  행" or data[0] == "연장/휴일 근무목적":  # 의미없는 데이터 건너뛰기
                continue
            print(data)
            writer.writerow(data)
            ws.append(data)
        print("=" * 100)
        writer.writerow(["'**************************","'**************************","'**************************","'**************************"])
        ws.append(["'**************************","'**************************","'**************************","'**************************"])

        browser.close()
        time.sleep(interval)

        browser.switch_to.window(allhandles[0])
        time.sleep(interval)

        browser.switch_to.frame("_content")

    if loop < max_loop:
        # 다음 버튼 클릭
        browser.find_element_by_xpath('//*[@id="grid"]/div[3]/div[1]/span[3]/a').click()
        time.sleep(interval)

browser.quit()
wb.save("연장-20210127.xlsx")
wb.close()

