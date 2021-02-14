from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
# import requests
# from bs4 import BeautifulSoup

interval = 2 # 페이지 전환시 2초 대기

# 크롬 브라우저 열기
browser = webdriver.Chrome()
# browser.maximize_window()

# 고용산재보험 사이트로 이동
browser.get("https://total.kcomwel.or.kr/")
time.sleep(interval)

# 자식 팝업 윈도우들이 많이 생성됨.
allhandles = browser.window_handles
# print(allhandles)
# print(browser.current_window_handle)
# for idx in range(1, len(allhandles)):
#     browser.switch_to.window(allhandles[idx])
#     browser.close()
#     time.sleep(1)
#
# allhandles = browser.window_handles
# print(allhandles)

browser.switch_to.window(allhandles[0]) # 메인 윈도우 선택

# 로그인 선택
browser.find_element_by_id('mf_wfm_header_btn_login').click()
time.sleep(interval)

# 사업장명의인증서 선택
browser.find_element_by_xpath('//*[@id="mf_wfm_content_u_group_fg1"]/span[3]/label').click()

# 사업자등록번호입력
elem = browser.find_element_by_id('mf_wfm_content_drno1_1')
elem.send_keys("144")
elem = browser.find_element_by_id('mf_wfm_content_drno1_2')
elem.send_keys("81")
elem = browser.find_element_by_id('mf_wfm_content_drno1_3')
elem.send_keys("03780")

# 공동인증서 로그인 선택
browser.find_element_by_id('mf_wfm_content_btn_login1').click()
time.sleep(interval)

# 공인인증서 로그인은 별도의 윈도우가 아님에 주의!
# 이동식디스크 선택
browser.find_element_by_xpath('//*[@id="xwup_media_removable"]').click() # 버튼 클릭
time.sleep(interval)
browser.find_element_by_xpath('//*[@id="xwup_location_4"]/div').click() # D: 드라이브 클릭
time.sleep(interval)

elem = browser.find_element_by_xpath('//*[@id="xwup_certselect_tek_input1"]')
elem.send_keys("eptmxls08!")
elem.send_keys(Keys.ENTER)
time.sleep(interval)
time.sleep(interval)

# 정보조회 선택
browser.find_element_by_id('mf_wfm_header_gen_firstGenerator_0_gen_SecondGenerator_1_get_second_li').click()
time.sleep(interval)
# 보험료정보 조회 선택
browser.find_element_by_xpath('//*[@id="mf_wfm_side_trv_menu_col_icon_navi_11"]/div').click()
time.sleep(interval)
# 부가고지 보험료 조회 선택
browser.find_element_by_id('mf_wfm_side_trv_menu_label_20').click()
time.sleep(interval)
# 고용관리/보수관리 비밀번호 선택
browser.find_element_by_id('mf_wfm_side_chongmuInjeung_wframe_wq_uuid_535').click()
time.sleep(interval)

# 고용관리/보수관리 비밀번호 입력
elem = browser.find_element_by_id('mf_wfm_side_chongmuInjeung_wframe_chongmuInjeung_check_wframe_CHANGE_PW')
elem.send_keys("destin08**")

# 조회 버튼 클릭
browser.find_element_by_id('mf_wfm_side_chongmuInjeung_wframe_chongmuInjeung_check_wframe_wq_uuid_558').click()
time.sleep(interval)

try:
    # 'browser'를 5초 기다리되 다음 조건이 나타나면 기다림을 중지하고 element를 가져옴.
    # 다음 조건은 Xpath '//*[@id="mf_wfm_content_alert794674078067883_wframe_btn_confirm"]'이 나타나는 것임
    # EC.presence_of_element_located()의 인자는 튜플이어야 함!
    elem = WebDriverWait(browser,5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[6]/div[2]/div[1]/div/div/div[2]/input[1]')))
    # 알림 창이 있으면 제거
    elem.click()
finally:
    # 알림 창이 없으면 pass
    pass

# 관리 번호 찾기
browser.find_element_by_id('mf_wfm_content_btnSaeopjangSearch').click()
time.sleep(interval)

# 선택 버튼 클릭
browser.find_element_by_xpath('//*[@id="mf_wfm_content_WZ0112_P01_wframe_GridMain_cell_0_0"]/button').click()
time.sleep(interval)

# 조회 버튼 클릭
browser.find_element_by_id('mf_wfm_content_wq_uuid_777').click()
time.sleep(interval)
time.sleep(interval)

# 당월보험료 부과내역조회 버튼 클릭(산재)
browser.find_element_by_id('mf_wfm_content_wq_uuid_1684').click()
time.sleep(interval)
time.sleep(interval)

# 엑셀저장 버튼 클릭(산재)
browser.find_element_by_id('mf_wfm_content_WL0502_P04_wframe_wq_uuid_1772').click()
time.sleep(interval)
time.sleep(interval)

# 알림 버튼 클릭(산재)
browser.find_element_by_xpath('/html/body/div[7]/div[2]/div[1]/div/div/div[2]/input[1]').click()
time.sleep(interval)

# 당월보험료 부과내역조회 화면 제거(산재)
browser.find_element_by_id('mf_wfm_content_WL0502_P04_close').click()
time.sleep(interval)

# 사업장세부상정내역 (산재)
browser.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[10]/div[1]/div/div[1]/div/div[3]/div[1]/div/table/tbody/tr[3]/td[8]').click()
time.sleep(interval)

# 근로자산정 상세 정보 엑셀저장 버튼 클릭(산재)
browser.find_element_by_id('mf_wfm_content_WL0502_P01_wframe_excelDownLoad').click()
time.sleep(interval)
time.sleep(interval)

# 알림화면 확인 버튼 클릭(산재)
browser.find_element_by_xpath('/html/body/div[8]/div[2]/div[1]/div/div/div[2]/input[1]').click()
time.sleep(interval)

# 근로자산정 상세 정보 화면 제거(산재)
browser.find_element_by_id('mf_wfm_content_WL0502_P01_close').click()
time.sleep(interval)

# =====================================================================================
# 고용 탭으로 전환
browser.find_element_by_id('mf_wfm_content_tabcont_tab_btnTabGy_tabHTML').click()
time.sleep(interval)

# 당월보험료 부과내역조회 버튼 클릭(고용)
browser.find_element_by_id('mf_wfm_content_wq_uuid_1684').click()
time.sleep(interval)
time.sleep(interval)

# 엑셀저장 버튼 클릭(고용)
browser.find_element_by_id('mf_wfm_content_WL0502_P04_wframe_wq_uuid_1984').click()
time.sleep(interval)
time.sleep(interval)

# 알림 화면 확인 버튼 클릭(고용)
browser.find_element_by_xpath('/html/body/div[8]/div[2]/div[1]/div/div/div[2]/input[1]').click()
time.sleep(interval)

# 당월보험료 부과내역조회 화면 제거(고용)
browser.find_element_by_id('mf_wfm_content_WL0502_P04_close').click()
time.sleep(interval)

# 사업장세부상정내역 (고용)
browser.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[10]/div[1]/div/div[2]/div/div[3]/div[1]/div/table/tbody/tr[1]/td[4]').click()
time.sleep(interval)

# 근로자산정 상세 정보 엑셀저장 버튼 클릭(고용)
browser.find_element_by_id('mf_wfm_content_WL0502_P00_wframe_excelDownLoad').click()
time.sleep(interval)
time.sleep(interval)

# 알림화면 확인 버튼 클릭(고용)
browser.find_element_by_xpath('/html/body/div[8]/div[2]/div[1]/div/div/div[2]/input[1]').click()
# browser.find_element_by_xpath('/html/body/div[9]/div[2]/div[1]/div/div/div[2]/input[1]').click()
time.sleep(interval)

# 근로자산정 상세 정보 화면 제거(고용)
browser.find_element_by_id('mf_wfm_content_WL0502_P00_close').click()
time.sleep(interval)


browser.quit()

