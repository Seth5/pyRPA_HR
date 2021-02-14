from selenium import webdriver
import time

interval = 2 # 페이지 전환시 2초 대기

# 크롬 브라우저 열기
browser = webdriver.Chrome()

# 국민연금 사이트로 이동
browser.get("https://edi.nps.or.kr/")
time.sleep(interval)

# master에서 다시 작업

# 마스터에서 다시 또