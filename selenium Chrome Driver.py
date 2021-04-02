from selenium import webdriver

driver = webdriver.Chrome("C:\\Users\\이영준\\python 파일들\\Chromedriver")

url = 'http://www.google.com'

browser = driver.get(url)

search_bar = browser.find_element_by_css_selector('.gLFyf.gsfi"')

search_bar.send_keys('젤다의 전설 야생의 숨결')