from selenium.webdriver import ActionChains
from selenium import webdriver
import time
import redis, copy
import numpy as np

def get_glass_book(page_num):
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome("E:\chromedriver\chromedriver.exe", chrome_options=options)
    driver.implicitly_wait(1)
    key = "http://m.26ksw.cc/book/58634/53430884.html"
    driver.get(key)
    for num in range(page_num):
        time.sleep(2)
        more_detail = driver.find_elements(by="xpath", value='/html/body/div[1]')
        text_get1 = more_detail[0].text
        time.sleep(1)
        write_book("glass_book_clawer.txt", text_get1)
        next_page = driver.find_elements(by="xpath", value='/html/body/p[2]/a[3]')
        action = ActionChains(driver)
        action.click(next_page[0]).perform()


def write_book(file,text):
    with open(file,"a+") as f:
        f.write(text)

if __name__ == '__main__':
    page_num = 2264
    get_glass_book(page_num)