from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import re
import openpyxl
from openpyxl import load_workbook

asin_codes = []
driver = webdriver.Chrome("C:\\Users\\admin\\Desktop\\Funko\\script\\chromedriver.exe")


def get_asin(barcode):
    try:
        url = "https://www.amazon.co.uk/s?k=%s&ref=nb_sb_noss"%barcode;
        driver.get(url)
        print(url)
        asin_link = driver.find_element_by_css_selector("h2.s-line-clamp-2:nth-child(1) > a:nth-child(1)").get_attribute('href')
        print (asin_link)
        asin_code = re.search('/dp/(.*)/ref', asin_link).group(1)
        return(asin_code)
    except NoSuchElementException:
        try:
            asin_link = driver.find_element_by_css_selector(".s-line-clamp-2 > a:nth-child(1)").get_attribute('href')
            print(asin_link)
        except NoSuchElementException:
            asin_link = driver.find_element_by_css_selector("#search > div.sg-row > div.sg-col-20-of-24.sg-col-28-of-32.sg-col-16-of-20.sg-col.s-right-column.sg-col-32-of-36.sg-col-8-of-12.sg-col-12-of-16.sg-col-24-of-28 > div > span:nth-child(4) > div.s-result-list.s-search-results.sg-row > div.sg-col-4-of-24.sg-col-4-of-12.sg-col-4-of-36.s-result-item.sg-col-4-of-28.sg-col-4-of-16.sg-col.sg-col-4-of-20.sg-col-4-of-32 > div > span > div > div > div:nth-child(2) > div:nth-child(3) > div > div.a-section.a-spacing-none.a-spacing-top-small > h2 > a").get_attribute('href')
            
        

def open_file():
    file_path =input("Please give path to the file: ")
    workbook = load_workbook(filename=file_path)
    wh_sheet= workbook.active
    for cell in wh_sheet['D']:
        asin_codes.append(get_asin(cell.value))
        







