from selenium import webdriver
# from selenium.common import StaleElementReferenceException
from selenium.webdriver.chrome.options import Options
from fake_useragent import UserAgent
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import time


def create_xls_file():
    wb = openpyxl.Workbook()
    file_name = 'products.xlsx'
    work_sheet = wb.active
    work_sheet.append(['num', 'product_name', 'url'])
    wb.save(file_name)
    return wb, file_name


# ----------------------------------------------------------------------------------------------------------------------
# https://selenium-python.readthedocs.io/waits.html

def main():
    website = "https://www.emersonecologics.com/shop/Adrenal-Stress-Support-S321"
    path = "C:\Program Files (x86)\chromedriver.exe"

    ua = UserAgent()
    userAgent = ua.random

    options = Options()
    # options.add_argument(f'user-agent={userAgent}')
    # options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(chrome_options=options, executable_path=path)
    driver.get(website)

    website_pages_num = len(WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located(
            (By.XPATH, '//md-option[@ng-value="page"]'))))

    print(f"Website pages number: {website_pages_num}")

    file, file_name = create_xls_file()[0], create_xls_file()[1]
    work_sheet_00 = file.active

    num = 1

    for page in range(1, website_pages_num + 1):

        print(f"Page number: {page}", end='\n')
        website = f"https://www.emersonecologics.com/shop/search?ff=321&fft=s&term=*&page={page}&s=321"
        driver.get(website)
        time.sleep(3)

        urls = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located(
                (By.XPATH, '//h3[@class="ng-scope"]//a[@class="ng-binding"]')))

        for i in urls:
            work_sheet_00.append([num, i.text, i.get_attribute('href')])
            file.save('products.xlsx')
            print(f"{num}: {i.text} {i.get_attribute('href')}")
            num += 1

    print(f"Web scraping processed!\nFound elements: {num}")


if __name__ == '__main__':
    main()
