from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import undetected_chromedriver as uc
import time
from openpyxl import Workbook

if __name__ == '__main__':
    options = Options()
    options.add_argument("--disable-extensions")
    options.binary_location = 'C:\Program Files\Google\Chrome Beta\Application\chrome.exe'
    driver = webdriver.Chrome(service=Service(ChromeDriverManager(version='104.0.5112.20').install()), options=options)
    driver.maximize_window()
    driver.implicitly_wait(30)
    driver.get('https://www.motosport.com/')


    time.sleep(1.2)
    oem_parts = driver.find_element(By.CSS_SELECTOR,"li[class='level-one-item nav_menu_link_drop_3']")
    oem_parts.click()

    time.sleep(1.2)
    machine_makes =  oem_parts.find_elements(By.CSS_SELECTOR, "a[class=' gtm-nav']")
    machine_make_text_list = []
    machine_make_links = []
    final_list = []

    for machine_make in machine_makes:
        machine_make_links.append(machine_make.get_attribute('href'))
        machine_make_text_list.append(machine_make.text)

    for machine_make_link,machine_make_text in zip(machine_make_links,machine_make_text_list):
        make_text = machine_make_text
        #print(make_text)
        #print(machine_make_link)
        driver.get(machine_make_link)
        try:
            years_grid = driver.find_element(By.CSS_SELECTOR,"div[class='ui agnostic-year grid']")
            years = years_grid.find_elements(By.TAG_NAME,"a")

            years_text_list = []
            years_links_list = []

            for year in years:
                years_links_list.append(year.get_attribute('href'))
                years_text_list.append(year.text)

            for year_link,year_text in zip(years_links_list,years_text_list):
                #time.sleep(1)
                driver.get(year_link)
                time.sleep(1)

                section = driver.find_element(By.CSS_SELECTOR, "div[class='twenty wide phablet twelve wide tablet twelve wide computer twelve wide large screen twelve wide widescreen column']")

                machine_types = section.find_elements(By.TAG_NAME,'h3')[1:]
                machine_models_grid = section.find_elements(By.CSS_SELECTOR,"div[class='ui grid oem-make-year-list']")

                for machine_type, machine_model_grid in zip(machine_types,machine_models_grid):
                    machine_type_text=machine_type.text
                    machine_models = machine_model_grid.find_elements(By.TAG_NAME,'a')

                    for machine_model in machine_models:
                        model_text = machine_model.text
                        list_to_append = [make_text,year_text,machine_type_text,model_text]
                        final_list.append(list_to_append)
                        print(list_to_append)
        except:
            pass

    workbook_name = 'Final_results_of_scraping.xlsx'
    wb = Workbook()
    page = wb.active
    headers = ['Type', 'Make', 'Series', 'Model']
    page.append(headers)
    for row in final_list:
        page.append(row)
    wb.save(filename='Final_results_of_scraping.xlsx')






