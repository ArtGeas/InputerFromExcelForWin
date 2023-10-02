from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from datetime import datetime
import time
import pyautogui
import os
import logging


# константы
URL = "https://rpachallenge.com/"
EXCEL_FILENAME = 'challenge.xlsx'
SCREENSHOT_NAME = "result.png"
LOGFILE_NAME = 'app_logs.log'

logging.basicConfig(level=logging.INFO, filename=LOGFILE_NAME,filemode="w")
_logger = logging.getLogger(__name__)

try:
    # создаем диалоговое окно для записи пути загрузки и путь до файла, который загрузим
    directory_path = pyautogui.prompt(text='Введите полный путь до папки сохранения', title='Куда сохранить excel файл' , default=os.getcwd())
    file_path = os.path.join(directory_path, EXCEL_FILENAME)
    _logger.info(f"{datetime.now()} - directories are chosen")

    # настройки загрузки для веб драйвера хрома и дальнейшая инициализация веб драйвера
    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory" : directory_path} 
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=options)
    _logger.info(f"{datetime.now()} - driver initialized")

    # разворачиваем окно на весь экран и переходим по нашему url, ждем 5 сек для прогрузки страницы
    driver.maximize_window()
    driver.get(URL)
    time.sleep(5)

    # находим по xpath кнопку download excel и нажимаем ее, ждем 5 сек чтобы загрузка успела выполниться
    driver.find_element(By.XPATH, "//*[contains(@class,'col s12 m12 l12 btn waves-effect waves-light uiColorPrimary center')]").click()
    time.sleep(5)  # set_page_load_timeout() or implicitly_wait() for error

    # читаем excel файл
    df = pd.read_excel(file_path, sheet_name='Sheet1')
    _logger.info(f"{datetime.now()} - read excel")


    # находим по xpath кнопку Start и нажимаем ее
    driver.find_element(By.XPATH, "//*[contains(@class,'waves-effect col s12 m12 l12 btn-large uiColorButton')]").click()
    _logger.info(f"{datetime.now()} - start clicked")

    # проходим циклом каждую строчку эксель файла и заполняем данные в форму сайта
    for index, row in df.iterrows():

        # находим поля по xpath в которые нужно вставлять данные из excel
        driver.find_element(By.XPATH, "//*[contains(@ng-reflect-name,'labelFirstName')]").send_keys(row['First Name'])
        driver.find_element(By.XPATH, "//*[contains(@ng-reflect-name,'labelLastName')]").send_keys(row['Last Name ']) # файле в имени этого столбика пробел в конце 
        driver.find_element(By.XPATH, "//*[contains(@ng-reflect-name,'labelCompanyName')]").send_keys(row['Company Name'])
        driver.find_element(By.XPATH, "//*[contains(@ng-reflect-name,'labelRole')]").send_keys(row['Role in Company'])
        driver.find_element(By.XPATH, "//*[contains(@ng-reflect-name,'labelAddress')]").send_keys(row['Address'])
        driver.find_element(By.XPATH, "//*[contains(@ng-reflect-name,'labelEmail')]").send_keys(row['Email'])
        driver.find_element(By.XPATH, "//*[contains(@ng-reflect-name,'labelPhone')]").send_keys(row['Phone Number'])
        driver.find_element(By.XPATH, "//*[contains(@class,'btn uiColorButton')]").click()

    _logger.info(f"{datetime.now()} - fields filled")

    # # ждем 5 сек чтобы физически успеть увидеть результат)))
    # time.sleep(5)
    driver.save_screenshot(SCREENSHOT_NAME)
    _logger.info(f"{datetime.now()} - screenshot taken")

except Exception as ex:
    _logger.info(f"{datetime.now()} - {ex}")
finally:
    # закрываем браузер и завершаем исполнение веб драйвера
    driver.quit()
    _logger.info(f"{datetime.now()} - driver quit")
