from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
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

# определяем настройки и инициализируем логгер
logging.basicConfig(level=logging.INFO, filename=LOGFILE_NAME, filemode="w")
_logger = logging.getLogger(__name__)

# в данном блоке обрабатываем 2 случая, 1)неверная директория 2)пользователь нажал отмену при вводе
try:
    # создаем диалоговое окно для записи пути загрузки и путь до файла, который загрузим
    directory_path = pyautogui.prompt(text='Введите полный путь до папки сохранения', title='Куда сохранить excel файл' , default=os.getcwd())
    file_path = os.path.join(directory_path, EXCEL_FILENAME)
    _logger.info(f"{datetime.now()} - directories are chosen")

    # проверка введенной директории
    if not os.path.exists(directory_path):
        _logger.info(f"{datetime.now()} - Directory not exists")
        raise Exception("Directory not exists") 
except Exception as ex:
    _logger.info(f"{datetime.now()} - {ex}")
    raise Exception

# в данном блоке обрабатываем ошибки связанные с работой веб драйвера
try:
    # настройки загрузки для веб драйвера хрома и дальнейшая инициализация веб драйвера
    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory" : directory_path} 
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=options)

    # инициализируем ожидание для веб драйвера
    wait = WebDriverWait(driver, 500)
    
    _logger.info(f"{datetime.now()} - driver initialized")

    # разворачиваем окно на весь экран и переходим по нашему url
    driver.maximize_window()
    driver.get(URL)

    # ожидаем пока не прогрузится кнопка с загрузкой файла
    wait.until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[contains(@class,'col s12 m12 l12 btn waves-effect waves-light uiColorPrimary center')]")))

    # находим по xpath кнопку download excel, проматываем до нее и нажимаем, ждем 10 сек чтобы загрузка успела выполниться
    button_download_excel = driver.find_element(By.XPATH, "//*[contains(@class,'col s12 m12 l12 btn waves-effect waves-light uiColorPrimary center')]")
    driver.execute_script("arguments[0].scrollIntoView(true);", button_download_excel)
    button_download_excel.click()
    time.sleep(10) 

    # читаем excel файл
    df = pd.read_excel(file_path, sheet_name='Sheet1')
    _logger.info(f"{datetime.now()} - read excel")

    # находим по xpath кнопку Start, проматываем до нее и нажимаем
    button_start = driver.find_element(By.XPATH, "//*[contains(@class,'waves-effect col s12 m12 l12 btn-large uiColorButton')]")
    driver.execute_script("arguments[0].scrollIntoView(true);", button_start)
    button_start.click()
    _logger.info(f"{datetime.now()} - start clicked")

    # проходим циклом каждую строчку эксель файла и заполняем данные в форму сайта
    for index, row in df.iterrows():
        # ждем пока прогрузится последнее поле которое мы обрабатываем из строки
        wait.until(expected_conditions.element_to_be_clickable((By.XPATH, "//*[contains(@ng-reflect-name,'labelPhone')]")))

        # находим поля по xpath в которые нужно вставлять данные из excel
        driver.find_element(By.XPATH, "//*[contains(@ng-reflect-name,'labelFirstName')]").send_keys(row['First Name'])
        driver.find_element(By.XPATH, "//*[contains(@ng-reflect-name,'labelLastName')]").send_keys(row['Last Name ']) # файле в имени этого столбика пробел в конце 
        driver.find_element(By.XPATH, "//*[contains(@ng-reflect-name,'labelCompanyName')]").send_keys(row['Company Name'])
        driver.find_element(By.XPATH, "//*[contains(@ng-reflect-name,'labelRole')]").send_keys(row['Role in Company'])
        driver.find_element(By.XPATH, "//*[contains(@ng-reflect-name,'labelAddress')]").send_keys(row['Address'])
        driver.find_element(By.XPATH, "//*[contains(@ng-reflect-name,'labelEmail')]").send_keys(row['Email'])
        driver.find_element(By.XPATH, "//*[contains(@ng-reflect-name,'labelPhone')]").send_keys(row['Phone Number'])
        button_submit = driver.find_element(By.XPATH, "//*[contains(@class,'btn uiColorButton')]")
        driver.execute_script("arguments[0].scrollIntoView(true);", button_submit)
        button_submit.click()

    _logger.info(f"{datetime.now()} - fields filled")

    # делаем скриншот 
    driver.save_screenshot(SCREENSHOT_NAME)
    _logger.info(f"{datetime.now()} - screenshot taken")

except Exception as ex:
    _logger.info(f"{datetime.now()} - {ex}")
finally:
    # закрываем браузер и завершаем исполнение веб драйвера
    driver.quit()
    _logger.info(f"{datetime.now()} - driver quit")
