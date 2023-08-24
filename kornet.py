import os, time, random, json, shutil
import pandas as pd

from datetime import date
from datetime import timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

from loguru import logger

reports_path = os.path.join(os.path.abspath(os.getcwd()), 'reports', 'from_kornet')

options = webdriver.ChromeOptions()
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_argument("--start-maximized")
options.add_argument("--disable-extensions")
options.add_argument("--disable-popup-blocking")
options.add_argument("--headless=new")
options.add_experimental_option("prefs", {
  "download.default_directory": reports_path,
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

service = Service('C:\changedetection\chromedriver\chromedriver.exe')
browser = webdriver.Chrome(options=options, service=service)
actions = ActionChains(browser)

def retry_with_backoff(retries = 5, backoff_in_seconds = 1):
    def rwb(f):
        def wrapper(*args, **kwargs):
          x = 0
          while True:
            try:
              return f(*args, **kwargs)
            except:
              if x == retries:
                raise
              sleep = (backoff_in_seconds * 2 ** x +
                       random.uniform(0, 1))
              time.sleep(sleep)
              x += 1
        return wrapper
    return rwb

def wait_for_document_ready(driver):
    WebDriverWait(driver, 10).until(lambda driver: driver.execute_script('return return document.readyState;' == 'complete'))

def download_wait(directory, timeout, nfiles=None):
    """
    Wait for downloads to finish with a specified timeout.

    Args
    ----
    directory : str
        The path to the folder where the files will be downloaded.
    timeout : int
        How many seconds to wait until timing out.
    nfiles : int, defaults to None
        If provided, also wait for the expected number of files.
    """
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < timeout:
        time.sleep(1)
        dl_wait = False
        files = os.listdir(directory)
        if nfiles and len(files) != nfiles:
            dl_wait = True
        for fname in files:
            if not 'ReestrDLO.xlsx' in fname:
                dl_wait = True
        seconds += 1
    return seconds

def autorization(login_data: str, password_data: str):
    browser.get('http://llo.emias.mosreg.ru/korvet/admin/signin')
    browser.refresh()
    login_field = browser.find_element(By.XPATH, '//*[@id="content"]/div/div/form/div[1]/input')
    login_field.send_keys(login_data)
    password_field = browser.find_element(By.XPATH, '//*[@id="content"]/div/div/form/div[2]/input')
    password_field.send_keys(password_data)
    browser.find_element(By.XPATH, '//*[@id="content"]/div/div/form/div[4]/button').click()
    logger.debug('Авторизация пройдена')

def open_report():
    logger.debug('Открываю страницу отчёта')
    browser.get('http://llo.emias.mosreg.ru/korvet/FiltersLocalReport.aspx?guid=85122D62-3F72-40B5-A7ED-B2AFBF27560B')
    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_plate_BeginDate"]')))
    element = browser.find_element(By.XPATH, '//*[@id="ctl00_plate_BeginDate"]')
    ActionChains(browser).click(element).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).send_keys("01.08.2023").perform()
    element = browser.find_element(By.XPATH, '//*[@id="ctl00_plate_EndDate"]')
    ActionChains(browser).click(element).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).send_keys("08.08.2023").perform()
    browser.find_element(By.XPATH, '//*[@id="ctl00_plate_sumbit"]').click()
    logger.debug('Отчет сформирован в браузере')

def open_dlo_report(begin_date, end_date):
    logger.debug('Открываю страницу отчёта')
    browser.get('http://llo.emias.mosreg.ru/korvet/LocalReportForm.aspx?guid=85122D62-3F72-40B5-A7ED-B2AFBF27560B&FundingSource=0&BeginDate=' + begin_date.strftime('%d.%m.%Y') + '&EndDate=' + end_date.strftime('%d.%m.%Y'))
    logger.debug('Отчет сформирован в браузере')

def save_report():
    logger.debug(f'Начинается сохранение файла с отчетом в папку: {reports_path}')
    # Создать папку с отчётами, если её нет в системе
    try:
        os.mkdir(reports_path)
    except FileExistsError:
        pass    
    # Ожидать загрузки отчёта в веб-интерфейсе
    WebDriverWait(browser, 30).until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/table/tbody/tr/td/div/span/div/table/tbody/tr[4]/td[3]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[8]')))
    # Выполнить javascript для выгрузки  в Excel, который прописан в кнопке
    browser.execute_script("$find('ctl00_plate_reportViewer').exportReport('EXCELOPENXML');")
    download_wait(reports_path, 20, 1)
    logger.debug('Сохранение файла с отчетом успешно')
    browser.get('http://llo.emias.mosreg.ru/korvet/Admin/SignOut')

@retry_with_backoff(retries=5)
def start_report_saving():
    shutil.rmtree(reports_path, ignore_errors=True) # Очистить предыдущие результаты
    credentials_path = os.path.join(os.path.abspath(os.getcwd()), 'auth-kornet.json')
    # С начала недели
    first_date = date.today() - timedelta(days=date.today().weekday()) # начало текущей недели
    last_date = date.today() # сегодня
    # Если сегодня понедельник, то берем всю прошлую неделю
    if date.today() == (date.today() - timedelta(days=date.today().weekday())):
        first_date = date.today() - timedelta(days=date.today().weekday()) - timedelta(days=7) # начало прошлой недели

    logger.debug(f'Выбран период: с {first_date.strftime("%d.%m.%Y")} по {last_date.strftime("%d.%m.%Y")}')
    f = open(credentials_path, 'r', encoding='utf-8')
    data = json.load(f)
    f.close()
    for _departments in data['departments']:
        df_list = []
        logger.debug(f'Начинается сохранение отчёта для подразденения: {_departments["department"]}')
        for _units in _departments["units"]:
            logger.debug(f'Начинается авторизация в отделение: {_units["name"]}')
            autorization(_units['login'], _units['password'])
            open_dlo_report(first_date, last_date)
            save_report()
            df_temp = pd.read_excel(os.path.join(reports_path, 'ReestrDLO.xlsx'), skiprows=range(1, 12), skipfooter=18, usecols = "C,D,E,I,L,N,P,S,X")
            df_temp.insert(0, 'Отделение', _units['name'])
            df_list.append(df_temp)
            #os.rename(os.path.join(reports_path, 'ReestrDLO.xlsx'), os.path.join(reports_path, _units['name'] + '.xlsx'))
            os.remove(os.path.join(reports_path, 'ReestrDLO.xlsx'))
        final_df = pd.concat(df_list)
        final_df.columns = ['Отделение', 'Серия и номер', 'Дата выписки', 'ФИО врача', 'СНИЛС', 'ФИО пациента', 'Код категории', 'Адрес', 'Препарат', 'Количество']
        final_df.to_excel(os.path.join(reports_path, _departments['department'] + '.xlsx'), index=False)
    logger.debug('Выгрузка из КОРНЕТА завершена')

start_report_saving()

browser.quit()