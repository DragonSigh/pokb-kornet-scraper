import os, time, random, json, shutil

from datetime import date, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from loguru import logger

reports_path = os.path.join(os.path.abspath(os.getcwd()), 'reports', 'from_emias')

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

service = Service('C:\chromedriver\chromedriver.exe')
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

def complex_function(x):
    if isinstance(x, str):
        first_name = x.split(' ')[1]
        second_name = x.split(' ')[2]
        last_name = x.split(' ')[3].replace(',', '')
        return f'{first_name} {second_name} {last_name}'
    else:
        return 0

def get_newest_file(path):
    files = os.listdir(path)
    paths = [os.path.join(path, basename) for basename in files]
    return max(paths, key=os.path.getctime)

def wait_for_document_ready(driver):
    WebDriverWait(driver, 60).until(lambda driver: driver.execute_script('return document.readyState;') == 'complete')

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
            if fname.endswith('.crdownload'):
                dl_wait = True
        seconds += 1
    return seconds

def autorization(login_data: str, password_data: str):
    browser.get('http://main.emias.mosreg.ru/MIS/Klimovsk_CGB/Main/Default')
    login_field = browser.find_element(By.XPATH, '//*[@id="Login"]')
    login_field.send_keys(login_data)
    password_field = browser.find_element(By.XPATH, '//*[@id="Password"]')
    password_field.send_keys(password_data)
    # Запомнить меня
    browser.find_element(By.XPATH, '//*[@id="Remember"]').click()
    browser.find_element(By.XPATH, '//*[@id="loginBtn"]').click()
    WebDriverWait(browser, 20).until(EC.invisibility_of_element((By.XPATH, '//*[@id="loadertext"]')))
    element = browser.find_element(By.XPATH, '/html/body/div[8]/div[3]/div/button/span')
    element.click()
    logger.debug('Авторизация пройдена')

def open_emias_report(cabinet_id, begin_date, end_date):
    logger.debug(f'Открываю страницу отчёта, ID кабинета: {cabinet_id}')
    element = browser.find_element(By.XPATH, '//*[@id="Portlet_9"]/div[2]/div[1]/a')
    element.click()
    browser.switch_to.window(browser.window_handles[1])
    WebDriverWait(browser, 20).until(EC.invisibility_of_element((By.XPATH, '//*[@id="loadertext"]')))
    element = browser.find_element(By.XPATH, '//*[@id="table_filter"]/label/input')
    ActionChains(browser).click(element).send_keys("v2").perform()   
    element = browser.find_element(By.XPATH, '//*[@id="table"]/tbody/tr/td[3]/a')
    element.click()
    element = browser.find_element(By.XPATH, '//*[@id="send-request-btn"]')
    WebDriverWait(browser, 20).until(EC.element_to_be_clickable(element))
    element = browser.find_element(By.XPATH, '//*[@id="Arguments_0__Value"]')
    browser.execute_script('''
        var elem = arguments[0];
        var value = arguments[1];
        elem.value = value;
    ''', element, begin_date.strftime('%d.%m.%Y') + '_' + end_date.strftime('%d.%m.%Y'))
    element = browser.find_element(By.XPATH, '//*[@id="Arguments_2__Value"]')
    browser.execute_script('''
        var elem = arguments[0];
        var value = arguments[1];
        elem.value = value;
    ''', element, cabinet_id)
    element = browser.find_element(By.XPATH, '//*[@id="Arguments_3__Value"]')
    browser.execute_script('''
        var elem = arguments[0];
        var value = arguments[1];
        elem.value = value;
    ''', element, '0')
    browser.find_element(By.XPATH, '//*[@id="send-request-btn"]').click()
    logger.debug('Отчет открыт в браузере')

def save_report(cabinet):
    logger.debug(f'Начинается сохранение файла с отчетом в папку: {reports_path}')
    # Создать папку с отчётами, если её нет в системе
    try:
        os.mkdir(reports_path)
    except FileExistsError:
        pass    
    # Сохранить в Excel
    WebDriverWait(browser, 300).until(EC.text_to_be_present_in_element_value((By.XPATH, '/html/body/div/div[2]/div/div/form[1]/input'), 'done'))
    element = browser.find_element(By.XPATH, '//*[@id="dlbId"]')
    element.click()
    download_wait(reports_path, 10)
    browser.close()
    browser.switch_to.window(browser.window_handles[0])
    logger.debug('Сохранение файла с отчетом успешно')

@retry_with_backoff(retries=5)
def start_report_saving():
    shutil.rmtree(reports_path, ignore_errors=True) # Очистить предыдущие результаты
    credentials_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'auth-emias.json')
    # С начала недели
    first_date = date.today() - timedelta(days=date.today().weekday()) # начало текущей недели
    last_date = date.today() # сегодня
     # Сегодня
    #first_date = date.today()
    #last_date = date.today()
    # За прошлую неделю
    #first_date = date.today() - timedelta(days=date.today().weekday()) - timedelta(days=7) # начало прошлой недели
    #last_date = first_date + timedelta(days=6) # конец недели
    # Задать даты вручную
    #first_date = datetime.datetime.strptime('24.05.2023', '%d.%m.%Y').date()
    #last_date  = datetime.datetime.strptime('25.05.2023', '%d.%m.%Y').date()
    # Если сегодня понедельник, то берем всю прошлую неделю
    if date.today() == (date.today() - timedelta(days=date.today().weekday())):
        first_date = date.today() - timedelta(days=date.today().weekday()) - timedelta(days=7) # начало прошлой недели
    # Открываем данные для авторизации и проходим по списку кабинетов
    logger.debug(f'Выбран период: с {first_date.strftime("%d.%m.%Y")} по {last_date.strftime("%d.%m.%Y")}')
    f = open(credentials_path, 'r', encoding='utf-8')
    data = json.load(f)
    f.close()
    for _departments in data['departments']:        
        logger.debug(f'Начинается сохранение отчёта для подразденения: {_departments["department"]}')
        for _units in _departments["units"]:
            logger.debug(f'Начинается авторизация в отделение: {_units["name"]}')
            autorization(_units['login'], _units['password'])
    # ID кабинетов выписки лекарств
    cabinets_list = ['2434', '2460', '2459', '2450', '636', '2458', '2343', '2457', '2449']
    for cabinet in cabinets_list:
        open_emias_report(cabinet, first_date, last_date)
        save_report(cabinet)
        #os.remove(get_newest_file(reports_path))
    logger.debug('Выгрузка из ЕМИАС завершена')

start_report_saving()

browser.quit()