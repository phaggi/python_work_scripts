#from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from elementium.drivers.se import SeElements
from yaml import load
from yaml import FullLoader
from datetime import datetime, timedelta
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import os
from pathlib import Path

FILIALS = {'30': 'Белгородский',
 '31': 'Брянский и Орловский',
 '86': 'Владимирский и Ивановский',
 '34': 'Воронежский',
 '35': 'Калужский',
 '36': 'Курский',
 '37': 'Липецкий',
 '9': 'МОСКВА',
 '44': 'Тульский и Рязанский',
 '40': 'Смоленский',
 '41': 'Тамбовский',
 '43': 'Тверской',
 '45': 'Ярославский и Костромской'}

class BetterChrome(webdriver.Chrome):

    default_timeout = 10
    better_setting = {
        'options': '', 'path':''
    }
    def set_better_setting(self, options, path):
        BetterChrome.better_setting['options'] = options
        BetterChrome.better_setting['path'] = path
        self.set_page_load_timeout(BetterChrome.default_timeout)

    def better_refresh(self):
        return self.better_load_page(self.current_url)


    def better_load_page(self, url):
        load_retry = 0
        timeout = BetterChrome.default_timeout
        default_setting = BetterChrome.better_setting
        while 1:
            print('retrying')
            if load_retry > 8:
                return False, False
            try:
                self.get(url)
                print('load success!')
                self.set_better_setting(default_setting['options'], default_setting['path'])
                return True, self
            except TimeoutException:
                load_retry += 1
                # print('retrying')
                self.quit()

                self = BetterChrome(chrome_options=default_setting['options'], executable_path=default_setting['path'])
                self.maximize_window()
                timeout += 15
                self.set_page_load_timeout(timeout)

def wait_new_file(path_to_watch):
    print('Ждем появления файла в папке сохранения')
    before = dict ([(f, None) for f in os.listdir (path_to_watch)])
    result = True
    while result:
        time.sleep (10)
        after = dict ([(f, None) for f in os.listdir (path_to_watch)])
        added = [f for f in after if not f in before]
        removed = [f for f in before if not f in after]
        if added: 
            print("Добавлено: ", ", ".join (added))
            result = False
        if removed: print("Удалено: ", ", ".join (removed))
        before = after
    return added

def get_dates():
    date_format = '%d.%m.%Y'
    today = datetime.now()
    yesterday = today + timedelta(days=-1)
    tomorrow = today + timedelta(days=1)
    after_tomorrow = today + timedelta(days=2)
    return {'today': today.strftime(date_format),
            'tomorrow': tomorrow.strftime(date_format),
            'after_tomorrow': after_tomorrow.strftime(date_format),
           'yesterday': yesterday.strftime(date_format)}

date_from = '01.01.2020'
#  date_to = get_dates()['yesterday']

#  date_to = get_dates()['today']

date_to = '31.07.2020'

home = str(Path.home())
pathsecret = home + '\\secret\\'
namesecret = 'config_hermes.yaml'
path = pathsecret + namesecret
f = open(path)
config = load(f, Loader=FullLoader)

pathtosave = os.getcwd()
if pathtosave == home + '\\_Python':
    pathtosave = 'C:\\test'

username = config['username']
password = config['password']

CHROMEPATH = 'C:\\selen\\chromedriver.exe'
if not os.path.exists(CHROMEPATH):
    print('Хромдравер не обнаружен по пути ', CHROMEPATH)
    raise SystemExit(1)

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {"profile.default_content_settings.popups": 0,
         "download.default_directory": pathtosave, # IMPORTANT - ENDING SLASH V IMPORTANT
         "directory_upgrade": True}
options.add_experimental_option("prefs", prefs)

driver = BetterChrome(options = options, executable_path = CHROMEPATH)
driver.set_better_setting(options, CHROMEPATH)
se = SeElements(driver) # указываем, что будем использовать драйвер Chrome, и в случае чего ChromeDriverManager установит его для нас

url = "https://hermes-prod.rt.ru/"
se.navigate(url)
field = 'username'
search_field = se.find("[name='" + field + "']", wait=True)
search_field.write(username) # вводим поисковый запрос и нажимаем Enter, как бы мы сделали это руками

field = 'password'
search_field = se.find("[name='" + field + "']", wait=True)
search_field.write(password + Keys.RETURN) # вводим поисковый запрос и нажимаем Enter, как бы мы сделали это руками

def wait(title):
    while not se.title() == title:
        time.sleep(0.5)
    return(True)

title = 'Гермес'
wait(title)

se.xpath('/html/body/nav/div[2]/ul[1]/li/a', wait=True).click() #  открываем меню вверху

try:
#    se.find('[href="/aggregator/menu"]', wait=True).click() #  жмем Агрегатор
    se.xpath('/html/body/nav/div[2]/ul[1]/li/ul/li[6]/a', wait=True).click() #  жмем Агрегатор
except Exception:
    print('Проблема с меню вверху')
se.scroll_bottom
try:
    se.xpath('//*[@id="menu_arm"]/div/a[1]', wait=True).write(Keys.END) #  Прокручиваем Агрегатор до низу
except Exception:
    print('проблема с Агрегатором')

se.find("[href='#collapse_b2c']", wait=True).click() #  Разворачиваем В2С

se.find("[href='#collapse_b2c']", wait=True).write(Keys.END) #  Прокручиваем до низу

link = '[href="#collapse_b2c_reports"]' #  Разворачиваем отчеты
se.find(link, wait=True).click()

se.find("[href='#collapse_b2c']", wait=True).write(Keys.END) #  Прокручиваем до низу

field = 'href="/report/b2c_agg_construction_by_techs_report_rf"' #  кликаем по ссылке отчета Отчет по фактическому строительству 
                                                  #  (с разбиением по технологиям)

link = '[' + field + ']'
se.find(link, wait=True).click()

link = 'https://hermes-prod.rt.ru/report/b2c_agg_construction_by_techs_report_rf' #  открываем еще одну закладку отчета, 
                                                                  #  потому что не знаю, как перейти на ту вкладку

se.navigate(link)

title = 'Сводный отчет по строительству B2C (с разбиением по технологиям)' #  Проверяю, что попали на правильную страничку
wait(title)

field = "availabilityDatePeriodStart" #  заполняю поле "от"
def fill_data(field, data):
    link = '[id=\"' + field + '\"]'
    pole = se.find(link, wait=True)[0]
    for _ in range(14):
        pole.write(Keys.BACKSPACE)
    pole.write(data)
fill_data(field, date_from)

field = "availabilityDatePeriodEnd" #  заполняю поле "до"
fill_data(field, date_to)

pole = se.find('[id="departmentIds"]', wait=True) #  прокручиваю список МРФов кнопками
pole.write(Keys.PAGE_DOWN)
pole.write(Keys.CONTROL+Keys.UP)


pole.select(value="7") #   выделяю Центр, снимаю выделение с остальных
for i in range(3,6):
    pole.deselect(value=str(i))

pole = se.find('[name="departmentIds[1][]"]', wait=True) #  прокручиваю список МРФов кнопками
for i in FILIALS.keys():
    pole.select(value=str(i))

pole = se.find('[name="addressPlanYears[]"]', wait=True) #  выбираю 2020 год
pole.select(value="2020") #   выделяю 2020    

se.find('body',wait=True).write(Keys.PAGE_DOWN) #  прокручиваю страницу до низу
se.find('[class="btn btn-default dropdown-toggle"]',wait=True)[0].click() #  раскрываю список возможных типов закрузки


se.find('[formaction="/report/b2c_agg_construction_by_techs_report_rf/advancedxlsx"]',wait=True).click() #  кликаю в нужный тип загрузки

def download(_pathtosave): #  жду, когда появится новый файл в целевой папке сохранения
    _result = False
    saved_file = ''
    saved_file = wait_new_file(_pathtosave)
    _result = True
    return _result, saved_file

result = download(pathtosave)

def makefinaltable(sourcefilename, pathtosave):
    import pandas as pd
    import numpy as np

    fullpathtofile = pathtosave + '\\' + sourcefilename
    print(fullpathtofile)
    df = pd.read_excel(fullpathtofile, encoding='utf-8', header=7, na_values='NaN')

    def droprows(dfObj, colname, value2drop):
        # Get names of indexes for which column Age has value 30
        indexNames = dfObj[dfObj[colname] == value2drop].index
        # Delete these row indexes from dataFrame
        dfObj.drop(indexNames , inplace=True)

    colname = 'Технология'
    value2drop = 'Суммарно по технологии'

    droprows(df, colname, value2drop)

    colname = 'МРФ'
    value2drop = 'Суммарно за МРФ Центр'
    droprows(df, colname, value2drop)
    resultfilename = 'result_' + sourcefilename
    fullpathtosave = pathtosave + '\\' + resultfilename
    writer = pd.ExcelWriter(fullpathtosave)
    df.to_excel(writer,'Sheet1',index=False)
    writer.save()
    return fullpathtosave
    
if result[0]: #  отчитываюсь и завершаю
    print()
    print('выгрузка ', end='')
    sourcefilename = result[1][0]
    print(sourcefilename) 
    print('сохранена в папке:', pathtosave, '\nЗагрузка завершена успешно.\nНачата обработка.\n')
    se.browser.quit()
    print('Обработанный файл ', makefinaltable(sourcefilename, pathtosave), ' сохранен в той же папке.')
    
else:
    print('че-та пошло не так...')
