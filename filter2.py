import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import time
from multiprocessing import Pool
import multiprocessing
from dotenv import load_dotenv
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pickle
from selenium.webdriver.chrome.service import Service
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning) 
load_dotenv()

email = os.getenv('LOGIN')
password_input = os.getenv('PASSWORD')




headers = {
'authority': 'catalog.aquamarine.kz',
'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="99", "Opera GX";v="85"',
'accept': 'application/json, text/javascript, */*; q=0.01',
'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
'x-requested-with': 'XMLHttpRequest',
'sec-ch-ua-mobile': '?0',
'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.84 Safari/537.36 OPR/85.0.4341.79',
'sec-ch-ua-platform': '"Windows"',
'origin': 'https://catalog.aquamarine.gold',
'sec-fetch-site': 'same-origin',
'sec-fetch-mode': 'cors',
'sec-fetch-dest': 'empty',
'referer': 'https://catalog.aquamarine.gold/catalog/index.aspx',
'accept-language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
# Requests sorts cookies= alphabetically
# 'cookie': 'ASP.NET_SessionId=i1gip0fre5uzl4iqlkubv1cp; SLG_G_WPT_TO=ru; SLG_GWPT_Show_Hide_tmp=1; SLG_wptGlobTipTmp=1; ICusrcartgd=be6d8ad2-c52e-49b8-83b2-f384a9feaa60; IWusrsesckgd=jojhbQMjYWEdV9ohRKijJKalgxKEvPEPzVqoH/F2376n50ziaNRcMA==',
}


data = {
'msearch': '',
}
def get_cookies():
    try:
        options = Options()
        service = Service()
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument("--start-maximized")
        
        driver = webdriver.Chrome(service=service, options=options)
        driver.get('https://catalog.aquamarine.gold/register/login.aspx?ref=%2fcatalog%2findex.aspx')
        time.sleep(2)
        login = driver.find_element(By.CSS_SELECTOR, '[placeholder="Email"]')

        password = driver.find_element(By.CSS_SELECTOR, '[placeholder="Пароль"]')

        login.send_keys(email)

        password.send_keys(password_input)
        enter = driver.find_element(By.CSS_SELECTOR, '[value="ВОЙТИ"]').click()

        time.sleep(2)

        cookies = driver.get_cookies()
        
        driver.close()
        driver.quit()

        with open('cookies.pkl', 'wb') as f:
            pickle.dump(cookies, f)  # Записываем cookies в файл
    except Exception as ex:
        print(ex)
        time.sleep(1000)
    return cookies

if os.path.exists('cookies.pkl'):
    with open('cookies.pkl', 'rb') as f:
        cookies_from_file = pickle.load(f)
else:
    # Если файла с cookies нет, получаем их с помощью Selenium и записываем в файл
    get_cookies()
    with open('cookies.pkl', 'rb') as f:
        cookies_from_file = pickle.load(f)

# Создаем сессию requests и устанавливаем cookies
s = requests.Session()

for cookie in cookies_from_file:
    s.cookies.set(cookie['name'], cookie['value'])


def paginations(headers, data, filterUrl): # Получение пагинации
    json = s.post(f'https://catalog.aquamarine.gold/catalog/products.ashx?rnd=955053327&q=&spec=&mip=34&map=2083%20566&mippg=195&mappg=852565&miw=0.11&maw=137.74&miq=1&maq=4476&miprcs=0.012&maprcs=6.52&page=1&sort=art-down&view=1&cbcg=1,&brid=1,&{filterUrl}', headers=headers, data=data).json()
    return json['totalPages']


def get_filters(headers): # получение всех имеющих фильтров к примеру {'Родирование': 'cid-3'} значение ключа нужно для получения данных по пост запросу
    list = {}
    r = s.get('https://catalog.aquamarine.gold/catalog/index.aspx', headers=headers)
    soup = BeautifulSoup(r.text, 'lxml')
    filter = soup.find('table', class_='selector')
    filters = filter.find_all('label')
    for i in filters:
        list[i.text.strip().lower()] = i['for']
    return list

def read_filters(): # Читает файл с фильтрами которые нужно спарсить и возращает список с фильтрами.
    with open('filter.txt', 'r',encoding='utf-8') as f:
        lines = [line.strip().lower() for line in f.readlines()]
        
    return lines

def get_page(headers, data,page): # Запросы к страницам
    return s.post(page, headers=headers, data=data)

def getUrls(): # Получение списка страниц
    sites = [] # Список для хранения ссылок
    filters = read_filters() # Список фильтров в файле
    filtersPage = get_filters(headers) # Словарь со всеми фильтрами из страницы
    for filter in filters:
        try:
            filter = filter.replace('\ufeff','').strip() # Убираем переносы строк и пробелы.
            filterUrl = filtersPage[filter].replace('-', '=', 1).replace('cgrs','cid') # Достаём ключ из словаря со всеми фильтрами и заменяем на знак = (для пост запроса)
            pagination = int(paginations(headers, data, filterUrl)) # Находим пагинацию для каждого фильтра
            if pagination > 0: 
                for page in range(1, pagination+1): # Создаём список со ссылками на страницы
                    sites.append([f'https://catalog.aquamarine.gold/catalog/products.ashx?q=&spec=&mip=34&mappg=852565&miw=0.11&maw=137.74&miq=1&maq=4476&miprcs=0.012&maprcs=6.52&page={page}&sort=art-down&view=2&{filterUrl}', filter])
        except Exception as ex:
            print(ex)
            print({filter}, 'не найден')
    return sites

def get_rodirov_html(soup, type):
    start_heading = f"Заказ со склада - {type}"
    end_heading = "Заказ в производство"

    start_found = False
    end_found = False

    selected_elements = []

    for element in soup.find_all():
        if not start_found and element.name == 'h3' and element.text.strip() in start_heading:
            start_found = True
            continue
        if start_found and not end_found:
            if element.name == 'h3' and element.text.strip() in end_heading:
                end_found = True
                break
            selected_elements.append(str(element))
    
    return selected_elements

def get_data(url,headers,data): # Получение данных
    soup = BeautifulSoup(get_page(headers, data,url[0]).text.replace('\\', ' '), 'lxml')
    contain = soup.find('div', class_='products')
    products = contain.find_all('div', class_='item wide') # Получение карточек
    data_page = []
    try:
        for product in products:
            link = 'https://catalog.aquamarine.gold'+(product.find('a').get('href')).strip() # Ссылка на карточку
            req_zoloch = s.get(link+',&avids=6', headers=headers)
            soup4 = BeautifulSoup(req_zoloch.text, 'lxml')
            req = s.get(link, headers=headers) 
            soup2 = BeautifulSoup(req.text, 'lxml')
            articul = soup2.find('td', text = 'Артикул').find_next_sibling('td').text # Артикул  карточки
            weight = soup2.find_all('tr', class_='data')[3].find_all('td')[1].text
            price = soup2.find('td', text = 'Цена').find_next_sibling('td').text
            gramm = soup2.find('span',class_='price' ).text
            metal = soup2.find('td', text = 'Металл').find_next_sibling('td').text
            try:
                count_in = product.find_all('div', class_='row')[-2].find('div', class_='val').text
            except:
                count_in = 0
            
            try:
                inserts = soup2.find('td', text = 'Вставки').find_next_sibling('td').text
            except:
                inserts = ''
            try:
                collections = soup2.find('td', text = 'Коллекция').find_next_sibling('td').text
            except:
                collections = '' 
            try:
                sizes = []
                info = soup2.find('table', class_='assetinfo').find_all('tr')
                count = len(info)
                info = soup2.find('table', class_='assetinfo').find_all('tr')
                for i in info:
                    try:
                        size = i.find_all('td')[2].text.replace(',', '.').strip()
                        if size != '-':
                            sizes.append(size)
                    except:
                        pass

                sizes = ','.join(sizes)
            except:
                sizes = ''
            try:
                coverage = []
                if soup2.find('td', text = 'Наличие').find_next_sibling('td').text == 'Есть':
                    coverage.append(soup2.find('option', {'selected':'selected'}).text.strip().split(' ')[-1])
                if soup4.find('td', text = 'Наличие').find_next_sibling('td').text == 'Есть':
                    coverage.append('золочение')
                coverage = '|'.join(coverage)
            except:
                coverage = ''
            try:        
                data_rodirov = [] 
                data_zoloch = []     
                rodirovs = BeautifulSoup(' '.join(get_rodirov_html(soup2,'Родирование')), 'lxml').find_all('div', class_='center ssize')                                 
                if rodirovs:
                    for rodirov in rodirovs: # Проходит по все дивам с родированием и забирает размер родирования
                        size = rodirov.find_all('span')[0].text.strip().replace('/', '')
                        if size == '0':
                            break
                        data_rodirov.append(size) # Добавляется в список с родированием
                data_rodirov = '|'.join(data_rodirov)  # Объединяется в строку, если их несколько
                zolochs = BeautifulSoup(' '.join(get_rodirov_html(soup4,'Золочение')), 'lxml').find_all('div', class_='center ssize') 
                if zolochs:
                    for zoloch in zolochs:
                        size = zoloch.find_all('span')[0].text.strip().replace('/', '')
                        if size == '0':
                            break
                        data_zoloch.append(size) # Добавляется в список с золочения
                data_zoloch = '|'.join(data_zoloch)  # Объединяется в строку, если их несколько



            except Exception as e:
                print(e)
                data_rodirov = ''
                data_zoloch = ''
            
            try:
                variants_rodirov = []

                if soup2.find('h3', text = 'Варианты модели'):
                    for img in soup2.find('h3', text = 'Варианты модели').find_next_sibling('div').find_all('img'):
                        variants_rodirov.append(img['title'].split(' ')[0])
                variants_rodirov = ','.join(variants_rodirov)

            except:
                variants_rodirov = ''
            try:
                headsets = []
                asets = soup2.find('div', {'style':'height:100px;overflow-y:auto;'}).find_all('a')
                for headset in asets:
                    url2 = 'https://catalog.aquamarine.gold'+headset.get('href')
                    req = s.get(url2, headers=headers)
                    soup3 = BeautifulSoup(req.text, 'lxml')
                    articul2 = soup3.find('td', text = 'Артикул').find_next_sibling('td').text
                    headsets.append(articul2)
                headsets = ','.join(headsets)
            except:
                headsets = ''
            links_images_rodirov = []
            links_images_zoloch = []
            try:
                image ='https://catalog.aquamarine.gold' + soup2.find('div', class_='imageview').find('img').get('src')
                links_images_rodirov.append(image)
            except:
                image = ''
            try:
                image ='https://catalog.aquamarine.gold' + soup4.find('div', class_='imageview').find('img').get('src')
                links_images_zoloch.append(image)
            except:
                image = ''
            try:
                image2 = soup2.find('div', class_='imagepview').find_all('img')[1:]
                for i in image2:
                    links_images_rodirov.append('https://catalog.aquamarine.gold' + i.get('src'))
            except:
                image2 = ''
            try:
                image3 = soup4.find('div', class_='imagepview').find_all('img')[1:]
                for i in image3:
                    links_images_zoloch.append('https://catalog.aquamarine.gold' + i.get('src'))
            except:
                image3 = ''
            links_images_rodirov = ','.join(links_images_rodirov)
            links_images_zoloch = ','.join(links_images_zoloch)
            data_page.append([articul,url[1], count_in,weight,price,gramm,metal,inserts,collections,data_rodirov,data_zoloch,sizes, coverage, headsets,links_images_rodirov,links_images_zoloch, variants_rodirov])
    except:
        data_page[''*17]

    print('Cтраница получена ', url)
    return data_page

def main():
    try:
        start_time = time.time()
        print('Parser started')
        print('Подождите пожалуйста, идёт парсинг')
        wb = Workbook()
        wb.remove(wb.active)
        ws = wb.create_sheet('Фильтры')
        ws.append(['Артикул','Вид изделия', 'На складе','Вес','Цена','За грамм','Металл','Вставка','Коллекция','Родирование','Золочение','Размеры','Покрытие','Гарнитуры', 'Изображения родирования', 'Изображения золочения', 'Варианты модели'])
        urls = getUrls()

        records = [(url, headers, data) for url in urls]
        with Pool(multiprocessing.cpu_count()) as p:
            name = p.starmap(get_data, records)
            p.close()
            p.join()
        for x in name:
            for y in x:
                ws.append(y) # Добавляет все артикулы
                
            # Сохранение файла
        wb.save('Data.xlsx')
        print(f'Парсинг занял {round((time.time()-start_time)/60, 1)} минут')
    except Exception as ex:
        print(ex)
        
    time.sleep(1000)
   
if __name__ == '__main__':
    multiprocessing.freeze_support()
    main()  