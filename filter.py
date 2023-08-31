import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import time
from multiprocessing import Pool
import multiprocessing
from art import tprint

cookies = {

}


headers = {
'authority': 'catalog.aquamarine.kz',
'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="99", "Opera GX";v="85"',
'accept': 'application/json, text/javascript, */*; q=0.01',
'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
'x-requested-with': 'XMLHttpRequest',
'sec-ch-ua-mobile': '?0',
'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.84 Safari/537.36 OPR/85.0.4341.79',
'sec-ch-ua-platform': '"Windows"',
'origin': 'https://catalog.aquamarine.kz',
'sec-fetch-site': 'same-origin',
'sec-fetch-mode': 'cors',
'sec-fetch-dest': 'empty',
'referer': 'https://catalog.aquamarine.kz/catalog/index.aspx',
'accept-language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
# Requests sorts cookies= alphabetically
# 'cookie': 'ASP.NET_SessionId=i1gip0fre5uzl4iqlkubv1cp; SLG_G_WPT_TO=ru; SLG_GWPT_Show_Hide_tmp=1; SLG_wptGlobTipTmp=1; ICusrcartgd=be6d8ad2-c52e-49b8-83b2-f384a9feaa60; IWusrsesckgd=jojhbQMjYWEdV9ohRKijJKalgxKEvPEPzVqoH/F2376n50ziaNRcMA==',
}


data = {
'msearch': '',
}
def paginations(cookies, headers, data, filterUrl): # Получение пагинации
    json = requests.post(f'https://catalog.aquamarine.kz/catalog/products.ashx?rnd=326806935&q=&spec=&mip=317&map=7777%20777&mippg=161&mappg=5466%20222&miw=0.14&maw=137.74&miq=1&maq=241&miprcs=999999.999&maprcs=0&page=1&sort=art-down&view=2&spc=1,&brid=7,&{filterUrl}', headers=headers, cookies = cookies, data=data).json()
    return json['totalPages']


def get_filters(cookies, headers): # получение всех имеющих фильтров к примеру {'Родирование': 'cid-3'} значение ключа нужно для получения данных по пост запросу
    list = {}
    r = requests.get('https://catalog.aquamarine.kz/catalog/index.aspx', headers=headers, cookies=cookies)
    soup = BeautifulSoup(r.text, 'lxml')
    filter = soup.find('table', class_='selector')
    filters = filter.find_all('label')
    for i in filters:
        list[i.text.strip().lower()] = i['for']
    return list

def read_filters(): # Читает файл с фильтрами которые нужно спарсить и возращает список с фильтрами.
    with open('filter.txt', 'r',encoding='utf-8') as f:
        lines = [line.lower() for line in f.readlines()]
    return lines

def get_page(cookies, headers, data,page): # Запросы к страницам
    return requests.post(page, headers=headers, cookies=cookies, data=data)

def getUrls(): # Получение списка страниц
    sites = [] # Список для хранения ссылок
    filters = read_filters() # Список фильтров в файле
    filtersPage = get_filters(cookies, headers) # Словарь со всеми фильтрами из страницы
    for filter in filters:
        try:
            filter = filter.replace('\n', '').replace('\ufeff','').strip() # Убираем переносы строк и пробелы.
            filterUrl = filtersPage[filter].replace('-', '=', 1).replace('cgrs','cid') # Достаём ключ из словаря со всеми фильтрами и заменяем на знак = (для пост запроса)
            pagination = int(paginations(cookies, headers, data, filterUrl)) # Находим пагинацию для каждого фильтра
            if pagination > 0: 
                for page in range(1, pagination+1): # Создаём список со ссылками на страницы
                    sites.append([f'https://catalog.aquamarine.kz/catalog/products.ashx?rnd=326806935&q=&spec=&mip=317&map=7777%20777&mippg=161&mappg=5466%20222&miw=0.14&maw=137.74&miq=1&maq=241&miprcs=999999.999&maprcs=0&page={page}&sort=art-down&view=2&spc=1,&brid=7,&{filterUrl}', filter])
        except:
            print(f'Фильтр {filter} не найден')
    return sites


def get_data(url,headers,cookies,data): # Получение данных
    try:
        soup = BeautifulSoup(get_page(cookies, headers, data,url[0]).text.replace('\\', ' '), 'lxml')
        contain = soup.find('div', class_='products')
        products = contain.find_all('div', class_='item wide') # Получение карточек
        data_page = []
        
        for product in products:
            link = 'https://catalog.aquamarine.kz'+(product.find('a').get('href')).strip() # Ссылка на карточку
            req_zoloch = requests.get(link+',&avids=6', headers=headers, cookies=cookies)
            soup4 = BeautifulSoup(req_zoloch.text, 'lxml')
            req = requests.get(link, headers=headers, cookies=cookies) 
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
                rodirovs = soup2.find_all('div', class_='center ssize') # Нахождение родирования
                count = len(rodirovs)
                for i in range(1, int(count)+1): # Проходит по все дивам с родированием и забирает размер родирования
                    rodirov = soup2.find_all('div', class_='center ssize')[i-1] 
                    size = rodirov.find_all('span')[0].text.strip().replace('/', '')
                    if size == '0':
                        break
                    data_rodirov.append(size) # Добавляется в список с родированием
                data_rodirov = '|'.join(data_rodirov)  # Объединяется в строку, если их несколько
                zolochs = soup4.find_all('div', class_='center ssize') # Нахождение родирования
                count = len(zolochs)
                for i in range(1, int(count)+1): # Проходит по все дивам с родированием и забирает размер родирования
                    zoloch = soup4.find_all('div', class_='center ssize')[i-1] 
                    size = zoloch.find_all('span')[0].text.strip().replace('/', '')
                    if size == '0':
                        break
                    data_zoloch.append(size) # Добавляется в список с родированием
                data_zoloch = '|'.join(data_zoloch)  # Объединяется в строку, если их несколько

            except: 
                data_rodirov = ''
                data_zoloch = ''
            
            try:
                headsets = []
                asets = soup2.find('div', {'style':'height:100px;overflow-y:auto;'}).find_all('a')
                for headset in asets:
                    url2 = 'https://catalog.aquamarine.kz'+headset.get('href')
                    req = requests.get(url2, headers=headers,cookies=cookies)
                    soup3 = BeautifulSoup(req.text, 'lxml')
                    articul2 = soup3.find('td', text = 'Артикул').find_next_sibling('td').text
                    headsets.append(articul2)
                headsets = ','.join(headsets)
            except:
                headsets = ''
            links_images_rodirov = []
            links_images_zoloch = []
            try:
                image ='https://catalog.aquamarine.kz' + soup2.find('div', class_='imageview').find('img').get('src')
                links_images_rodirov.append(image)
            except:
                image = ''
            try:
                image ='https://catalog.aquamarine.kz' + soup4.find('div', class_='imageview').find('img').get('src')
                links_images_zoloch.append(image)
            except:
                image = ''
            try:
                image2 = soup2.find('div', class_='imagepview').find_all('img')[1:]
                for i in image2:
                    links_images_rodirov.append('https://catalog.aquamarine.kz' + i.get('src'))
            except:
                image2 = ''
            try:
                image3 = soup4.find('div', class_='imagepview').find_all('img')[1:]
                for i in image3:
                    links_images_zoloch.append('https://catalog.aquamarine.kz' + i.get('src'))
            except:
                image3 = ''
            links_images_rodirov = ','.join(links_images_rodirov)
            links_images_zoloch = ','.join(links_images_zoloch)
            data_page.append([articul,url[1], count_in,weight,price,gramm,metal,inserts,collections,data_rodirov,data_zoloch,sizes, coverage, headsets,links_images_rodirov,links_images_zoloch])
    except Exception as ex:
        data_page[''*16]
        
    print('Cтраница получена ', url)
    return data_page

def main():
    start_time = time.time()
    tprint('Parser started')
    print('Подождите пожалуйста, идёт парсинг')
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Фильтры')
    ws.append(['Артикул','Вид изделия', 'На складе','Вес','Цена','За грамм','Металл','Вставка','Коллекция','Родирование','Золочение','Размеры','Покрытие','Гарнитуры', 'Изображения родирования', 'Изображения золочения'])
    urls = getUrls()

    records = [(url, headers, cookies,data) for url in urls]
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
    time.sleep(1000)
    
   
if __name__ == '__main__':
    multiprocessing.freeze_support()
    main()