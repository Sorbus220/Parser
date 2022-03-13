import time
from lxml import html
import requests
from pprint import pprint
import pandas as pd
import openpyxl as ox
from datetime import datetime

start_time = datetime.now()

timestr = time.strftime("%Y%m%d-%H%M%S")

m = True
pagenumber = 1
bad_chars = ['\xa0','\n',' ']
compl_list = []
links = []
while m == True:  
    while pagenumber <=14:
        time.sleep(4)
        main_link = "https://www.avito.ru/"
        params = {'p':pagenumber} #//////////////////////
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36'}
        response = requests.get(main_link + 'volgograd/kvartiry/prodam/novostroyka-ASgBAQICAUSSA8YQAUDmBxSOUg',
                                headers=headers, params = params)
        avito_dom = html.fromstring(response.text)
        avito_links = avito_dom.xpath('//div[@class="iva-item-titleStep-pdebR"]//a/@href')
        pagenumber += 1
        for link_list in avito_links:
            news_dict = link_list
            links.append(main_link+news_dict)
        pprint(len(links))
    else:
        m = False
    pprint(len(links))
    for content in links:
        response_t = requests.get(content, headers=headers)
        time.sleep(3)
        dom2 = html.fromstring(response_t.text)
        container = dom2.xpath("//div[contains(@class,'item-view__new-style')]")
        for i in container:
            fil_dict = {}
            name = i.xpath("//span[contains(@class,'title-info-title-text')]/text()")
            number = i.xpath("//span[@data-marker='item-view/item-id']/text()")
            try:
                fil_dict['Номер'] = number[0].replace('№', '').replace('[', '').replace(']', '')
                fil_dict['url'] = content
                atribut = i.xpath("//span[@class='item-params-label']/text()")
                varib = i.xpath("//li[@class='item-params-list-item']/text()[last()]")
                price = i.xpath("//span[@class='js-item-price']/text()")
                fil_dict['Цена'] = int(price[0].replace('\xa0', ''))
                adress = i.xpath("//span[@class='item-address__string']/text()")
                fil_dict['Адрес'] = adress[0].replace('\xa0', '').replace('\n', '')
                raion = i.xpath("//span[@class='item-address-georeferences-item__content']/text()")
                fil_dict['Район'] = raion[0].replace('№', '').replace('[', '').replace(']', '')
                novostroika = i.xpath("//li[@class='item-params-list-item']/a/text()")
                fil_dict['Название новостройки'] = novostroika[0].replace('\xa0', '').replace('\n', '')
            except Exception:
                pprint('Все норм')
            for z in range(len(atribut)):
                atrib_dict = {atribut[z].replace(':', ''): varib[z].replace('\xa0', '').replace('\n', '').replace('м²','').replace('.',',')}
                fil_dict.update(atrib_dict)
            compl_list.append(fil_dict)
            pprint(len(compl_list))
df = pd.DataFrame(compl_list)
df = df[['url','Номер', 'Цена', 'Адрес', 'Район', 'Название новостройки', 'Количество комнат '
    , 'Общая площадь ', 'Площадь кухни ', 'Жилая площадь ', 'Этаж ', 'Балкон или лоджия ', 'Санузел '
    , 'Окна ', 'Отделка ', 'Корпус, строение ', 'Официальный застройщик ', 'Тип участия ', 'Срок сдачи ', 'Тип дома ', 'Этажей в доме '
    , 'Пассажирский лифт ', 'Грузовой лифт ', 'Двор ', 'Парковка ', 'Тип комнат ', 'Высота потолков ', 'Способ продажи ', 'Вид сделки ']]
pd.options.mode.chained_assignment
df['Общая площадь '] = df['Общая площадь '].str.replace('([0-9]+)', r'\1', regex=True)

xlsx_file = 'Parser3.xlsx'


def update_spreadsheet(xlsx_file: str, df, starcol : int = 1,startrow : int = 2,sheet_name : str = 'Wokrsheet 1'):
    '''
    :param xlsx_file: Путь до шаблона файла Excel
    :param df: Датафрейм pandas для записи
    :param starcol: Стартовая колонка в таблице Excel, где будут перезаписываться данные
    :param startrow: Стартовая строка в таблице Excel, где будут перезаписываться данные
    :param sheet_name: Название страницы в Excel
    :return:
    '''
    wb = ox.load_workbook(xlsx_file)
    for ir in range(0, len(df)): # Перебираем двумерный массив, сначала строки потом серию
        for ic in range(0, len(df.iloc[ir])): # iloc позволяет выбрать конкретную ячейку
            wb[sheet_name].cell(startrow + ir, starcol + ic).value = df.iloc[ir][ic] # Присваиваем ячейке значения по заданным координатам датафрейма
            wb.save(timestr+'.xlsx') # Сохраняем изменения

update_spreadsheet(xlsx_file, df, sheet_name="Worksheet 1", starcol=1, startrow= 2)
end_time = datetime.now()
print('Duration: {}'.format(end_time - start_time))