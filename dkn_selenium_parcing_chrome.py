#Python3!
#Парсинг mos.ru и сохранение в csv

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from fake_useragent import UserAgent
import time
import datetime
import os
import csv


#Формирование списков CSS элементов
name_doc_CSS = []
name_date_start_CSS = []
name_date_finish_CSS = []
summary_CSS = 'body > div.b-fix-wrapper > div > div:nth-child(1) > div > div > div.mos-oiv-layout__wrapper > div > div.mos-oiv-layout__main-content > div > div.documents-projects-one-url.ng-scope.ng-isolate-scope > div > div > div > div.department-onedoc__attachs.ng-scope > div:nth-child(2) > a'
for i in range(1, 16):
    names_doc = 'body > div.b-fix-wrapper > div > div:nth-child(1) > div > div > div.mos-oiv-layout__wrapper > div > div.mos-oiv-layout__main-content > div > div.documents-projects-list.ng-scope.ng-isolate-scope > div > div.projects-list__items > div:nth-child' + '(' + str(i)+')' + ' > div > a'
    name_doc_CSS.append(names_doc)
    names_date_start = 'body > div.b-fix-wrapper > div > div:nth-child(1) > div > div > div.mos-oiv-layout__wrapper > div > div.mos-oiv-layout__main-content > div > div.documents-projects-list.ng-scope.ng-isolate-scope > div > div.projects-list__items > div:nth-child' + '(' + str(i)+')' + ' > div > div.mos-oiv-project__public-date > span.mos-oiv-project__text.ng-binding'
    name_date_start_CSS.append(names_date_start)
    names_date_finish = 'body > div.b-fix-wrapper > div > div:nth-child(1) > div > div > div.mos-oiv-layout__wrapper > div > div.mos-oiv-layout__main-content > div > div.documents-projects-list.ng-scope.ng-isolate-scope > div > div.projects-list__items > div:nth-child' + '(' + str(i)+')' + ' > div > div.mos-oiv-project__date-end > span.mos-oiv-project__text.ng-binding'
    name_date_finish_CSS.append(names_date_finish)


try:
    #Запрос последнего наименования csv
    File = open('Public_dicscussion_to_excel.csv', encoding='utf-8')
    Reader = csv.reader(File, delimiter='~')
    exampleData = list(Reader)
    first_name = exampleData[1][0]
    File.close()

    #Сбор и обновление данных в таблице
    useragent = UserAgent()
    option = webdriver.ChromeOptions()
    option.add_experimental_option('excludeSwitches', ['enable-logging'])
    option.add_argument(f'user-agent={useragent.random}')
    url = 'https://www.mos.ru/dkn/documents/discussions/?page=1'
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=option)
    driver.get(url=url)
    time.sleep(5)

    
    #Сбор данных со страниц и переключение
    name_docs = []
    date_start = []
    date_finish = []
    a = 1
    while first_name not in name_docs:
        for CSS_doc in name_doc_CSS:
            if driver.find_elements(By.CSS_SELECTOR, CSS_doc):
                if first_name not in name_docs:
                    name_doc = driver.find_element(By.CSS_SELECTOR, CSS_doc).text
                    name_docs.append(name_doc)
        for CSS_start in name_date_start_CSS:
            if driver.find_elements(By.CSS_SELECTOR, CSS_start):
                dates_start = driver.find_element(By.CSS_SELECTOR, CSS_start).text
                date_start.append(dates_start)
        for CSS_finish in name_date_finish_CSS:
            if driver.find_elements(By.CSS_SELECTOR, CSS_finish):
                dates_finish = driver.find_element(By.CSS_SELECTOR, CSS_finish).text
                date_finish.append(dates_finish)
        if a == 1:
            page_nxt = driver.find_element(By.CSS_SELECTOR, 'body > div.b-fix-wrapper > div > div:nth-child(1) > div > div > div.mos-oiv-layout__wrapper > div > div.mos-oiv-layout__main-content > div > div.documents-projects-list.ng-scope.ng-isolate-scope > div > div.projects-list__pagination > div > div > a > span > span')
            page_nxt.click()
            time.sleep(3)
            a += 1
        else:
            page_nxt = driver.find_element(By.CSS_SELECTOR, 'body > div.b-fix-wrapper > div > div:nth-child(1) > div > div > div.mos-oiv-layout__wrapper > div > div.mos-oiv-layout__main-content > div > div.documents-projects-list.ng-scope.ng-isolate-scope > div > div.projects-list__pagination > div > div > a.mos-oiv-pagination__link.mos-oiv-pagination__link_next.ng-scope > span > span')
            page_nxt.click()
            time.sleep(3)
    name_docs = name_docs[ :-1]        
    date_start = date_start[0 : len(name_docs)]
    date_finish = date_finish[0 : len(name_docs)]

    driver.close()
    driver.quit()
    # Добавление новых данных в csv
    if first_name not in name_docs and name_docs != []:
        rows = zip(name_docs,
                   date_start,
                   date_finish)
        with open('New_data_DKN.csv', 'w', newline="", encoding='utf-8') as file:
             writer = csv.DictWriter(file, ['Наименование документа', 'Дата начала обсуждения',
         'Дата окончания обсуждения', 'Сводка на сайте'], delimiter='~')
             writer.writeheader()
             for row in rows:
                writer.writerow({'Наименование документа': row[0], 
                'Дата начала обсуждения': row[1], 
                'Дата окончания обсуждения': row[2]})
        oldDatatoExcel = open('Public_dicscussion_to_excel.csv', encoding='utf-8')
        oldDatatoExcelRead = csv.DictReader(oldDatatoExcel, ['Наименование документа', 'Дата начала обсуждения',
         'Дата окончания обсуждения', 'Сводка на сайте'], delimiter='~')
        oldDatatoExcelRead.__next__()
        writertoexcell = open('New_data_DKN.csv', 'a', newline='', encoding='utf-8')
        outputwritertoexcel = csv.DictWriter(writertoexcell, ['Наименование документа', 'Дата начала обсуждения',
         'Дата окончания обсуждения', 'Сводка на сайте'], delimiter='~')
        for row in oldDatatoExcelRead:
            outputwritertoexcel.writerow({'Наименование документа' : row['Наименование документа'],'Дата начала обсуждения' : row['Дата начала обсуждения'], 
            'Дата окончания обсуждения': row['Дата окончания обсуждения'], 'Сводка на сайте': row['Сводка на сайте']})

        oldDatatoExcel.close()
        writertoexcell.close()


        os.remove('Public_dicscussion_to_excel.csv')
        os.rename('New_data_DKN.csv', 'Public_dicscussion_to_excel.csv')
        print("Список актов обновлён")
        time.sleep(4)

#Проверка наличия сводок в актах
    #Выгрузка обновлённого словаря данных
    Name1 = []
    date1 = []
    date2 = []
    summary = []
    exampleFile2 = open('Public_dicscussion_to_excel.csv', encoding='utf-8')
    exampleDictReader2 = csv.DictReader(exampleFile2, ['Наименование документа', 'Дата начала обсуждения',
         'Дата окончания обсуждения', 'Сводка на сайте'], delimiter='~')
    for row in exampleDictReader2:
        Name1.append(row['Наименование документа'])
        date1.append(row['Дата начала обсуждения'])
        date2.append(row['Дата окончания обсуждения'])
        summary.append(row['Сводка на сайте'])
    exampleFile2.close()

    #Список дат, сегодня минус 22 дня
    date_today = datetime.datetime.today()
    date_list = []
    for i in range(23): 
        date_today -= datetime.timedelta(days=1)
        date_list.append(date_today.strftime('%d.%m.%Y'))
    
    # Cписок страниц
    urls_list =[]
    for i in range(3, 13):
        url_scan = 'https://www.mos.ru/dkn/documents/discussions/?page=' + str(i)
        urls_list.append(url_scan) 

    #Акты со сводкой
    Akts_with_summary = []
    #Проверка перебором наличия сводок в актах
    def get_data(url):
        useragent = UserAgent()
        option = webdriver.ChromeOptions()
        option.add_experimental_option('excludeSwitches', ['enable-logging'])
        option.add_argument(f'user-agent={useragent.random}')
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=option)
        driver.get(url=url)
        time.sleep(5)
        for n, CSS_finish in enumerate(name_date_finish_CSS, 0):
            dates_finish = driver.find_element(By.CSS_SELECTOR, CSS_finish).text
            for date in date_list:
                if dates_finish == date:
                    akt_name = driver.find_element(By.CSS_SELECTOR, name_doc_CSS[n]).text
                    elementDetected = driver.find_element(By.CSS_SELECTOR, name_doc_CSS[n])
                    elementDetected.click()
                    time.sleep(3)
                    if driver.find_elements(By.CSS_SELECTOR, summary_CSS):
                        Akts_with_summary.append(akt_name)
                        driver.back()
                        time.sleep(2)
                        break            
                    else: 
                        driver.back()
                        time.sleep(2)
                        break
        for akt in Akts_with_summary:
            for N, name in enumerate(Name1, 0):
                if akt == name:
                    summary[N] = '+'
                    break                    

        return summary                        
        driver.close()
        driver.quit()

    list(map(get_data, urls_list))

    #Перезапись обновлённых элементов csv
    upgrade_web_pars = zip(Name1,
    date1,
    date2,
    summary)
    writertoexcell = open('Public_dicscussion_to_excel.csv', 'w', newline='', encoding='utf-8')
    outputwritertoexcel= csv.DictWriter(writertoexcell, ['Наименование документа', 'Дата начала обсуждения',
         'Дата окончания обсуждения', 'Сводка на сайте'], delimiter='~')
    for row in upgrade_web_pars:
        outputwritertoexcel.writerow({'Наименование документа': row[0], 
        'Дата начала обсуждения': row[1], 
        'Дата окончания обсуждения': row[2], 
        'Сводка на сайте': row[3]})
    writertoexcell.close()
    
    print("Программа завершена.")
    time.sleep(4)
        

except:
    useragent = UserAgent()
    option = webdriver.ChromeOptions()
    option.add_experimental_option('excludeSwitches', ['enable-logging'])
    option.add_argument(f'user-agent={useragent.random}')
    url = 'https://www.mos.ru/dkn/documents/discussions/?page=1'
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=option)
    driver.get(url=url)
    time.sleep(5)
    
    name_docs = []
    date_start = []
    date_finish = []
    CSS_end = 'body > div.b-fix-wrapper > div > div:nth-child(1) > div > div > div.mos-oiv-layout__wrapper > div > div.mos-oiv-layout__main-content > div > div.documents-projects-list.ng-scope.ng-isolate-scope > div > div.projects-list__pagination > div > div > a.mos-oiv-pagination__link.mos-oiv-pagination__link_next.ng-scope > span > span'
    a = 1
    while True:
        for CSS_doc in name_doc_CSS:
            if driver.find_elements(By.CSS_SELECTOR, CSS_doc):
                name_doc = driver.find_element(By.CSS_SELECTOR, CSS_doc).text
                name_docs.append(name_doc)
        for CSS_start in name_date_start_CSS:
            if driver.find_elements(By.CSS_SELECTOR, CSS_start):
                dates_start = driver.find_element(By.CSS_SELECTOR, CSS_start).text
                date_start.append(dates_start)
        for CSS_finish in name_date_finish_CSS:
            if driver.find_elements(By.CSS_SELECTOR, CSS_finish):
                dates_finish = driver.find_element(By.CSS_SELECTOR, CSS_finish).text
                date_finish.append(dates_finish)
        if not driver.find_elements(By.CSS_SELECTOR, 'body > div.b-fix-wrapper > div > div:nth-child(1) > div > div > div.mos-oiv-layout__wrapper > div > div.mos-oiv-layout__main-content > div > div.documents-projects-list.ng-scope.ng-isolate-scope > div > div.projects-list__pagination > div > div > a.mos-oiv-pagination__link.mos-oiv-pagination__link_next.ng-scope > span > span'):
            break
        if a == 1:
            page_nxt = driver.find_element(By.CSS_SELECTOR, 'body > div.b-fix-wrapper > div > div:nth-child(1) > div > div > div.mos-oiv-layout__wrapper > div > div.mos-oiv-layout__main-content > div > div.documents-projects-list.ng-scope.ng-isolate-scope > div > div.projects-list__pagination > div > div > a > span > span')
            page_nxt.click()
            time.sleep(3)
            a += 1
        else:
            page_nxt = driver.find_element(By.CSS_SELECTOR, 'body > div.b-fix-wrapper > div > div:nth-child(1) > div > div > div.mos-oiv-layout__wrapper > div > div.mos-oiv-layout__main-content > div > div.documents-projects-list.ng-scope.ng-isolate-scope > div > div.projects-list__pagination > div > div > a.mos-oiv-pagination__link.mos-oiv-pagination__link_next.ng-scope > span > span')
            page_nxt.click()
            time.sleep(3)
            a += 1
    #Копирование данных в csv
    date_start = date_start[0 : len(name_docs)]
    date_finish = date_finish[0 : len(name_docs)]
    print(str(a-1) + " страниц скопированно в папку с программой, формат - csv")
    print(name_docs)
    rows = zip(name_docs,
               date_start,
               date_finish)
    date_today = datetime.datetime.today()
    CSV_name = 'Public_dicscussion_to_excel.csv'
    with open(CSV_name, 'w', newline="", encoding='utf-8') as file:
         writer = csv.DictWriter(file, ['Наименование документа', 'Дата начала обсуждения',
     'Дата окончания обсуждения', 'Сводка на сайте'], delimiter='~')
         writer.writeheader() 
         for row in rows:
            writer.writerow({'Наименование документа': row[0], 
            'Дата начала обсуждения': row[1], 
            'Дата окончания обсуждения': row[2]})
    
    driver.close()
    driver.quit()
    print('Программа завершена')
    time.sleep(4)

