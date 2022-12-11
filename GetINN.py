'''
Данный модуль предназначен для компиляции в формат исполняемого файла или
динамической библиотеки и запуска из листа excel.
Читает инн из ячеек "G4 - G13" активного листа Excel,
по каждому из полученных инн запрашивает с сайта ЕГРЮЛ
информацию о лице, имеющем право действовать от именю организации
без доверенности. Записывает полученную информацию в стобец "K".
Данные с сайта скачиваются в формате pdf. Для сохранения загруженных файлов
используется папка temp_pdf. Для парсинга pdf используется tabula.
'''


import time
import requests

import pandas as pd
from win32com.client import Dispatch
from urllib import parse
import tabula

url = "https://egrul.nalog.ru/"
url_search_result = "https://egrul.nalog.ru/search-result/"
url_download_file = "https://egrul.nalog.ru/vyp-download/"

form = {
    "vyp3CaptchaToken": "",
    "page": "",
    "query": "",  # сюда нужно подставить инн
    "region": "",
    "PreventChromeAutocomplete": "",
}
headers = {
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"
}

pdf_dir = 'temp_pdf/'
excel_filename = 'Тестовое.xlsm'


def main():
    inn_list = _read_inn_from_excel()
    authorized_person_list = _get_authorized_person_list(inn_list)
    _write_inn_to_excel(authorized_person_list)


def _read_inn_from_excel():
    '''
    Читает данные из ячеек "G4 - G13" активного листа Excel,
    возвращает список значений, преобразованных в строку
    '''
    xl = Dispatch('Excel.Application')
    inn_list = []
    for i in range(4, 14):
        inn_list.append(xl.ActiveSheet.Cells(i, 7).Value)
    for i in range(0, len(inn_list)):
        inn_list[i] = str(inn_list[i])
        inn_list[i] = inn_list[i].split('.')[0]
    return inn_list


def _get_authorized_person_list(inn_list):
    '''
    Получает информацию с сайта по каждому инн в списке
    МЕЖДУ ИТЕРАЦИЯМИ УСТАНОВЛЕНА ЗАДЕРЖКА 0,5 СЕК, ЧТОБЫ НЕ РУГАЛАСЬ КАПЧА
    '''
    authorized_person_list = []
    for inn in inn_list:
        pdf_filename = _get_pdf_by_inn(inn)
        if pdf_filename:
            authorized_person_list.append(_get_inn_from_pdf(pdf_filename))
        else:
            authorized_person_list.append('Ошибка получения ИНН')
        time.sleep(0.5)
    return authorized_person_list


def _get_pdf_by_inn(inn):
    '''
    Ищет на сайте фирму по инн, скачивает выписку pdf,
    возвращает имя файла. Если скачать не удалось, возвращает пустую строку
    '''
    try:
        firm_code = _search_request(inn)
        download_link = _search_download_link_request(firm_code)
        response_file = _get_file_request(download_link)
        pdf_filename = _save_pdf(response_file)
    except:
        pdf_filename = ''
    return pdf_filename


def _search_request(inn):
    '''
    Отправляет поисковую форму на сайте, возвращает код для поиска организации
    '''
    form['query'] = inn
    data = parse.urlencode(form)
    response = requests.post(url=url, headers=headers, data=data)
    response_json = response.json()
    firm_code = response_json["t"]
    return firm_code


def _search_download_link_request(firm_code):
    '''
    По коду организации отправляет get запрос на получение информации со ссылкой на скачку выписки,
    возвращает код для скачивания pdf
    '''
    params = {
        "r": str(round(time.time(), 3)).replace(".", ""),
        "_": str(round(time.time(), 3)).replace(".", "")
    }

    response_search_result = requests.get(url=url_search_result + firm_code, params=params)
    response_search_result_json = response_search_result.json()
    download_link = response_search_result_json["rows"][0]["t"]
    return download_link


def _get_file_request(download_link):
    '''
    Запрашивает файл, возвращает ответ с файлом внутри
    '''
    response_download_file = requests.get(url=url_download_file + download_link)
    return response_download_file


def _save_pdf(response_download_file):
    '''
    Читает ответ. Если файл скачался, то сохраняет его в pdf_dir,
    возвращает имя файла. Если файл не скачался, возвращает пустую строку.
    Ну здесь, конечно, try - except на грани добра и зла
    '''
    try:
        pdf_filename = (response_download_file.headers.get("content-disposition")).split('filename=')[1]
        pdf_filename = pdf_filename.split('.pdf')[0]
        pdf_filename = pdf_dir + pdf_filename + '.pdf'
    except AttributeError:
        pdf_filename = ''

    if response_download_file.headers.get("Content-Type") == "application/pdf":
        with open(pdf_filename, 'wb') as f:
            f.write(response_download_file.content)
    return pdf_filename


def _get_inn_from_pdf(pdf_path):
    '''
    Читает файл выписки, возвращает инн организации, имеющей право действовать от имени юр. лица
    без доверенности.
    Сначала делает из pdf набор pd.DataFrame, потом соединяет таблицы из первых страниц pdf в одну,
    потом ищет в них данные
    '''
    try:
        tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True, lattice=True)
    except tabula.errors.JavaNotFoundError:
        print('Ошибка: для работы программы необходимо установить интерпретатор Java')
        print('Нажмите Enter, чтобы выйти из программы')
        input()
        exit()
    df_full_table = pd.DataFrame()
    for table in tables[1:4]:
        if not table.empty:
            table.columns = ['1', '2', '3']
            df_full_table = pd.concat([df_full_table, table], axis=0, ignore_index=True)
    return _get_inn_from_dataframe(df_full_table)


def _get_inn_from_dataframe(df: pd.DataFrame):
    '''
    Перебирает pd.DataFrame. Находит сначала строку, включающую подстроку
    "Сведения о лице, имеющем право без доверенности действовать от имени".
    После нее перебирает дальше и забирает первый попавшийся ИНН.
    '''
    i = 0
    while i <= df.shape[0]:
        if 'Сведения о лице, имеющем право без доверенности действовать от имени' in str(df.iloc[i, 0]):
            break
        else:
            i += 1
    while i <= df.shape[0]:
        if 'ИНН' == str(df.iloc[i, 1]):
            break
        else:
            i += 1
    if i == df.shape[0] + 1:
        return 'ИНН не найден в файле pdf'
    inn = str(df.iloc[i, 2])
    return inn


def _write_inn_to_excel(inn_list):
    '''
    Записывает данные из списка в активный лист Excel в диапазон "K4:K13"
    '''
    xl = Dispatch('Excel.Application')
    for i in range(4, 14):
        xl.ActiveSheet.Cells(i, 11).Value = inn_list[i - 4]


main()


