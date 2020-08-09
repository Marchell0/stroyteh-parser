import os
import re
from pprint import pprint
from typing import Generator

import requests
from bs4 import BeautifulSoup as bs
from fake_useragent import UserAgent
from openpyxl import load_workbook


base_dir = os.path.dirname(__file__)
file = base_dir + '\\for_parsing.xlsx'
wb = load_workbook(file)
sheet = wb.get_sheet_by_name('only_product')


def read_xlsx() -> Generator:
    """Читаем xlsx файл"""
    row_count = sheet.max_row
    for row in range(2, row_count + 1):
        url = sheet['A' + str(row)].value
        yield url, row


def get_user_agent() -> dict:
    """Генерим user agent"""
    ua = UserAgent()
    header = {'user-agent': ua.random}
    return header


def get_response(url: str, header: dict) -> requests.Response:
    """Получем http(s) ответ от сервера"""
    r = requests.get(url, headers=header, verify=False)
    if r.ok:
        return r
    else:
        print(f'Ошибка {r.status_code}')


def parse_data(response) -> dict:
    """Сбор необходимых данных"""
    html = response.text
    soup = bs(html, 'lxml')
    url = response.url
    print(url)
    try:
        price_sourse = soup.find('div', class_='price').text.strip()
        # Валюта
        price_currency = price_sourse.split()[-1]
        current_price = soup.find('span', itemprop='price').text
        price = current_price + ' ' + price_currency
    except AttributeError:
        price = ''
    try:
        old_price = soup.find('div', class_='old-price').text.strip()
    except AttributeError:
        old_price = ''
    try:
        description = soup.find('div', class_='col-md-6').text.strip()
        description = re.sub(r' +', ' ', description)
        description = re.sub(r' \n', '\n', description)
        description = re.sub(r'\n ', '\n', description)
    except AttributeError:
        description = ''
    try:
        ware_cod = soup.find('span', class_='sku cod').text.strip()
        cod = re.findall(r'\d+', ware_cod)
    except AttributeError:
        cod = ''
    try:
        image_url = soup.find('meta', property='og:image')['content']
    except AttributeError:
        image_url = ''
    
    site_data = {
        'price': price,
        'old_price': old_price,
        'description': description,
        'ware_cod': cod,
        'image_url': image_url,
    }
    pprint(site_data)
    return site_data


def write_xlsx(data: dict, row: int):
    """Записываем данные в эксель файл"""
    pass


def main():
    user_agent = get_user_agent()
    xlsx_data = read_xlsx()
    for url, row in xlsx_data:
        response = get_response(url, user_agent)
        site_data = parse_data(response)
        write_xlsx(site_data, row)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(e)
