import openpyxl
import requests
from bs4 import BeautifulSoup


def read_excel_file(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    products = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        asin = row[0]
        url = f'https://www.amazon.co.uk/dp/{asin}'
        products.append({'asin': asin, 'url': url})

    return products


def check_availability(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, 'lxml')

    out_of_stock = soup.find('span', {'class': 'a-size-medium a-color-price'})

    if out_of_stock and 'out of stock' in out_of_stock.text.lower():
        return 'yes'
    else:
        return 'no'


def update_excel_file(file_path, products):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    sheet.cell(row=1, column=2).value = 'URL'
    sheet.cell(row=1, column=3).value = 'Out of Stock'

    for index, product in enumerate(products, start=2):
        sheet.cell(row=index, column=2).value = product['url']
        sheet.cell(row=index, column=3).value = product['out_of_stock']

    workbook.save(file_path)


def main():
    file_path = 'C:/Users/gabriel.konopnicki/OneDrive - funko.com/input/list.xlsx'
    products = read_excel_file(file_path)

    for product in products:
        product['out_of_stock'] = check_availability(product['url'])

    update_excel_file(file_path, products)


if __name__ == '__main__':
    main()
