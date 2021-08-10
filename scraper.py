#from openpyxl.workbook import Workbook
import requests
import datetime
import os
from openpyxl import Workbook, load_workbook
import lxml.html as html


HOME_URL = 'https://github.com/MinCiencia/Datos-COVID19/blob/master/output/producto1/Covid-19.csv'

XPATH_HEADERS = '//tr[@id="LC1"]/th/text()'


def home():
    try:
        response = requests.get(HOME_URL)

        if response.status_code == 200:
            home = response.content.decode('utf-8')
            parsed = html.fromstring(home)
            comuna = input('Ingresa comuna: ')
            comuna = '//tr[contains(.,"' + comuna + '")]/td/text()'
            # print(comuna)
            headers = parsed.xpath(XPATH_HEADERS)
            data = parsed.xpath(comuna)
            today = datetime.date.today().strftime('%d-%m-%Y')
            wb = Workbook()
            ws = wb.active
            ws.title = "reporte " + today
            # for header in headers:
            ws.append(headers)
            wb.save(f'reporte {today}.xlsx')
            # if not os.path.isdir(f'reporte ', today):
            #   os.mkdir(today)

            #parse_news(urlhome+link, today)

        else:
            raise ValueError(f'Error: {response.status_code}')
    except ValueError as e:
        print(e)


def run():
    home()


if __name__ == '__main__':
    run()
