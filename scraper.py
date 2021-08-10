import requests
import datetime
import os
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers
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
            comunaxpath = '//tr[contains(.,"' + comuna + '")]/td/text()'
            headers = parsed.xpath(XPATH_HEADERS)
            data = parsed.xpath(comunaxpath)
            today = datetime.date.today().strftime('%d-%m-%Y')
            wb = Workbook()
            ws = wb.active
            ws.title = "reporte " + today
            for i in range(0, 5):
                _ = ws.cell(column=i+1, row=1, value=headers[i])
                _ = ws.cell(column=i+1, row=2, value=data[i])

            _ = ws.cell(column=1, row=4, value="Fecha")
            _ = ws.cell(column=1, row=5, value="Casos Totales")
            _ = ws.cell(column=1, row=6, value="Casos Diarios")
            for i in range(5, len(data)):
                _ = ws.cell(column=i-3, row=4, value=headers[i])
                _ = ws.cell(column=i-3, row=5, value=int(float(data[i])))
            #_ = ws.cell(column=2, row=6, value='=SUM(B5:C5)')

            for i in range(1, len(headers)):
                ws.column_dimensions[get_column_letter(i)].auto_size = True

            wb.save(f'Reporte {comuna} {today}.xlsx')

        else:
            raise ValueError(f'Error: {response.status_code}')
    except ValueError as e:
        print(e)


def run():
    home()


if __name__ == '__main__':
    run()
