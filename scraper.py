import requests
import datetime
import os
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers
from openpyxl.chart import (
    LineChart,
    Reference,
)
from openpyxl.chart.axis import DateAxis
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
            if comuna == 'Arica':
                comunaxpath = '//tr[@id="LC2"]/td/text()'
            elif comuna == 'Antofagasta':
                comunaxpath = '//tr[@id="LC15"]/td/text()'
            elif comuna == 'Coquimbo':
                comunaxpath = '//tr[@id="LC38"]/td/text()'
            elif comuna == 'Valparaiso':
                comunaxpath = '//tr[@id="LC85"]/td/text()'
            elif comuna == 'Maule':
                comunaxpath = '//tr[@id="LC188"]/td/text()'
            elif comuna == "Talca":
                comunaxpath = '//tr[@id="LC202"]/td/text()'
            elif comuna == 'Aysén':
                comunaxpath = '//tr[@id="LC341"]/td/text()'
            else:
                comunaxpath = '//tr[contains(.,"' + comuna + '")]/td/text()'

            headers = parsed.xpath(XPATH_HEADERS)
            print(f'path: {comunaxpath}')
            data = parsed.xpath(comunaxpath)
            print(f'data: {data}')
            today = datetime.date.today().strftime('%d-%m-%Y')
            wb = Workbook()
            ws = wb.active
            ws.title = "reporte " + today
            print(len(data))
            for i in range(0, 5):
                _ = ws.cell(column=i+1, row=1, value=headers[i])
                _ = ws.cell(column=i+1, row=2, value=data[i])

            # _ = ws.cell(column=1, row=4, value="Fecha")
            # _ = ws.cell(column=1, row=5, value="Casos Totales")
            # _ = ws.cell(column=1, row=6, value="Casos Diarios")
            ws['A4'] = 'Fecha'
            ws['B4'] = 'Casos Totales'
            ws['C4'] = 'Casos Diarios'

            # for i in range(5, len(data)-1):
            print(len(data)-2)
            for i in range(len(data)-2, 4, -1):
                _ = ws.cell(column=1, row=len(data)-i+4, value=headers[i])

            for i in range(len(data)-2, 4, -1):
                _ = ws.cell(column=1, row=len(data)-i+3, value=headers[i])
                _ = ws.cell(column=2, row=len(data)-i +
                            3, value=int(float(data[i])))

            for i in range(len(data)-2, 4, -1):
                _ = ws.cell(column=3, row=len(data)-i+3, value='=B' +
                            str(len(data)-i+3)+'-B'+str(len(data)-i+4))
            # for i in range(1, len(headers)):
            for i in range(1, 6):
                ws.column_dimensions[get_column_letter(i)].auto_size = True

            chart = LineChart()
            chart.title = "Grafico"
            chart.style = 13
            chart.y_axis.title = "Casos diarios"
            chart.x_axis.title = "Fecha"
            chart.x_axis.number_format = 'dd-mm-yyyy'
            chart.height = 15
            chart.width = 30

            values = Reference(ws, min_col=3, min_row=len(data),
                               max_col=3, max_row=5)
            chart.add_data(values, titles_from_data=True)

            dates = Reference(ws, min_col=1, min_row=5,
                              max_col=1, max_row=len(data)+4)
            chart.set_categories(dates)
            ws.add_chart(chart, "E4")
            wb.save(f'Reporte {comuna} {today}.xlsx')

        else:
            raise ValueError(f'Error: {response.status_code}')
    except ValueError as e:
        print(e)


def run():
    home()


if __name__ == '__main__':
    run()
