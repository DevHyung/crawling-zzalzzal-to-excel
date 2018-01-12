"""
title           :requests_main.py
description     :
author          :DevHyung
date            :2018.01.11
version         :1.0.0
python_version  :3.6
required module :requests, lxml, BeautifulSoup4, XlsxWriter
develope env    : Mac OS X High Sierra, intel i5 (2Ghz), using cpu
[ref]
1.

"""
import requests
from bs4 import BeautifulSoup
import xlsxwriter
def makeExcel():
    # Create an new Excel file and add a worksheet.
    global workbook
    workbook = xlsxwriter.Workbook('demo.xlsx')
    worksheet = workbook.add_worksheet()

    # Widen the first column to make the text clearer.
    worksheet.set_column('A:A', 20)

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})

    # Write some simple text.
    worksheet.write('A1', 'Hello')

    # Text with formatting.
    worksheet.write('A2', 'World', bold)

    # Write some numbers, with row/column notation.
    worksheet.write(2, 0, 123)
    worksheet.write(3, 0, 123.456)

    workbook.close()
if __name__=="__main__":
    makeExcel()
    print(workbook)
    exit(-1)

    html = requests.get('http://www.zzalzzal.com/gogo/upbit')
    bs4 = BeautifulSoup(html.text,'lxml')
    div = bs4.find('div',class_='tab-content')
    five_minute = div.find('div',id='5m').find('table',id='go1')
    fifteen_minute = div.find('div',id='15m').find('table',id='go3')
    thirty_minute = div.find('div', id='30m').find('table', id='go4')
    sixty_minute = div.find('div', id='60m').find('table', id='go5')
    for tr in five_minute.find_all('tr')[1:]:
        tdlist =tr.find_all('td')
        print(tdlist[0].get_text(),":",tdlist[1].get_text().strip())
    #print( len(five_minute.find_all('tr')),five_minute.find_all('tr')[1])
    #print(len(fifteen_minute.find_all('tr')),fifteen_minute.find_all('tr')[1])
    #print(len(thirty_minute.find_all('tr')))
    #print(len(sixty_minute.find_all('tr')))