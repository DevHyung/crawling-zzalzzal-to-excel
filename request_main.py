import test
exit(-1)
import requests
from bs4 import BeautifulSoup

if __name__=="__main__":
    html = requests.get('http://www.zzalzzal.com/gogo/upbit')
    bs4 = BeautifulSoup(html.text,'lxml')
    div = bs4.find('div',class_='tab-content')
    five_minute = div.find('div',id='5m').find('table',id='go1')
    fifteen_minute = div.find('div',id='15m').find('table',id='go3')
    thirty_minute = div.find('div', id='30m').find('table', id='go4')
    sixty_minute = div.find('div', id='60m').find('table', id='go5')
    print( len(five_minute.find_all('tr')))
    print(len(fifteen_minute.find_all('tr')))