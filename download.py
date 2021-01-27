from bs4 import BeautifulSoup
import re
import urllib.request


html_page = urllib.request.urlopen("your url")
soup = BeautifulSoup(html_page, features="lxml")
for link in soup.findAll('a'):
    if len(link.get('href')) < 14 and 'apc2011' in link.get('href'):
        print(link.get('href'))
        urllib.request.urlretrieve(
            'xmlurl/dailyxml/applications/' +
            link.get('href'),
            link.get('href'))
