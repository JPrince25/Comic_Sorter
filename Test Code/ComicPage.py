import pandas as pd
import requests
from bs4 import BeautifulSoup
import re

URL = "https://dc.fandom.com/wiki/Batman_Vol_2_8"
page = requests.get(URL) 
soup = BeautifulSoup(page.content, "html.parser")
f = soup.find('b', string=["Featured Characters:"])
featured = f.find_next('ul').find_all('a')
for link in featured:
    print(link.get('href'))
s = soup.find('b', string=["Supporting Characters:"])
supporting = soup.find_next('ul').find_all('a')
for link in supporting:
    print(link.get('href'))
#antagonists = soup.find('b', string=["Antagonists:"]).find_next('ul')