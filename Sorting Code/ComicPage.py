import pandas as pd
import requests
from bs4 import BeautifulSoup
import re

URL = "https://dc.fandom.com/wiki/Batman_Vol_2_8"
page = requests.get(URL) 
soup = BeautifulSoup(page.content, "html.parser")
characterLinks = []
for f in soup.find_all('b', string=["Featured Characters:","Supporting Characters:","Antagonists:"]):
    print(f)
    featured = f.find_next('ul').find_all('a')
    for link in featured:
        if (link.get('href') not in characterLinks):
            characterLinks.append(link.get('href'))
print(characterLinks)
