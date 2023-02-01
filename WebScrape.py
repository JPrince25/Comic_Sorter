import requests
from bs4 import BeautifulSoup
import pandas as pd


year = 2000
month = 1
monthC = "01"
while (year < 2008):
    if (month == 1):
        monthC = "01"
    if (month == 2):
        monthC = "02"
    if (month == 3):
        monthC = "03"
    if (month == 4):
        monthC = "04"
    if (month == 5):
        monthC = "05"
    if (month == 6):
        monthC = "06"
    if (month == 7):
        monthC = "07"
    if (month == 8):
        monthC = "08"
    if (month == 9):
        monthC = "09"
    if (month == 10):
        monthC = "10"
    if (month == 11):
        monthC = "11"
    if (month == 12):
        monthC = "12"

    URL = "https://www.comichron.com/monthlycomicssales/"+str(year)+"/"+str(year)+"-"+monthC+".html"
    page = requests.get(URL) 
    soup = BeautifulSoup(page.content, "html.parser")     

    #Issues
    iName = str(year)+"_"+monthC
    issues = soup.find('table', {'id':"Top300Comics"})
    columns = [i.get_text(strip=True) for i in issues.find_all("th")]
    data = []
    for tr in issues.find("tbody").find_all("tr"):
        data.append([td.get_text(strip=True) for td in tr.find_all("td")])
    df = pd.DataFrame(data, columns=columns)
    with pd.ExcelWriter('IssuesPre.xlsx', mode = "a") as writer:  
        df.to_excel(writer, sheet_name = iName, index=False)

    #Novels
    gName = str(year)+"_"+monthC
    try:  
        novels = soup.find('table', {'id':"Top300GNs"})
        columns = [i.get_text(strip=True) for i in novels.find_all("th")]
        data = []
        for tr in novels.find("tbody").find_all("tr"):
            data.append([td.get_text(strip=True) for td in tr.find_all("td")])
        df = pd.DataFrame(data, columns=columns)
        with pd.ExcelWriter('GraphicNovelsPre.xlsx', mode ="a") as writer:  
            df.to_excel(writer, sheet_name= gName, index=False)
    except AttributeError: 
        print(gName+" Not Available")

    if (month == 12):
        year = year + 1
        month = 1
    else: 
        month = month + 1
