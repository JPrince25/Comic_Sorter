import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
from pandas import ExcelWriter

class Titles():
    def __init__(self,name):
        self.name = name
        self.monthlyTitleSales = [0]*147
        self.monthlyTitleUnits = [0]*147

        self.issueNumber = []
        self.volNumber = []
        self.issueSales = []
        self.issueUnits = []
        self.mainCharacters = []
    
    def addTitleSale(self,sale,col):
        self.monthlyTitleSales[col] = self.monthlyTitleSales[col] + sale
    
    def addTitleUnit(self,unit,col):
        self.monthlyTitleUnits[col] = self.monthlyTitleUnits[col] + unit
        
    def addUnits(self,unit):
        self.issueUnits.append(unit)

    def addSales(self,sale):
        self.issueSales.append(sale)

    def addIssueNumber(self,number):
        self.issueNumber.append(number)
        
    def addCharacter(self,character):
        self.mainCharacters.append(character)
    
    def addVol(self,vol):
        self.volNumber.append(vol)
  

    
    
marvel = 'Marvel'
dc = 'DC'

year = 2008
month = 1
monthC = "01"
monthCol = 0

monthList = []
comicTitle = ""
comicUnits = 0
comicSales = 0
comicExpiration = []       

issueCount = []

dcTitleList = []
dcCharacterString = []
dcURL = "https://dc.fandom.com/wiki/"
marvelURL = "https://marvel.fandom.com/wiki/"

#marvelNames = pd.read_excel('MarvelNames.xlsx', sheet_name= "Titles", names=['DC','Track','Super','DataBaseTitle','Group','Group Name','Vol 1 Month Start','Vol 1 Year Start','Vol 1 Issue Start','Vol 1 Month End','Vol 1 Year End','Vol 1 Issue End','Vol 2 Month Start','Vol 2 Year Start','Vol 2 Issue Start','Vol 2 Month End','Vol 2 Year End','Vol 2 Issue End','Vol 3 Month Start','Vol 3 Year Start','Vol 3 Issue Start','Vol 3 Month End','Vol 3 Year End','Vol 3 Issue End','Vol 4 Month Start','Vol 4 Year Start','Vol 4 Issue Start','Vol 4 Month End','Vol 4 Year End','Vol 4 Issue End','Vol 5 Month Start','Vol 5 Year Start','Vol 5 Issue Start','Vol 5 Month End','Vol 5 Year End','Vol 5 Issue End','Vol 6 Month Start','Vol 6 Year Start','Vol 6 Issue Start','Vol 6 Month End','Vol 6 Year End','Vol 6 Issue End','Vol 7 Month Start','Vol 7 Year Start','Vol 7 Issue Start','Vol 7 Month End','Vol 7 Year End','Vol 7 Issue End','Vol 8 Month Start','Vol 8 Year Start','Vol 8 Issue Start','Vol 8 Month End','Vol 8 Year End','Vol 8 Issue End','Vol 9 Month Start','Vol 9 Year Start','Vol 9 Issue Start','Vol 9 Month End','Vol 9 Year End','Vol 9 Issue End','Vol 10 Month Start','Vol 10 Year Start','Vol 10 Issue Start','Vol 10 Month End','Vol 10 Year End','Vol 10 Issue End','Vol 11 Month Start','Vol 11 Year Start','Vol 11 Issue Start','Vol 11 Month End','Vol 11 Year End','Vol 11 Issue End','Vol 12 Month Start','Vol 12 Year Start','Vol 12 Issue Start','Vol 12 Month End','Vol 12 Year End','Vol 12 Issue End'])
dcNames = pd.read_excel('DCNames.xlsx', sheet_name= "Titles", names=['DC','Track','Super','DataBaseTitle','Group','Group Name','Vol 1 Month Start','Vol 1 Year Start','Vol 1 Issue Start','Vol 1 Month End','Vol 1 Year End','Vol 1 Issue End','Vol 2 Month Start','Vol 2 Year Start','Vol 2 Issue Start','Vol 2 Month End','Vol 2 Year End','Vol 2 Issue End','Vol 3 Month Start','Vol 3 Year Start','Vol 3 Issue Start','Vol 3 Month End','Vol 3 Year End','Vol 3 Issue End','Vol 4 Month Start','Vol 4 Year Start','Vol 4 Issue Start','Vol 4 Month End','Vol 4 Year End','Vol 4 Issue End','Vol 5 Month Start','Vol 5 Year Start','Vol 5 Issue Start','Vol 5 Month End','Vol 5 Year End','Vol 5 Issue End','Vol 6 Month Start','Vol 6 Year Start','Vol 6 Issue Start','Vol 6 Month End','Vol 6 Year End','Vol 6 Issue End','Vol 7 Month Start','Vol 7 Year Start','Vol 7 Issue Start','Vol 7 Month End','Vol 7 Year End','Vol 7 Issue End','Vol 8 Month Start','Vol 8 Year Start','Vol 8 Issue Start','Vol 8 Month End','Vol 8 Year End','Vol 8 Issue End','Vol 9 Month Start','Vol 9 Year Start','Vol 9 Issue Start','Vol 9 Month End','Vol 9 Year End','Vol 9 Issue End','Vol 10 Month Start','Vol 10 Year Start','Vol 10 Issue Start','Vol 10 Month End','Vol 10 Year End','Vol 10 Issue End','Vol 11 Month Start','Vol 11 Year Start','Vol 11 Issue Start','Vol 11 Month End','Vol 11 Year End','Vol 11 Issue End','Vol 12 Month Start','Vol 12 Year Start','Vol 12 Issue Start','Vol 12 Month End','Vol 12 Year End','Vol 12 Issue End'])


while (year < 2020 or month!=4):
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
    print("-------------------------------------------------------------------------------------------------------------------------------\n\n")
    print( str(year)+"_"+monthC)
    print("\n\n-------------------------------------------------------------------------------------------------------------------------------")
    monthList.append(str(year)+"_"+monthC)
    issueData = pd.read_excel('Issues.xlsx',sheet_name = str(year)+"_"+monthC, names = ['Amt','Dollars','Comic-Title','Issue','Price','Publisher','Units','Sales'])
    issueData.reset_index(inplace = True, drop = True)
    marvelIssueData = issueData.apply(lambda r: r.astype('string').str.contains(marvel).any(), axis=1)
    dcIssueData = issueData.apply(lambda r: r.astype('string').str.contains(dc).any(), axis=1)
    

    # for index, row in issueData[marvelIssueData].iterrows():
    #    if(index > 300):
    #       break
    #    comicTitle = row['Comic-Title']
    #    sales = row['Sales']
    #    units = row['Units']
    #    issueNum = row['Issue']
    #    if (issueNum not in issueCount):
    #        issueCount.append(issueNum)
    #        print(issueNum)
    
    for index, row in issueData[dcIssueData].iterrows():
        if(index > 300):
            break
        comicTitle = row['Comic-Title']
        sales = row['Sales']
        units = row['Units']
        issueString = row['Issue']
        dcTitleWikiURL = dcURL
        vol = 1
        tracked = False
        for index, nameRow in dcNames.iterrows():
            if comicTitle in nameRow['DC']:
                if (nameRow['Track'] == 1 and nameRow['Super'] == 1) or (nameRow['Track'] == 3):
                    tracked = True
                    if (nameRow['Track'] == 1):
                        issueNum = int(issueString)
                    else:
                        issueNum = float(issueString)
                    dcTitleWikiURL = dcTitleWikiURL + nameRow['DataBaseTitle'] + "_Vol_" 
                    comicTitle = nameRow['DataBaseTitle']
                    if year <= nameRow['Vol 1 Year End'] and year >= nameRow['Vol 1 Year Start']:
                        if issueNum <= float(nameRow['Vol 1 Issue End']) and issueNum >= float(nameRow['Vol 1 Issue Start']):
                            if (year == int(nameRow['Vol 1 Year Start']) and month >= int(nameRow['Vol 1 Month Start'])):
                                dcTitleWikiURL = dcTitleWikiURL + "1_" + str(issueNum) 
                                vol = 1
                            elif (year == int(nameRow['Vol 1 Year End']) and month <= int(nameRow['Vol 1 Month End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "1_" + str(issueNum) 
                                vol = 1
                            elif (year > int(nameRow['Vol 1 Year Start']) and year < int(nameRow['Vol 1 Year End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "1_" + str(issueNum) 
                                vol = 1
                    if year <= nameRow['Vol 2 Year End'] and year >= nameRow['Vol 2 Year Start']:
                        if issueNum <= float(nameRow['Vol 2 Issue End']) and issueNum >= float(nameRow['Vol 2 Issue Start']):
                            if (year == int(nameRow['Vol 2 Year Start']) and month >= int(nameRow['Vol 2 Month Start'])):
                                dcTitleWikiURL = dcTitleWikiURL + "2_" + str(issueNum) 
                                vol = 2
                            elif (year == int(nameRow['Vol 2 Year End']) and month <= int(nameRow['Vol 2 Month End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "2_" + str(issueNum) 
                                vol = 2
                            elif (year > int(nameRow['Vol 2 Year Start']) and year < int(nameRow['Vol 2 Year End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "2_" + str(issueNum) 
                                vol = 2
                    if year <= nameRow['Vol 3 Year End'] and year >= nameRow['Vol 3 Year Start']:
                        if issueNum <= int(nameRow['Vol 3 Issue End']) and issueNum >= int(nameRow['Vol 3 Issue Start']):
                            if (year == int(nameRow['Vol 3 Year Start']) and month >= int(nameRow['Vol 3 Month Start'])):
                                dcTitleWikiURL = dcTitleWikiURL + "3_" + str(issueNum) 
                                vol = 3
                            elif (year == int(nameRow['Vol 3 Year End']) and month <= int(nameRow['Vol 3 Month End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "3_" + str(issueNum) 
                                vol = 3
                            elif (year > int(nameRow['Vol 3 Year Start']) and year < int(nameRow['Vol 3 Year End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "3_" + str(issueNum) 
                                vol = 3
                    if year <= nameRow['Vol 4 Year End'] and year >= nameRow['Vol 4 Year Start']:
                        if issueNum <= int(nameRow['Vol 4 Issue End']) and issueNum >= int(nameRow['Vol 4 Issue Start']):
                            if (year == int(nameRow['Vol 4 Year Start']) and month >= int(nameRow['Vol 4 Month Start'])):
                                dcTitleWikiURL = dcTitleWikiURL + "4_" + str(issueNum) 
                                vol = 4
                            elif (year == int(nameRow['Vol 4 Year End']) and month <= int(nameRow['Vol 4 Month End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "4_" + str(issueNum) 
                                vol = 4
                            elif (year > int(nameRow['Vol 4 Year Start']) and year < int(nameRow['Vol 4 Year End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "4_" + str(issueNum) 
                                vol = 4
                    if year <= nameRow['Vol 5 Year End'] and year >= nameRow['Vol 5 Year Start']:
                        if issueNum <= int(nameRow['Vol 5 Issue End']) and issueNum >= int(nameRow['Vol 5 Issue Start']):
                            if (year == int(nameRow['Vol 5 Year Start']) and month >= int(nameRow['Vol 5 Month Start'])):
                                dcTitleWikiURL = dcTitleWikiURL + "5_" + str(issueNum) 
                                vol = 5
                            elif (year == int(nameRow['Vol 5 Year End']) and month <= int(nameRow['Vol 5 Month End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "5_" + str(issueNum) 
                                vol = 5
                            elif (year > int(nameRow['Vol 5 Year Start']) and year < int(nameRow['Vol 5 Year End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "5_" + str(issueNum) 
                                vol = 5
                    if year <= nameRow['Vol 6 Year End'] and year >= nameRow['Vol 6 Year Start']:
                        if issueNum <= int(nameRow['Vol 6 Issue End']) and issueNum >= int(nameRow['Vol 6 Issue Start']):
                            if (year == int(nameRow['Vol 6 Year Start']) and month >= int(nameRow['Vol 6 Month Start'])):
                                dcTitleWikiURL = dcTitleWikiURL + "6_" + str(issueNum) 
                                vol = 6
                            elif (year == int(nameRow['Vol 6 Year End']) and month <= int(nameRow['Vol 6 Month End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "6_" + str(issueNum) 
                                vol = 6
                            elif (year > int(nameRow['Vol 6 Year Start']) and year < int(nameRow['Vol 6 Year End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "6_" + str(issueNum) 
                                vol = 6
                    if year <= nameRow['Vol 7 Year End'] and year >= nameRow['Vol 7 Year Start']:
                        if issueNum <= int(nameRow['Vol 7 Issue End']) and issueNum >= int(nameRow['Vol 7 Issue Start']):
                            if (year == int(nameRow['Vol 7 Year Start']) and month >= int(nameRow['Vol 7 Month Start'])):
                                dcTitleWikiURL = dcTitleWikiURL + "7_" + str(issueNum) 
                                vol = 7
                            elif (year == int(nameRow['Vol 7 Year End']) and month <= int(nameRow['Vol 7 Month End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "7_" + str(issueNum) 
                                vol = 7
                            elif (year > int(nameRow['Vol 7 Year Start']) and year < int(nameRow['Vol 7 Year End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "7_" + str(issueNum) 
                                vol = 7
                    if year <= nameRow['Vol 8 Year End'] and year >= nameRow['Vol 8 Year Start']:
                        if issueNum <= int(nameRow['Vol 8 Issue End']) and issueNum >= int(nameRow['Vol 8 Issue Start']):
                            if (year == int(nameRow['Vol 8 Year Start']) and month >= int(nameRow['Vol 8 Month Start'])):
                                dcTitleWikiURL = dcTitleWikiURL + "8_" + str(issueNum) 
                                vol = 8
                            elif (year == int(nameRow['Vol 8 Year End']) and month <= int(nameRow['Vol 8 Month End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "8_" + str(issueNum) 
                                vol = 8
                            elif (year > int(nameRow['Vol 8 Year Start']) and year < int(nameRow['Vol 8 Year End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "8_" + str(issueNum) 
                                vol = 8
                    if year <= nameRow['Vol 9 Year End'] and year >= nameRow['Vol 9 Year Start']:
                        if issueNum <= int(nameRow['Vol 9 Issue End']) and issueNum >= int(nameRow['Vol 9 Issue Start']):
                            if (year == int(nameRow['Vol 9 Year Start']) and month >= int(nameRow['Vol 9 Month Start'])):
                                dcTitleWikiURL = dcTitleWikiURL + "9_" + str(issueNum) 
                                vol = 9
                            elif (year == int(nameRow['Vol 9 Year End']) and month <= int(nameRow['Vol 9 Month End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "9_" + str(issueNum) 
                                vol = 9
                            elif (year > int(nameRow['Vol 9 Year Start']) and year < int(nameRow['Vol 9 Year End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "9_" + str(issueNum) 
                                vol = 9
                    if year <= nameRow['Vol 10 Year End'] and year >= nameRow['Vol 10 Year Start']:
                        if issueNum <= int(nameRow['Vol 10 Issue End']) and issueNum >= int(nameRow['Vol 10 Issue Start']):
                            if (year == int(nameRow['Vol 10 Year Start']) and month >= int(nameRow['Vol 10 Month Start'])):
                                dcTitleWikiURL = dcTitleWikiURL + "10_" + str(issueNum) 
                                vol = 10
                            elif (year == int(nameRow['Vol 10 Year End']) and month <= int(nameRow['Vol 10 Month End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "10_" + str(issueNum) 
                                vol = 10
                            elif (year > int(nameRow['Vol 10 Year Start']) and year < int(nameRow['Vol 10 Year End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "10_" + str(issueNum) 
                                vol = 10
                    if year <= nameRow['Vol 11 Year End'] and year >= nameRow['Vol 11 Year Start']:
                        if issueNum <= int(nameRow['Vol 11 Issue End']) and issueNum >= int(nameRow['Vol 11 Issue Start']):
                            if (year == int(nameRow['Vol 11 Year Start']) and month >= int(nameRow['Vol 11 Month Start'])):
                                dcTitleWikiURL = dcTitleWikiURL + "11_" + str(issueNum) 
                                vol = 11
                            elif (year == int(nameRow['Vol 11 Year End']) and month <= int(nameRow['Vol 11 Month End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "11_" + str(issueNum) 
                                vol = 11
                            elif (year > int(nameRow['Vol 11 Year Start']) and year < int(nameRow['Vol 11 Year End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "11_" + str(issueNum) 
                                vol = 11
                    if year <= nameRow['Vol 12 Year End'] and year >= nameRow['Vol 12 Year Start']:
                        if issueNum <= int(nameRow['Vol 12 Issue End']) and issueNum >= int(nameRow['Vol 12 Issue Start']):
                            if (year == int(nameRow['Vol 12 Year Start']) and month >= int(nameRow['Vol 12 Month Start'])):
                                dcTitleWikiURL = dcTitleWikiURL + "12_" + str(issueNum) 
                                vol = 12
                            elif (year == int(nameRow['Vol 12 Year End']) and month <= int(nameRow['Vol 12 Month End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "12_" + str(issueNum) 
                                vol = 12
                            elif (year > int(nameRow['Vol 12 Year Start']) and year < int(nameRow['Vol 12 Year End'])):
                                dcTitleWikiURL = dcTitleWikiURL + "12_" + str(issueNum) 
                                vol = 12
                    if (nameRow['Track'] == 3):
                        extension = nameRow['Group Name']
                        dcTitleWikiURL = dcTitleWikiURL + str(extension)
                    break     
        if (tracked): 
            if comicTitle not in dcCharacterString:
                newTitle = Titles(comicTitle)
                dcTitleList.append(newTitle)
                dcCharacterString.append(comicTitle)
                newTitle.addTitleSale(sales,monthCol)
                newTitle.addTitleUnit(units,monthCol)
                newTitle.addUnits(units)
                newTitle.addSales(sales)
                newTitle.addIssueNumber(issueString)
                newTitle.addVol(vol)
                print(comicTitle+" Added")
                try:     
                    print("Begin Search:"+dcTitleWikiURL)
                    page = requests.get(dcTitleWikiURL) 
                    soup = BeautifulSoup(page.content, "html.parser")
                    characterLinks = []
                    characterString = ""
                    for f in soup.find_all('b', string=["Featured Characters:"]):
                        featured = f.find_next('ul').find_all('a')
                        for link in featured:
                            if (link.get('href') not in characterLinks):
                                characterLinks.append(link.get('href'))
                    print("Charater Search:"+dcTitleWikiURL) 
                    for character in characterLinks:
                        dcCharacterURL = 'https://dc.fandom.com' + character
                        characterPage = requests.get(dcCharacterURL)
                        characterSoup = BeautifulSoup(characterPage.content, "html.parser")
                        characterName = characterSoup.find('h1').text.strip()
                        characterString = characterString + "/" + characterName
                    newTitle.addCharacter(characterString)
                except:
                    dcTitle.addCharacter("Character Not Found")
                    print(str(month) +" / " + str(year))
                    print(comicTitle)
                    print(issueString) 
            else:
                for dcTitle in dcTitleList:
                    if comicTitle in dcTitle.name:
                        dcTitle.addTitleSale(sales,monthCol)
                        dcTitle.addTitleUnit(units,monthCol)
                        dcTitle.addUnits(units)
                        dcTitle.addSales(sales)
                        dcTitle.addIssueNumber(issueString)
                        dcTitle.addVol(vol)
                        try:
                            print("Begin Search:"+dcTitleWikiURL)     
                            page = requests.get(dcTitleWikiURL) 
                            soup = BeautifulSoup(page.content, "html.parser")
                            characterLinks = []
                            characterString = ""
                            for f in soup.find_all('b', string=["Featured Characters:"]):
                                featured = f.find_next('ul').find_all('a')
                                for link in featured:
                                    if (link.get('href') not in characterLinks):
                                        characterLinks.append(link.get('href'))
                            print("Charater Search:"+dcTitleWikiURL) 
                            for character in characterLinks:
                                dcCharacterURL = 'https://dc.fandom.com' + character
                                characterPage = requests.get(dcCharacterURL)
                                characterSoup = BeautifulSoup(characterPage.content, "html.parser")
                                characterName = characterSoup.find('h1').text.strip()
                                characterString = characterString + "/" + characterName
                            dcTitle.addCharacter(characterString)
                            
                        except:
                            dcTitle.addCharacter("Character Not Found")
                            print(str(month) +" / " + str(year))
                            print(comicTitle)
                            print(issueString) 
                        print(comicTitle+" Updated")
            
                            
    monthCol = monthCol + 1
    if (month == 12):
        year = year + 1
        month = 1
    else: 
        month = month + 1

countTotal = 0
dataTrack = []
for finalTitle in dcTitleList:
    finalDisplay = []
    finalDisplay.append(finalTitle.monthlyTitleUnits)
    finalDisplay.append(finalTitle.monthlyTitleSales)
    finalDisplay.append(finalTitle.issueUnits)
    finalDisplay.append(finalTitle.issueSales)
    finalDisplay.append(finalTitle.volNumber)
    finalDisplay.append(finalTitle.issueNumber)
    finalDisplay.append(finalTitle.mainCharacters)
    dcDataFrame = pd.DataFrame(finalDisplay, index = ["Monthly Units","Monthly Sales","Issue Units","Issue Sales","Vol Number","Issue Number","Characters"])#
    try:
        dcDataFrame.to_excel(str(finalTitle.name)+"Info.xlsx")
    except:
        print(finalTitle.name+" Failed")



