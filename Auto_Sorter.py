import pandas as pd
import requests
from bs4 import BeautifulSoup

class CharacterComics():
    def __init__(self, name):
        self.name = name  
        self.comicList = []
        self.characterUnits = [0]*147 #Amount of Month
        self.characterSales = [0]*147
        self.comicsIn = [0]*147
        self.weightedIn = [0]*147

    def __str__(self):
        return f"{self.name}"

    def addComic(self, comic):
        self.comicList.append(comic)
    
    def addSale(self,sales,col):
        self.characterSales[col] = self.characterSales[col] + sales
    
    def addUnit(self,units,col):
        self.characterUnits[col] = self.characterUnits[col] + units
    
    def addComicIn(self, comicIn, col):
        self.comicsIn[col] = self.comicsIn[col] + comicIn
    
    def addWeightedIn(self, weightIn, col):
        self.weightedIn[col] = self.weightedIn[col] + weightIn
        
class Titles():
    def __init__(self,name):
        self.name = name
        self.titleSales = [0]*147
        self.titleUnits = [0]*147
        self.issueNumber = []
        self.issueSales = []
        self.issueUnits = []
    
    def addTitleSale(self,sale,col):
        self.titleSales[col] = self.titleSales[col] = sale
    
    def addTitileUnit(self,unit,col):
        self.titleUnits[col] = self.titleUnits[col] = unit
        
    def addIssueNumber(self,number):
        self.issueNumber.append(number)
        
    def addIssueSale(self,sale):
        self.issueSales.append(sale)
    
    def addIssueUnit(self,unit):
        self.issueUnits.append(unit)
    
    
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

dcURL = "https://dc.fandom.com/wiki/"
marvelURL = "https://marvel.fandom.com/wiki/"

marvelNames = pd.read_excel('MarvelNames.xlsx', sheet_name= "Titles", names=['DC','Track','Super','DataBaseTitle','Group','Group Name','Vol 1 Month Start','Vol 1 Year Start','Vol 1 Issue Start','Vol 1 Month End','Vol 1 Year End','Vol 1 Issue End','Vol 2 Month Start','Vol 2 Year Start','Vol 2 Issue Start','Vol 2 Month End','Vol 2 Year End','Vol 2 Issue End','Vol 3 Month Start','Vol 3 Year Start','Vol 3 Issue Start','Vol 3 Month End','Vol 3 Year End','Vol 3 Issue End','Vol 4 Month Start','Vol 4 Year Start','Vol 4 Issue Start','Vol 4 Month End','Vol 4 Year End','Vol 4 Issue End','Vol 5 Month Start','Vol 5 Year Start','Vol 5 Issue Start','Vol 5 Month End','Vol 5 Year End','Vol 5 Issue End','Vol 6 Month Start','Vol 6 Year Start','Vol 6 Issue Start','Vol 6 Month End','Vol 6 Year End','Vol 6 Issue End','Vol 7 Month Start','Vol 7 Year Start','Vol 7 Issue Start','Vol 7 Month End','Vol 7 Year End','Vol 7 Issue End','Vol 8 Month Start','Vol 8 Year Start','Vol 8 Issue Start','Vol 8 Month End','Vol 8 Year End','Vol 8 Issue End','Vol 9 Month Start','Vol 9 Year Start','Vol 9 Issue Start','Vol 9 Month End','Vol 9 Year End','Vol 9 Issue End','Vol 10 Month Start','Vol 10 Year Start','Vol 10 Issue Start','Vol 10 Month End','Vol 10 Year End','Vol 10 Issue End','Vol 11 Month Start','Vol 11 Year Start','Vol 11 Issue Start','Vol 11 Month End','Vol 11 Year End','Vol 11 Issue End','Vol 12 Month Start','Vol 12 Year Start','Vol 12 Issue Start','Vol 12 Month End','Vol 12 Year End','Vol 12 Issue End'])
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
    print( str(year)+"_"+monthC)
    monthList.append(str(year)+"_"+monthC)
    issueData = pd.read_excel('Issues.xlsx',sheet_name = str(year)+"_"+monthC, names = ['Amt','Dollars','Comic-Title','Issue','Price','Publisher','Units','Sales'])
    issueData.reset_index(inplace = True, drop = True)
    marvelIssueData = issueData.apply(lambda r: r.astype('string').str.contains(marvel).any(), axis=1)
    dcIssueData = issueData.apply(lambda r: r.astype('string').str.contains(dc).any(), axis=1)
    

    #for index, row in issueData[marvelIssueData].iterrows():
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
        if comicTitle in ['Batman']:
            try: 
                issueNum = int(issueString)
                test = issueNum / 2
            except:
                issueNum = int(input(issueString + "\n>>>"))
            print(comicTitle)
            print(issueString)
            for index, nameRow in dcNames.iterrows():
                if comicTitle in nameRow['DC']:
                    if nameRow['Track'] == 1 and nameRow['Super'] == 1:
                        dcTitleWikiURL = dcTitleWikiURL + nameRow['DataBaseTitle'] + "_Vol_" 
                        if year <= nameRow['Vol 1 Year End'] and year >= nameRow['Vol 1 Year Start']:
                            if issueNum <= float(nameRow['Vol 1 Issue End']) and issueNum >= float(nameRow['Vol 1 Issue Start']):
                                if year == float(nameRow['Vol 1 Year Start'] and month >= int(nameRow['Vol 1 Month Start'])):
                                    dcTitleWikiURL = dcTitleWikiURL + "1_" + str(issueNum) 
                                    print (dcTitleWikiURL)
                                elif year == float(nameRow['Vol 1 Year End'] and month <= int(nameRow['Vol 1 Month End'])):
                                    dcTitleWikiURL = dcTitleWikiURL + "1_" + str(issueNum) 
                                    print (dcTitleWikiURL)
                                else:
                                    dcTitleWikiURL = dcTitleWikiURL + "1_" + str(issueNum) 
                                    print (dcTitleWikiURL)
                        if year <= nameRow['Vol 2 Year End'] and year >= nameRow['Vol 2 Year Start']:
                            if issueNum <= float(nameRow['Vol 2 Issue End']) and issueNum >= float(nameRow['Vol 2 Issue Start']):
                                if year == float(nameRow['Vol 2 Year Start'] and month >= int(nameRow['Vol 2 Month Start'])):
                                    dcTitleWikiURL = dcTitleWikiURL + "2_" + str(issueNum) 
                                    print (dcTitleWikiURL)
                                elif year == float(nameRow['Vol 2 Year End'] and month <= int(nameRow['Vol 2 Month End'])):
                                    dcTitleWikiURL = dcTitleWikiURL + "2_" + str(issueNum) 
                                    print (dcTitleWikiURL)
                                else:
                                    dcTitleWikiURL = dcTitleWikiURL + "2_" + str(issueNum) 
                                    print (dcTitleWikiURL)
                        if year <= nameRow['Vol 3 Year End'] and year >= nameRow['Vol 3 Year Start']:
                            if issueNum <= float(nameRow['Vol 3 Issue End']) and issueNum >= float(nameRow['Vol 3 Issue Start']):
                                if year == float(nameRow['Vol 3 Year Start'] and month >= int(nameRow['Vol 3 Month Start'])):
                                    dcTitleWikiURL = dcTitleWikiURL + "3_" + str(issueNum) 
                                    print (dcTitleWikiURL)
                                if year == float(nameRow['Vol 3 Year End'] and month <= int(nameRow['Vol 3 Month End'])):
                                    dcTitleWikiURL = dcTitleWikiURL + "3_" + str(issueNum) 
                                    print (dcTitleWikiURL)
                                else:
                                    dcTitleWikiURL = dcTitleWikiURL + "3_" + str(issueNum) 
                                    print (dcTitleWikiURL)
                        issuePage = requests.get(URL) 
                        soup = BeautifulSoup(page.content, "html.parser") 
                        
                    break  
    
    monthCol = monthCol + 1
    if (month == 12):
        year = year + 1
        month = 1
    else: 
        month = month + 1
    



