import pandas as pd

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
        issueNum = row['Issue']
        if (issueNum not in issueCount):
            issueCount.append(issueNum)
            print(comicTitle)
            print(issueNum)
            
    
    monthCol = monthCol + 1
    if (month == 12):
        year = year + 1
        month = 1
    else: 
        month = month + 1
    



