import pandas as pd

class CharacterComics():
    def __init__(self, name):
        self.name = name  
        self.comicList = []
        self.characterUnits = [0]*147 #Amount of Month
        self.characterSales = [0]*147
        self.expirations = []

    def __str__(self):
        return f"{self.name}"

    def addComic(self, comic):
        self.comicList.append(comic)
    
    def addSale(self,sales,col):
        self.characterSales[col] = self.characterSales[col] + sales
    
    def addUnit(self,units,col):
        self.characterUnits[col] = self.characterUnits[col] + units

    def addExpiration(self,expiration):
        self.expirations.append(expiration)

    def deleteComic(self,comic):
        self.comicList.remove(comic)

    def removeExpiration(self,expiration):
        self.expirations.remove(expiration)

class ExpirationComics():
    def __init__(self,comic,year,month):
        self.comic = comic
        self.year = year
        self.month = month



marvel = 'Marvel'
dc = 'DC'

year = 2008
month = 1
monthC = "01"
monthCol = 0

saveCheck = 0
saves = 1

marvelIssueCharacters = []
marvelIssueCharactersString = []



monthList = []
comicTitle = ""
comicUnits = 0
comicSales = 0
comicExpiration = []

while (month == 1 or month == 2 or month == 3):#(year < 2020 or month!=4):
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
    


   
    for index, row in issueData[marvelIssueData].iterrows():
        if(index > 300):
            break
        comicTitle = row['Comic-Title']
        sales = row['Sales']
        units = row['Units']
        if (len(marvelIssueCharacters) == 0):
            characterToAdd = input("What Should I Do With: " + comicTitle + "  /  "+str(sales)+"\n>>>")
            marvelIssueCharacters.append(CharacterComics(characterToAdd))
            marvelIssueCharactersString.append(characterToAdd)
        for marvelCharacter in marvelIssueCharacters:
            for expiration in marvelCharacter.expirations:
                if(month == int(expiration.month) and year == int(expiration.year)):
                    marvelCharacter.deleteComic(expiration.comic)
                    marvelCharacter.removeExpiration(expiration)
                    
        comicDetection = 0
        for marvelCharacter in marvelIssueCharacters:
            if (comicTitle in marvelCharacter.comicList):
                marvelCharacter.addSale(sales,monthCol)
                print(marvelCharacter.characterSales[monthCol])
                comicDetection = comicDetection + 1
        if (comicDetection == 0):
            commandControl = 0
            expiration = ExpirationComics(comicTitle,0,0)
            expirationCheck = False
            while(commandControl == 0):
                indexNum = 0
                characterList = "\n"
                for marvelCharacter in marvelIssueCharacters:
                    characterList = characterList + '_' + str(marvelCharacter)
                    if (indexNum % 6 == 5):
                        characterList = characterList + '\n'
                print("Known Characters: "+characterList+"\n")
                userInput = input("What Should I Do With: " + comicTitle + "  /  "+str(sales)+"\nOptions: O(One Character) N(New Character) M(Multiple Characters) E(Expiration)\n>>>")
                if (userInput in ['e','E']):
                    yearInput = input('What Year Will the Comic Expire:\n>>>')
                    monthInput = input('What Month Will the Comic Expire:\n>>>')
                    try:
                        expiration.year = yearInput
                    except:
                        print('You Did Not Enter A Valid Year')
                    try:
                        expiration.month = monthInput
                        expirationCheck = True
                        print(expiration.comic+' / '+expiration.year+' / '+expiration.month+' / '+str(expirationCheck))
                        print('Expiration Validated')
                    except:
                        print('You Did Not Enter A Valid Month')
                    
                if (userInput in ['o','O']):
                    knownCharacterControl = 0
                    while(knownCharacterControl == 0):
                        knownInput = input("Which Character: " + comicTitle + "\n>>>")
                        if (knownInput in ["?"]):
                            knownCharacterControl = 1
                        if (knownInput in marvelIssueCharactersString):
                            marvelIssueCharacters[marvelIssueCharactersString.index(knownInput)].addComic(comicTitle)
                            marvelIssueCharacters[marvelIssueCharactersString.index(knownInput)].addUnit(units,monthCol)
                            marvelIssueCharacters[marvelIssueCharactersString.index(knownInput)].addSale(sales,monthCol)
                            if (expirationCheck):
                                marvelIssueCharacters[marvelIssueCharactersString.index(knownInput)].addExpiration(expiration)
                            knownCharacterControl = 1
                            commandControl = 1
                            print(str(marvelIssueCharacters[marvelIssueCharactersString.index(knownInput)])+ " - Comic Added")
                            print(marvelIssueCharacters[marvelIssueCharactersString.index(knownInput)].characterSales[monthCol])
                            print("")
                        else:
                            print("Unknown Character")
                elif (userInput in ['n', 'N']):
                    newCharacterControl = 0
                    while(newCharacterControl == 0):
                        newInput = input("What New Character Should Be Added: "+ comicTitle + "\n>>>")
                        check = input("Is This Name Correct: "+newInput+"\n>>>")
                        if (check in ["?"]):
                            newCharacterControl = 1
                        if (check in ['y','Y']):
                            marvelIssueCharacters.append(CharacterComics(newInput))
                            marvelIssueCharactersString.append(newInput)
                            marvelIssueCharacters[marvelIssueCharactersString.index(newInput)].addComic(comicTitle)
                            marvelIssueCharacters[marvelIssueCharactersString.index(newInput)].addUnit(units,monthCol)
                            marvelIssueCharacters[marvelIssueCharactersString.index(newInput)].addSale(sales,monthCol)
                            if (expirationCheck):
                                marvelIssueCharacters[marvelIssueCharactersString.index(newInput)].addExpiration(expiration)
                            print(str(marvelIssueCharacters[marvelIssueCharactersString.index(newInput)])+ " - Character/Comic Added")
                            print(marvelIssueCharacters[marvelIssueCharactersString.index(newInput)].characterSales[monthCol])
                            print("")
                            newCharacterControl = 1
                            commandControl = 1
                elif (userInput in ['m','M']):
                    multipleCharacterControl = 0
                    multipleCharacterList = "Current Characters in Comic: "
                    while(multipleCharacterControl == 0):
                        characterAmount = input("How Many Characters: "+ comicTitle + "\n>>>")
                        i = 0
                        try:
                            while (i < int(characterAmount)):
                                mCC = 0
                                while (mCC == 0):
                                    print(multipleCharacterList)
                                    print("Character "+str(i+1)+": "+"\n>>>")
                                    characterKnowledge = input("Is the character (K)nown or (U)nknown: "+comicTitle+"\n>>>")
                                    if (characterKnowledge in ["?"]):
                                        mCC = 1
                                        multipleCharacterControl = 1
                                    if(characterKnowledge in ['k','K']):
                                        mKnownCharacterControl = 0
                                        while(mKnownCharacterControl == 0):
                                            mKnownInput = input("Which Character: " + comicTitle + "\n>>>")
                                            if (mKnownInput in ["?"]):
                                                mKnownCharacterControl = 1
                                            if (mKnownInput in marvelIssueCharactersString):
                                                marvelIssueCharacters[marvelIssueCharactersString.index(mKnownInput)].addComic(comicTitle)
                                                marvelIssueCharacters[marvelIssueCharactersString.index(mKnownInput)].addUnit(units,monthCol)
                                                marvelIssueCharacters[marvelIssueCharactersString.index(mKnownInput)].addSale(sales,monthCol)
                                                if (expirationCheck):
                                                    marvelIssueCharacters[marvelIssueCharactersString.index(mKnownInput)].addExpiration(expiration)
                                                multipleCharacterList = multipleCharacterList + "_"+mKnownInput
                                                print(str(marvelIssueCharacters[marvelIssueCharactersString.index(mKnownInput)])+ " - Comic Added")
                                                print(marvelIssueCharacters[marvelIssueCharactersString.index(mKnownInput)].characterSales[monthCol])
                                                print("")

                                                mKnownCharacterControl = 1
                                                mCC = 1
                                                multipleCharacterControl = 1
                                                commandControl = 1
                                            else:
                                                print("Unknown Character")
                                    if(characterKnowledge in ['u','U']):
                                        mNewCharacterControl = 0
                                        while(mNewCharacterControl == 0):
                                            mNewInput = input("What New Character Should Be Added: "+ comicTitle + "\n>>>")
                                            check = input("Is This Name Correct: "+mNewInput+"\n>>>")
                                            if (mNewInput in ["?"]):
                                                mNewCharacterControl = 1
                                            if (check in ['y','Y']):
                                                marvelIssueCharacters.append(CharacterComics(mNewInput))
                                                marvelIssueCharactersString.append(mNewInput)
                                                multipleCharacterList = multipleCharacterList + "_"+mNewInput
                                                marvelIssueCharacters[marvelIssueCharactersString.index(mNewInput)].addComic(comicTitle)
                                                marvelIssueCharacters[marvelIssueCharactersString.index(mNewInput)].addUnit(units,monthCol)
                                                marvelIssueCharacters[marvelIssueCharactersString.index(mNewInput)].addSale(sales,monthCol)
                                                if (expirationCheck):
                                                    marvelIssueCharacters[marvelIssueCharactersString.index(mNewInput)].addExpiration(expiration)
                                                    print(marvelIssueCharacters[marvelIssueCharactersString.index(mNewInput)].expirations[0].comic)
                                                    print('Unknown Expiration')
                                                print(str(marvelIssueCharacters[marvelIssueCharactersString.index(mNewInput)])+ " - Character/Comic Added")
                                                print(marvelIssueCharacters[marvelIssueCharactersString.index(mNewInput)].characterSales[monthCol])
                                                print("")
                                                mNewCharacterControl = 1
                                                mCC = 1
                                                multipleCharacterControl = 1
                                                commandControl = 1
                                i = i + 1    
                        except:
                               print("You Did Not Type in a Number") 
                else:
                    print('Unknown Command. Please Type in a Correct Command\n')
            #Save State
            if (saveCheck == 11):
                saveMarvelSalesData = []
                saveMarvelUnitsData = []
                saveMarvelName = []
                for marvelCharacter in marvelIssueCharacters:
                    saveMarvelSalesData.append(marvelCharacter.characterSales)
                    saveMarvelUnitsData.append(marvelCharacter.characterSales)
                    saveMarvelName.append(str(marvelCharacter.name))
                marvelDataFrame = pd.DataFrame( saveMarvelSalesData,columns=monthList, index = saveMarvelName)
                marvelDataFrame.to_excel("SAVE_"+str(saves)+"_MarvelIssueCharacterData.xlsx")


    print(issueData[marvelIssueData])
    monthCol = monthCol + 1
    if (month == 12):
        year = year + 1
        month = 1
        saveCheck = 0
        saves = saves + 1
    else: 
        month = month + 1
        saveCheck = saveCheck + 1


marvelIssueSalesData = []
marvelIssueUnitsData = []
marvelName = []
for marvelCharacter in marvelIssueCharacters:
    marvelIssueSalesData.append(marvelCharacter.characterSales)

    marvelName.append(str(marvelCharacter.name))
marvelDataFrame = pd.DataFrame(marvelNovelSalesData,columns=monthList, index = marvelName)
marvelDataFrame.to_excel("MarvelIssueCharacterSalesData.xlsx")

