import sys
import openpyxl
import os
import glob
#from os import listdir
#from os.path import isfile, join
from howlongtobeatpy import HowLongToBeat
#credit: https://github.com/ScrappyCocco/HowLongToBeat-PythonAPI



from pathlib import Path
#print(Path.cwd())
#C:\Users\Nathan\Documents\PySteamTimer


def findFile():
    excelFilesinDir = []
    cwd = Path.cwd()
    os.chdir(cwd)
    for i in glob.glob("*.xlsx"):
        excelFilesinDir.append(i)
    if len(excelFilesinDir) == 1:
        return excelFilesinDir[0]
    elif len(excelFilesinDir) == 0:
        print("Please run importSteam.py first.")
        sys.exit()
    else:
        print("Multiple .xlsx files detected:")
        for i in range(0,len(excelFilesinDir)):
            print(i+1, " ", excelFilesinDir[i])
        x = input("Enter the listed number of the one you would like to use.\n")
        try:
            x = int(x)
            if x < 1:
                print("Invalid input.")
                findFile()
            print(excelFilesinDir[x-1], "selected.\n")
            return excelFilesinDir[x-1]
        except Exception as e:
            print("Invalid input:", e)
            findFile()

try:
    #search for xlsx files in directory?
    #!!!
    wb = openpyxl.load_workbook(findFile())
    ws = wb["Sheet"]
    wsList = wb.sheetnames
    gameCount = ws.max_row
except Exception as e:
    print("A problem occurred:", e)
    #print("If your error is not apparent, please check the README.txt file for assistance.")
    sys.exit()

# pulls games from the excel sheet, puts in a list
def getGames(): 
    gameList = []
    first_col = list(ws['A'])
    #try and get first_col in here instead of the whole data file
    #print(rows)
    try:
        for i in first_col:
            gameList.append(i.value)
    except Exception as e:
        print(e)
        print("Game count invalid. Check if your sheet is empty.")
        sys.exit()
    return gameList

# takes in JSON, returns completion time as int
def completion_time(game, playstyle):
    time = 0
    if game is not None:
        try:
            if playstyle == 1:
                time = round(game.main_story, 2)
            elif playstyle == 2:
                time = round(game.main_extra, 2)
            elif playstyle == 3:
                time = round(game.completionist, 2)
            elif playstyle == 4:
                time = round((game.main_story + game.main_extra) / 2, 2)
            elif playstyle == 5:
                time = round((game.main_extra + game.completionist) / 2, 2)
            elif playstyle == 6:
                time = round((game.main_story + game.main_extra + game.completionist) / 3, 2)
            else:
                print("Time not found")
        except Exception as e:
            print(e)
            sys.exit()
        return time
    else:
        return "Not found in HLTB database."

# takes in list of games, returns list of times for the games
def getTimes(games):
    x = input("Select the type of times you would like to retrieve:\n"
              "1. Main story completion only\n"
              "2. Main story + extra content\n"
              "3. Full completion\n"
              "4. Average of 1 and 2 (recommended)\n"
              "5. Average of 2 and 3\n"
              "6. Average of all\n\n"
              "If you would like multiple of the above, please re-run the script at the end.\n")
    print("")
    try:
        x = int(x) #this only works if x is an int, filters input
        if x < 1 or x > 6:
            print("Invalid input.")
            getTimes(games)            
    except:
        print("Invalid input.")
        getTimes(games)

    result = searchResults(games, x)

    return result

def searchResults(games, x):
    timeList = []
    typeList = []
    #global searchResults
    #searchResults = []
    for i in range(0, gameCount):
        gameSearch = searchforGame(games[i])
        # print(gameSearch)
        #searchResults.append(gameSearch)
        timeList.append(completion_time(gameSearch, x))
        try:
            typeList.append(gameSearch.game_type)
        except:
            typeList.append("")
    return [timeList, typeList]



# takes in a string, returns a JSON result to be interpreted
def searchforGame(name):
    results = HowLongToBeat().search(name, similarity_case_sensitive=False)
    best_element = ""
    if results is not None and len(results) > 0:
        best_element = max(results, key=lambda element: element.similarity)
        print(name, "found")
        #count += 1
        #maybe someday put a number here
        return best_element
    else:
        print(name,  "NOT found. Check the README file for reasons this may be.")
        return None

# takes in JSON, returns name of game as string
def getName(game):
    return game.game_name
