from howlongtobeatpy import HowLongToBeat
#credit: https://github.com/ScrappyCocco/HowLongToBeat-PythonAPI

import sys
import os
import openpyxl
from openpyxl import load_workbook


#filePath = "C:/Users/Nathan/Documents/PySteamTimer/SteamGames.xlsx"

try:
    #search for xlsx files in directory?
    #ignore column B on second run
    wb = load_workbook("SteamGames.xlsx")
    ws = wb["Sheet"]
    #ws = wb.active
    wsList = wb.sheetnames
    if len(wsList) > 1:
        print("These are the sheets in your excel file:")
        for i in range(0,len(wsList)):
            if i != (len(wsList)-1): #cuts off the comma on wsList[-1]
                print(wsList[i], end=", ")
            else:
                print(wsList[i])
        sheet = input("Choose the sheet you would like to query: ")
        ws = wb[sheet]
    gameCount = ws.max_row
except Exception as e:
    print("A problem occurred:", e)
    print("If your error is not apparent, please check the README.txt file for assistance.")
    sys.exit()

# takes in JSON, returns completion time as int
def completion_time(game):
    if game is not None:
        time = round((game.main_story + game.main_extra) / 2, 2)
        return time
    else:
        return "Not found in HLTB database."

# pulls games from the excel sheet, puts in a list
def getGames(): 
    gameList = []
    rows = list(ws.rows)
    try:
        for i in range(0,gameCount):
            for j in rows[i]:
                gameList.append(j.value)
    except:
        print("Game count invalid. Check if your sheet is empty.")
        sys.exit()
    return gameList

# takes in list of games, returns list of times for the games
def getTimes(games):
    timeList = []
    for i in range(0, gameCount):
        gameSearch = searchforGame(games[i])
        # print(gameSearch)
        timeList.append(completion_time(gameSearch))
    return timeList

def exportToSheet(times):
    for i in range(0, len(times)):
        ws_cell = ws.cell(row=i + 1, column=2)
        ws_cell.value = times[i]
        # ws.write(i,1,times[i])
    wb.save('SteamGames.xlsx')
    print("Exported to file successfully.")
    print("Refer to the README.txt in order to complete results.")

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

class HLTBSteam():
    def main():
        gameList = getGames()
        timeList = getTimes(gameList)
        count = 1
        #for i in range(0,gameCount):
            #print(gameList[i], end=": ")
            #time = timeList[i]
            #if time == "Not found in HLTB database.":
                #print(time)
            #else:
                #print(time, "hours")
        exportToSheet(timeList)
        return 0

    if __name__ == "__main__":
        main()
