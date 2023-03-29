from howlongtobeatpy import HowLongToBeat
#credit: https://github.com/ScrappyCocco/HowLongToBeat-PythonAPI

import sys
import os
import openpyxl
from openpyxl import load_workbook

try:
    #search for xlsx files in directory?
    wb = load_workbook("SteamGames.xlsx")
    ws = wb["Sheet"]
    #ws.delete_cols(2,3) #deletes 3 columns starting at B
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
def completion_time(game, playstyle):
    if game is not None:
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
            print("Should be impossible (completion_time)")
            sys.exit()
        return time
    else:
        return "Not found in HLTB database."

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

# takes in list of games, returns list of times for the games
def getTimes(games):
    timeList = []
    x = input("Select the type of times you would like to retrieve:\n"
              "1. Main story completion only\n"
              "2. Main story + extra content\n"
              "3. Full completion\n"
              "4. Average of 1 and 2 (recommended)\n"
              "5. Average of 2 and 3\n"
              "6. Average of all\n\n"
              "If you would like multiple of the above, please re-run the script at the end.\n")
    try:
        x = int(x) #this only works if x is an int, filters input
    except:
        print("Invalid input.")
        getTimes(games)

    for i in range(0, gameCount):
        gameSearch = searchforGame(games[i])
        # print(gameSearch)
        timeList.append(completion_time(gameSearch, x))

    return timeList

def exportToSheet(times):
    y = input("Is this your first time running through this script? Y/N\n")
    if y == "N" or y == "n":
        x = input("Select the Excel column you would like to put the times in. If you would like to enter in column C, please enter 3, etc.\n")
    else:
        x = "2"
    try:
        x = int(x)
    except:
        print("Invalid column input.")
        exportToSheet(times)
    for i in range(0, len(times)):
        ws_cell = ws.cell(row=i + 1, column=x)
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
