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
    print("")
    try:
        x = int(x) #this only works if x is an int, filters input
        if x < 1 or x > 6:
            print("Invalid input.")
            getTimes(games)            
    except:
        print("Invalid input.")
        getTimes(games)

    global searchResults
    searchResults = []
    for i in range(1, gameCount):
        gameSearch = searchforGame(games[i])
        # print(gameSearch)
        searchResults.append(gameSearch)
        timeList.append(completion_time(gameSearch, x))

    return timeList

def exportToSheet(times):
    typeList = []
    y = input("Is this your first time running through this script? Y/N\n")
    if y == "N" or y == "n":
        x = input("Select the Excel column you would like to put the times in. If you would like to enter in column C, please enter 3, etc. "
                  "Please keep in mind that the game list is in column A and the first set of times created was put in column B.\n")
        if x == "1":
            z = input("Are you sure? This will overwrite your games. Y/N\n")
            if z == "N" or z == "n":
                exportToSheet(times)
        multi = "N"
        endless = "N"
    else:
        x = "2"
        multi = input("Would you like to remove primarily multiplayer games? These are often given inflated completion times. Y/N\n")
        endless = input("Would you like to remove games marked as endless? These are often given arbitrary completion times. Y/N\n")

    try:
        x = int(x)
        if x < 1:
            print("Invalid column input.")
            exportToSheet(times) 
    except:
        print("Invalid column input.")
        exportToSheet(times)
        
        
    ### Insert into Excel ###
    for i in range(0, len(times)):
        ws_cell = ws.cell(row=i+1, column=x)
        ws_cell2 = ws.cell(row=i+1, column=x+1)
        try:
            ws_cell.value = times[i]
            ws_cell2.value = searchResults[i].game_type
        except Exception as e:
            print(searchResults)
            print(e)
            ws_cell2.value = ""
        typeList.append(searchResults[i].game_type)
    wb.save('SteamGames.xlsx')

    if (multi == "Y" or multi == "y"):
        #find all rows, sort, reverse, then delete
        for i in range(1,gameCount):
            if typeList[i] == "multi":
                ws.delete_rows(i,1)
    if endless == "Y" or endless == "y":     
        for i in range(1,gameCount):
            if typeList[i] == "endless":
                ws.delete_rows(i,1)
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
        exportToSheet(timeList)
        return 0

    if __name__ == "__main__":
        main()
