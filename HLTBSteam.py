#modules needed: aiohttp, requests, fake_useragent, openpyxl, xlwt, steamfront
#scrape games via ParseHub

import steamfront #https://steamfront.readthedocs.io/en/latest/code-reference.html
import xlwt
from xlwt import Workbook #https://www.geeksforgeeks.org/writing-excel-sheet-using-python/
from howlongtobeatpy import HowLongToBeat #https://pypi.org/project/howlongtobeatpy/
from openpyxl import load_workbook #https://ehmatthes.github.io/pcc_2e/beyond_pcc/extracting_from_excel/

file = "C:/Users/Nathan/Documents/PySteamTimer/SteamGames.xlsx"
wb = load_workbook(file)
ws = wb["run_results"] #raw input in main?
rows = list(ws.rows)
gameCount = ws.max_row


steamClient = steamfront.Client()






#takes in JSON, returns completion time as int
def completion_time(game):
    if game is not None:
        type(game)
        return game.main_story
    else:
        return 0
        #return ((game.main_story + game.main_extra) // 2)
    

#returns list of steam games from excel file
def getGames(): 
    gameList = []
    for i in range(0,gameCount):
        #try:
        for j in rows[i]:
            gameList.append(j.value)
    return gameList

#takes in list of games, returns list of times for the games
def getTimes(games):
    timeList = []
    for i in range(0,gameCount):
        gameSearch = searchforGame(games[i])
        #print(gameSearch)
        timeList.append(completion_time(gameSearch))
    return timeList

def exportToSheet(times):
    for i in range(0,len(times)):
        ws.cell(row=i+1, column=2).value = times[i]
        #ws.write(i,1,times[i])
    wb.save('SteamGameTimes_Finished.xlsx')

#takes in a string, returns a JSON result to be interpreted
def searchforGame(name): 
    results = HowLongToBeat().search(name, similarity_case_sensitive=False)
    best_element = ""
    if results is not None and len(results) > 0:
        best_element = max(results, key=lambda element: element.similarity)
        print(name, "found")
        return best_element
    else:
        print(name, "NOT found. Check the name listed in your spreadsheet and ensure it does not include extraneous info such as Definitive Edition.")

#takes in JSON, returns name of game as string
def getName(game):
    return game.game_name 


class HLTBSteam():

    def main():
        gameList = getGames()
        gameTimes = getTimes(gameList)
        for i in range(0,gameCount):
            print(gameList[i], end=": ")
            print(gameTimes[i], "hours")
        exportToSheet(gameTimes)
            
    if __name__ == "__main__":
        main()


    


