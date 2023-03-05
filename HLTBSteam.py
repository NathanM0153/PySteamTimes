#modules needed: aiohttp, requests, fake_useragent, openpyxl, xlwt, steamfront


import steamfront
#https://steamfront.readthedocs.io/en/latest/code-reference.html
#https://github.com/4Kaylum/Steamfront
from howlongtobeatpy import HowLongToBeat
#https://pypi.org/project/howlongtobeatpy/

import xlwt
from xlwt import Workbook
#https://www.geeksforgeeks.org/writing-excel-sheet-using-python/
from openpyxl import load_workbook
#https://ehmatthes.github.io/pcc_2e/beyond_pcc/extracting_from_excel/




steamID64 = "76561198272854176"
steamAPIKey = "6D5F599D289CCFC212F5D824F212CCB4"
steamClient = steamfront.Client(steamAPIKey)


def appIDList(games):
    IDList = []
    for i in range(0,len(games)):
        IDList.append(games[i].appid)
        #print(games[i].appid)
    #print(IDList)
    return IDList

def gameNameList(IDList):
    games = []
    app = ""
    for i in IDList:
        app = steamClient.getApp(appid=i)
        games.append(app.name)
        #print(str(i))
    print(games)
    return games
    

#takes in JSON, returns completion time as int
def completion_time(game):
    if game is not None:
        time = (game.main_story + game.main_extra) // 2
        return time
    else:
        return 0
    


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
        user = steamClient.getUser(id64=steamID64)
        games = user.apps
        idlist = appIDList(games)
        gamelist = gameNameList(idlist)
        print(gamelist)
        return 0
    
    if __name__ == "__main__":
        main()


    


