#modules needed: aiohttp, requests, fake_useragent, openpyxl
#scrape games via ParseHub

from howlongtobeatpy import HowLongToBeat #https://pypi.org/project/howlongtobeatpy/
from openpyxl import load_workbook #https://ehmatthes.github.io/pcc_2e/beyond_pcc/extracting_from_excel/

#class howlongtobeat():
file = "C:/Users/Nathan/Documents/PySteamTimer/SteamGames.xlsx"
wb = load_workbook(file)
ws = wb["run_results"]

def getGames(): #returns list of steam games from excel file
    rows = list(ws.rows)
    i = 0
    gameList = []
    while i < 734:
        for j in rows[i]:
            gameList.append(j.value)
        i+=1
    return gameList
    
def findGame(name): #returns an integer for hour completion
    gameList = getGames()
    results = HowLongToBeat().search(name)
    if results_list is not None and len(results_list) > 0:
        best_element = max(results_list, key=lambda element: element.similarity)
        
































def scratch():
    rows = list(ws.rows)
    i = 0
    j = 0
    gameList = []
    while True:
        for j in rows[i]:
            if i.value == "":
                return gameList
            gameList.append(i.value)
        i+=1
    return gameList


    rows = list(ws.rows)
    i = 0
    gameList = []
    while i < 734:
        for j in rows[i]:
            gameList.append(j.value)
        i+=1
    return gameList
