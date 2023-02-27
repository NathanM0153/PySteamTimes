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
    while i < 734: #game count, make this dynamic later
        for j in rows[i]:
            gameList.append(j.value)
        i+=1
    return gameList
    
def findGame(name): #returns a JSON result to be interpreted
    results = HowLongToBeat().search(name, similarity_case_sensitive=False)
    if results is not None and len(results) > 0:
        best_element = max(results, key=lambda element: element.similarity)
        return best_element
    else:
        print("Not found")
        return

def parse():
    game = "Darkest Dungeon"
    json = findGame(game)
    result = __parse_web_result(game,json,None,False)
    return result
    































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
