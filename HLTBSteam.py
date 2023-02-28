#modules needed: aiohttp, requests, fake_useragent, openpyxl
#scrape games via ParseHub

from howlongtobeatpy import HowLongToBeat #https://pypi.org/project/howlongtobeatpy/
from openpyxl import load_workbook #https://ehmatthes.github.io/pcc_2e/beyond_pcc/extracting_from_excel/

class HLTBSteam():
    file = "C:/Users/Nathan/Documents/PySteamTimer/SteamGames.xlsx"
    wb = load_workbook(file)
    ws = wb["run_results"]

    
    #returns list of steam games from excel file
    def getGames(): 
        rows = list(ws.rows)
        i = 0
        gameList = []
        #game count, make this dynamic later
        for i in range(0,734):
            for j in rows[i]:
                gameList.append(j.value)
        return gameList

    #takes in list of games, returns list of times for the games
    def getTimes(games):
        timeList = []
        for i in games:
            timeList += completion_time(searchforGame(games[i]))
        return timeList
        

    #returns a JSON result to be interpreted
    def searchforGame(name): 
        results = HowLongToBeat().search(name, similarity_case_sensitive=False)
        if results is not None and len(results) > 0:
            best_element = max(results, key=lambda element: element.similarity)
            return best_element
        else:
            print("Not found.")
            return None

    #takes in JSON, returns name of game as string
    def getName(game):
        return game.game_name

    
    #takes in JSON, returns completion time as int
    def completion_time(game):
        return game.main_story
        #return ((game.main_story + game.main_extra) // 2)

    def main():
        gameList = getGames()
        print(getTimes(gameList))
            
    if __name__ == "__main__":
        main()


    

































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
