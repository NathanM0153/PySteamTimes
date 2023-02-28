#modules needed: aiohttp, requests, fake_useragent, openpyxl
#scrape games via ParseHub

from howlongtobeatpy import HowLongToBeat #https://pypi.org/project/howlongtobeatpy/
from openpyxl import load_workbook #https://ehmatthes.github.io/pcc_2e/beyond_pcc/extracting_from_excel/

file = "C:/Users/Nathan/Documents/PySteamTimer/SteamGames.xlsx"
wb = load_workbook(file)
ws = wb["run_results"] #raw input in main?
rows = list(ws.rows)

#make this dynamic later
gameCount = 717

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
        

#takes in a string, returns a JSON result to be interpreted
def searchforGame(name): 
    results = HowLongToBeat().search(name, similarity_case_sensitive=False)
    if results is not None and len(results) > 0:
        best_element = max(results, key=lambda element: element.similarity)
        print(name, "found")
        return best_element
    else:
        print(name, "NOT FOUND")

#takes in JSON, returns name of game as string
def getName(game):
    return game.game_name

    


class HLTBSteam():

    def main():
        gameList = getGames()
        gameTimes = getTimes(gameList)
        print(gameTimes)
            
    if __name__ == "__main__":
        main()


    


