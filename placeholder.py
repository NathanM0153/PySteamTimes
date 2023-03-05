
#file = "C:/Users/Nathan/Documents/PySteamTimer/SteamGames.xlsx"
#wb = load_workbook(file)
#ws = wb["run_results"] #raw input in main?
#rows = list(ws.rows)
#gameCount = ws.max_row


#returns list of steam games from excel file
def getGameList(): 
    gameList = []
    for i in range(0,gameCount):
        for j in rows[i]:
            gameList.append(j.value)
    return gameList

main
        #gameList = getGames()
        #gameTimes = getTimes(gameList)
        #for i in range(0,gameCount):
        #    print(gameList[i], end=": ")
        #    print(gameTimes[i], "hours")
        #exportToSheet(gameTimes)
