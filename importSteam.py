import steamfront
#https://steamfront.readthedocs.io/en/latest/code-reference.html
import xlwt
from xlwt import Workbook
#https://www.geeksforgeeks.org/writing-excel-sheet-using-python/
from openpyxl import load_workbook
#https://ehmatthes.github.io/pcc_2e/beyond_pcc/extracting_from_excel/



#file = "C:/Users/Nathan/Documents/PySteamTimer/SteamGames.xlsx"
#wb = load_workbook(file)
#ws = wb["run_results"] #raw input in main?
#rows = list(ws.rows)

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


class importSteam():

    def main():
        user = steamClient.getUser(id64=steamID64)
        games = user.apps
        idlist = appIDList(games)
        gamelist = gameNameList(idlist)
        print(gamelist)
        return 0
    
    if __name__ == "__main__":
        main()
