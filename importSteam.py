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
    return IDList

def gameNameList(IDList):
    games = []
    for i in IDList:
        try:
            #print(i)
            game = steamClient.getApp(appid=i)
            name = game.name
            print(name)
            games.append(name)
        except:
            print("ID not found. App ID is", i)
            continue
    return games

#https://github.com/ValvePython/steam

class importSteam():

    def main():
        #test()
        user = steamClient.getUser(id64=steamID64)
        #print(user.name)
        games = user.apps #list of <steamfront.userapp.UserApp object at 0x00000XXXXXXXXX>
        print(games)
        idlist = appIDList(games) #steam IDs in mostly(?) ascending numerical order
        #print(idlist)
        gamelist = gameNameList(idlist)
        #it seems like this errors with too large of lists. split this up into multiple?
        print(gamelist)
        return 0
    
    if __name__ == "__main__":
        main()
