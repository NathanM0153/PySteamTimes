import time
import re
import steamfront
#https://steamfront.readthedocs.io/en/latest/code-reference.html
import xlwt
from xlwt import Workbook
#https://www.geeksforgeeks.org/writing-excel-sheet-using-python/
from openpyxl import load_workbook
#https://ehmatthes.github.io/pcc_2e/beyond_pcc/extracting_from_excel/

filePath = "C:/Users/Nathan/Documents/PySteamTimer/"
steamID64 = "76561198272854176"
steamAPIKey = "6D5F599D289CCFC212F5D824F212CCB4"
steamClient = steamfront.Client(steamAPIKey)
waitTime = 1.25


def appIDList(games):
    IDList = []
    for i in games:
        IDList.append(i.appid)
        #print(games[i].appid)
    return IDList

def gameNameList(IDList):
    games = []
    errors = []
    for i in IDList:
        try:
            game = steamClient.getApp(appid=i)
            name = game.name
            print(name)
            games.append(name)
        except:
            print("ID not found. App ID is", i)
            errors.append(i)
            continue
        time.sleep(waitTime) #rate limit on api requests
        #would it be faster to go full speed until error then wait for longer?
    print("Game IDs not found:")
    print(errors)
    return games

def doctorOutput(gameList):
    fixedList = []
    for i in gameList:
        j = i.replace("™","")
        k = j.replace("®","")
        m = k.replace(";"," ")
        n = m.replace("Definitive Edition", "")
        o = n.replace("The Final Cut", "")
        p = o.replace("Deluxe Edition", "")
        r = p.replace("Enhanced Edition", "")
        s = r.replace("Enhanced Plus Edition", "")
        re.sub("[\(\[].*?[\)\]]", "", s) #what does this do?
        fixedList.append(s)
    #™, ®, ; to be removed
    return fixedList

def exportToExcel(gameList):
    wb = Workbook()
    gameSheet = wb.add_sheet("run_results")
    for i in range(0,len(gameList)):
        gameSheet.write(i, 0, gameList[i])
    file = filePath + "SteamGames.xlsx"
    wb.save(file)
    print("File successfully saved.")


class importSteam():

    def main():
        user = steamClient.getUser(id64=steamID64)
        games = user.apps #list of <steamfront.userapp.UserApp object at 0x00000XXXXXXXXX>
        IDList = appIDList(games)
        print("This process is expected to take roughly", round(len(IDList) * waitTime / 60, 1), "minutes.")
        gameList = gameNameList(IDList)
        doctoredList = doctorOutput(gameList)
        exportToExcel(doctoredList)
        return 0
    
    if __name__ == "__main__":
        main()
