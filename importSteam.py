import time
import steamfront
#https://steamfront.readthedocs.io/en/latest/code-reference.html
import xlwt
import xlrd
from xlwt import Workbook
#https://www.geeksforgeeks.org/writing-excel-sheet-using-python/
from openpyxl import load_workbook
from openpyxl import Workbook
#https://ehmatthes.github.io/pcc_2e/beyond_pcc/extracting_from_excel/

filePath = "C:/Users/Nathan/Documents/PySteamTimer/"
steamID64 = "76561198272854176"
steamAPIKey = "6D5F599D289CCFC212F5D824F212CCB4"
steamClient = steamfront.Client(steamAPIKey)
waitTime = 1.5


def appIDList(games):
    IDList = []
    for i in games:
        IDList.append(i.appid)
        #print(games[i].appid)
    return IDList

def gameNameList(IDList):
    games = []
    errors = []
    count = 1
    #check if game is free? exclude?
    for i in IDList:
        try:
            game = steamClient.getApp(appid=i)
            name = game.name
            print(count, "  ", name)
            games.append(name)
        except:
            print("ID not found. App ID is", i)
            errors.append(i)
            continue
        time.sleep(waitTime) #rate limit on api requests
        #would it be faster to go full speed until error then wait for longer?
        count += 1
    print("Game IDs not found:")
    print(errors)
    print("You can find them at https://steamcommunity.com/app/######, inserting the relevant ID.")
    print("For complete accuracy, please add them to the bottom of the spreadsheet SteamGames.xlsx.")
    return games

def doctorOutput(gameList):
    fixedList = []
    bracketbool = False
    
    for i in gameList:
        i = i.replace("™","")
        i = i.replace("®","")
        i = i.replace(";"," ")
        #i = i.replace(":","")
        i = i.replace("- ", "")
        i = i.replace("– ", "")
        
        #string = ""
        #for j in i:
        #    if j == '(':
        #        bracketbool = True
        #        #turns off adding to string
        #    elif j == ')':
        #        bracketbool = False
        #        #turns on adding to string
        #    elif not bracketbool:
        #        string += j
        
        if i[-1] == ":":
            i = i[:-1] #cuts off if there is one
        
        #find a way to remove anything in parentheses
        fixedList.append(i)
    return fixedList

def exportToExcel(gameList):
    file = "C:\\Users\\Nathan\\Documents\\PySteamTimer\\SteamGames.xlsx"
    wb = Workbook()
    wb.save(file)
    gameSheet = wb.active #grab the active worksheet
    #gameSheet = wb.add_sheet("run_results")
    for i in range(0,len(gameList)):
        gameSheet.cell(row=i+1, column=1, value=gameList[i])
    #sort??
    wb.save(file)
    print("File successfully saved.")
    wb.close()
    


class importSteam():

    def main():
        user = steamClient.getUser(id64=steamID64)
        games = user.apps #list of <steamfront.userapp.UserApp object at 0x00000XXXXXXXXX>
        IDList = appIDList(games)
        time = round(len(IDList) * waitTime / 60 + (len(IDList)/60/2.5), 1)
        print("This process is expected to take roughly", time, "minutes.")
        gameList = gameNameList(IDList)
        doctoredList = doctorOutput(gameList)
        exportToExcel(doctoredList)
        return 0
    
    if __name__ == "__main__":
        main()
