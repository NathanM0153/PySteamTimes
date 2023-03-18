import time
import steamfront
#credit: https://github.com/4Kaylum/Steamfront
import xlwt
from xlwt import Workbook
from openpyxl import load_workbook
from openpyxl import Workbook


#steamID64 = "76561198272854176"
#steamAPIKey = "6D5F599D289CCFC212F5D824F212CCB4"
steamID64 = input("Enter your Steam ID64 key. This is the number at the end of your profile URL.\n")
steamAPIKey = input("Enter your Steam API Key. After registering, you can find yours at this link: \nhttps://steamcommunity.com/dev/apikey\n")
steamClient = steamfront.Client(steamAPIKey)
waitTime = 1.25 #rate limit on steamfront api requests


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
        time.sleep(waitTime)
        #would it be faster to go full speed until error then wait for longer?
        count += 1
    print("Game IDs not found:")
    print(errors)
    print("You can find them at https://steamcommunity.com/app/######, inserting the relevant ID.")
    print("Most of these will be game accessories or discontinued games, but there is a possibility the script missed a game so check if you would like.")
    return games

def exportToExcel(gameList):
    file = "SteamGames.xlsx"
    wb = Workbook()
    #wb.save(file)
    gameList.sort()
    gameSheet = wb.create_sheet("Games")
    wb.remove("Sheet")
    for i in range(0,len(gameList)):
        gameSheet.cell(row=i+1, column=1, value=gameList[i])
    wb.save(file)
    print("File successfully saved.")
    wb.close()
    

def doctorOutput(gameList):
    fixedList = []
    bracketbool = False
    string = ""
    for i in gameList:
        i = i.replace("™","")
        i = i.replace("®","")
        i = i.replace(";"," ")
        i = i.replace("- ", "")
        i = i.replace("– ", "")
        
        string = ""
        for j in i: #removes everything in parentheses
            if j == '(':
                bracketbool = True
                #turns off adding to string
            elif j == ')':
                bracketbool = False
                #turns on adding to string
            elif not bracketbool:
                string += j
        i = string
        
        if i[-1] == " ":
            i = i[:-1] #cuts off if there is one
            
        fixedList.append(i)
    return fixedList


class importSteam():

    def main():
        user = steamClient.getUser(id64=steamID64)
        games = user.apps #list of <steamfront.userapp.UserApp object at 0x00000XXXXXXXXX>
        IDList = appIDList(games)
        time = round(len(IDList) * waitTime / 60 + (len(IDList)/60/2.5), 1) #roughly 2.5 queries/second at max speed
        print("This process is expected to take roughly", time, "minutes.")
        gameList = gameNameList(IDList)
        doctoredList = doctorOutput(gameList)
        exportToExcel(doctoredList)
        return 0
    
    if __name__ == "__main__":
        main()
