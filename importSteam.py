import time
import steamfront
import sys
#credit: https://github.com/4Kaylum/Steamfront
from openpyxl import Workbook



steamID64 = input("Enter your Steam ID64 key. This is the number at the end of your profile URL.\n")
steamAPIKey = input("Enter your Steam API Key. After registering, you can find yours at this link: \nhttps://steamcommunity.com/dev/apikey\n")
freebool = input("Include free games? Y/N\n")
steamClient = steamfront.Client(steamAPIKey)
waitTime = 1.3 #rate limit on steam api requests


def appIDList(games):
    IDList = []
    try:
        for i in games:
            IDList.append(i.appid)
    except:
        print("Internet connection required.")
        sys.exit()
    return IDList
        

def gameNameList(IDList):
    games = []
    errors = []
    count = 1
    
    for i in IDList:
        try:
            game = steamClient.getApp(appid=i)
            name = game.name
            if freebool == "Y" or freebool == "y":
                print(count, "  ", name)
                count += 1
                games.append(name)
            else:
                if not game.is_free: #exclude free games
                    print(count, "  ", name)
                    count += 1
                    games.append(name)
        except:
            print("ID not found. App ID is", i)
            errors.append(i)
            continue
        time.sleep(waitTime)
    if len(errors) > 0:
        print("Game IDs not found:")
        for i in errors:
            i = int(i)
            print(i, end="\n")

        print("You can find them at https://steamcommunity.com/app/######, inserting the relevant ID.")
        print("Most of these will be game accessories or discontinued games, but there is a possibility the script missed a game so check if you would like.")
    else:
        print("All games found successfully.")
    return games

def exportToExcel(gameList):
    file = "SteamGames.xlsx"
    wb = Workbook()
    wb.save(file)
    gameList.sort()
    gameSheet = wb.active
    for i in range(0,len(gameList)):
        gameSheet.cell(row=i+1, column=1, value=gameList[i])
    wb.save(file)
    print("File successfully saved.")
    wb.close()
    

def doctorOutput(gameList):
    fixedList = []
    string = ""
    bracketbool = False
    for i in gameList:
        i = i.replace("™","")
        i = i.replace("®","")
        i = i.replace("–", "-") #en dash
        i = i.replace("—", "-") #em dash
        i = i.replace(" - ", " ")
        i = i.replace("’", "'")

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
            i = i[:-1] #cuts off extra space if there is one
            
        fixedList.append(i)
    return fixedList


class importSteam():

    def main():
        user = steamClient.getUser(id64=steamID64)
        games = user.apps #list of <steamfront.userapp.UserApp object at 0x00000XXXXXXXXX>
        IDList = appIDList(games)
        gameList = gameNameList(IDList)
        doctoredList = doctorOutput(gameList)
        exportToExcel(doctoredList)
        return 0
    
    if __name__ == "__main__":
        main()
