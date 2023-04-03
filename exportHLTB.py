import sys
import openpyxl
import returnQuery #in working directory
from pathlib import Path


try:

    print(Path.cwd())
    #C:\Users\Nathan\Documents\PySteamTimer
    
    #search for xlsx files in directory and rename
    wb = openpyxl.load_workbook("SteamGames.xlsx")
    ws = wb["Sheet"]
    #ws.delete_cols(2,3) #deletes 3 columns starting at B
    wsList = wb.sheetnames
    if len(wsList) > 1:
        print("These are the sheets in your excel file:")
        for i in range(0,len(wsList)):
            if i != (len(wsList)-1): #cuts off the comma on wsList[-1]
                print(wsList[i], end=", ")
            else:
                print(wsList[i])
        sheet = input("Choose the sheet you would like to query: ")
        ws = wb[sheet]
    gameCount = ws.max_row
    typeList = []
    searchResult = [] 
except Exception as e:
    print("A problem occurred:", e)
    print("If your error is not apparent, please check the README.txt file for assistance.")
    sys.exit()


def exportToSheet(games, times, typeList):
    y = input("Is this your first time running through this script? Y/N\n")
    if y == "N" or y == "n":
        x = input("Select the Excel column you would like to put the times in. If you would like to enter in column C, please enter 3, etc. "
                  "Please keep in mind that the game list is in column A and the first set of times created was put in column B.\n")
        if x == "1":
            z = input("Are you sure? This will overwrite your games. Y/N\n")
            if z == "N" or z == "n":
                exportToSheet(times)
                
        #if not first time, they already answered these questions, no need to ask again
        multi = "N"
        endless = "N"
    else:
        #defaults to column B
        x = "2"
        multi = input("Would you like to remove primarily multiplayer games? These are often given inflated completion times. Y/N\n")
        endless = input("Would you like to remove games marked as endless? These are often given arbitrary completion times. Y/N\n")

    #check if x input is integer
    try:
        x = int(x)
        if x < 1:
            print("Invalid column input.")
            exportToSheet(times) 
    except:
        print("Invalid column input.")
        exportToSheet(times)

    try:
        insert(times, x)
        delGenre(typeList, multi, endless)
        print("Exported to file successfully.")
        print("Refer to the README.txt in order to complete results.")
    except Exception as e:
        print(e)
        sys.exit()


def insert(times, x):
    ### Insert into Excel ###
    for i in range(0, len(times)):
        ws_cell = ws.cell(row=i+1, column=x)
        #try:
        ws_cell.value = times[i]
        #except Exception as e:
        #    print("test")
        #    print(e)
        #    typeList.append("")
    return typeList
            
    wb.save('SteamGames.xlsx')

def delGenre(typeList, multi, endless):
    if (multi == "Y" or multi == "y"):
        #find all rows, sort, reverse, then delete
        for i in range(1,gameCount):
            if typeList[i] == "multi":
                ws.delete_rows(i,1)
    if endless == "Y" or endless == "y":     
        for i in range(1,gameCount):
            if typeList[i] == "endless":
                ws.delete_rows(i,1)
    wb.save('SteamGames.xlsx')


class HLTBSteam():
    def main():
        gameList = returnQuery.getGames()
        result = returnQuery.getTimes(gameList)
        timeList = result[0]
        typeList = result[1]
        print("test 2")
        exportToSheet(gameList, timeList, typeList)
        return 0

    if __name__ == "__main__":
        main()
