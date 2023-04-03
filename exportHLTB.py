import sys
import openpyxl
import returnQuery #in working directory
from pathlib import Path
import glob
import os
            

try:
    fileName = returnQuery.findFile()
    wb = openpyxl.load_workbook(fileName)
    ws = wb["Sheet"]
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



    

def exportToSheet(timeList, typeList):
    y = "Y"#input("Is this your first time running through this script? Y/N\n")
    if y == "N" or y == "n":
        #possibly delete rightmost column? idk if i'm keeping the genres or not
        timeColumn = input("Select the Excel column you would like to put the times in. If you would like to enter in column C, please enter 3, etc. "
                  "Please keep in mind that the game list is in column A and the first set of times created was put in column B.\n")
        if timeColumn == "1":
            z = input("Are you sure? This will overwrite your games. Y/N\n")
            if z == "N" or z == "n":
                exportToSheet(timeList, typeList)
            elif z != "Y" or z != "y":
                print("Invalid input.")
                exportToSheet(timeList, typeList)
        elif timeColumn == "2":
            z = input("Are you sure? This will overwrite your times in column B. Y/N\n")
            if z == "N" or z == "n":
                exportToSheet(timeList, typeList)
            elif z != "Y" or z != "y":
                print("Invalid input.")
                exportToSheet(timeList, typeList)
                
        #if not first time, they already answered these questions, no need to ask again
        multi = "N"
        endless = "N"
    elif y == "Y" or y == "y":
        #defaults to column B
        timeColumn = "2"
        multi = input("Would you like to remove primarily multiplayer games? These are often given inflated completion times. Y/N\n")
        endless = input("Would you like to remove games marked as endless? These are often given arbitrary completion times. Y/N\n")
    else:
        print("Invalid input.")
        exportToSheet(timeList, typeList)
    #check if column input is integer
    try:
        timeColumn = int(timeColumn)
        if timeColumn < 1:
            print("Invalid column input.")
            exportToSheet(timeList) 
    except:
        print("Invalid column input.")
        exportToSheet(timeList)

    try:
        insert(typeList, timeList, timeColumn)
        delGenre(typeList, multi, endless)
        print("Exported to file successfully.")
        print("Refer to the README.txt in order to complete results.")
    except Exception as e:
        print("Export not successful. Refer to the README for instructions.")
        print(e)
        sys.exit()


def insert(typeList, times, col):
    ### Insert into Excel ###
    for i in range(0, len(times)):
        gameTime = ws.cell(row=i+1, column=col)
        #gameType = ws.cell(row=i+1, column=col+1)
        gameTime.value = times[i]
        #try:
        #    gameType.value = typeList[i]
        #except Exception as e:
        #    gameType.value = ""
    wb.save('SteamGames.xlsx')

def delGenre(typeList, multi, endless):
    if (multi == "Y" or multi == "y"):
        #delete rows bottom to top or it'll get screwed up
        for i in reversed(range(0,gameCount)):
            if typeList[i] == "multi":
                ws.delete_rows(i+1,1)
    if endless == "Y" or endless == "y":     
        for i in reversed(range(0,gameCount)):
            if typeList[i] == "endless":
                ws.delete_rows(i+1,1)
    wb.save('SteamGames.xlsx')


class HLTBSteam():
    def main():
        gameList = returnQuery.getGames()
        result = returnQuery.getTimes(gameList) #returns nested list
        timeList = result[0]
        typeList = result[1]
        exportToSheet(timeList, typeList)
        return 0

    if __name__ == "__main__":
        main()
