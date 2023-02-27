from howlongtobeatpy import HowLongToBeat #https://pypi.org/project/howlongtobeatpy/
from openpyxl import load_workbook #https://ehmatthes.github.io/pcc_2e/beyond_pcc/extracting_from_excel/

#class howlongtobeat():
file = "C:/Users/Nathan/Documents/PySteamTimer/SteamGames.xlsx"
wb = load_workbook(file)
ws = wb["Sheet1"]

def accessSheet(): 
    rows = list(ws.rows)
    i = 0
    while i < 5:
        for j in rows[i]:
            print(j.value, end = " ")
        print("") #new line
        i+=1

    
def findGame(gameList):

    for i in gameList:
        results += HowLongToBeat().search(gameList[i])
        #gamestring = join([str(item) for item in results])

