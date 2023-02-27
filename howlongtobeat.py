from howlongtobeatpy import HowLongToBeat

class howlongtobeat():
    def findGame(gameList):

        for i in gameList:
            results = HowLongToBeat().search(gameList[i])
            #gamestring = join([str(item) for item in results])
