Libraries needed:
- aiohttp
- requests
- fake_useragent
- openpyxl
- xlwt
- steamfront
- howlongtobeatpy

Install these in Windows Powershell using the command:
pip install [name of library]


HOW TO RUN:

OPTION 1:
- Open Windows Powershell
- Navigate to your installation folder using this command
cd [file path]

You can get your file path by clicking on the bar to the left of the search bar in File Explorer.
For example:
cd C:\Users\Admin\Documents\PySteamTimer

Once in the folder, run the following command:
python importSteam.py

This program will take several minutes to complete.

After completing the steps specified by the program, run this command:
python exportTimes.py



OPTION 2:
- Open IDLE, included with your installation of Python
- Open importSteam.py under File
- Run the script using Run -> Run Module

This program will take several minutes to complete.

After completing the steps specified by the program, open exportTimes.py and run it.





TROUBLESHOOTING

The Steam app ID was not found.

- Honestly 99% of the time this happens (to roughly 1% of my library) it's still unexplained.
If it's just a few, go to the links specified and find the games. If it's many more, contact me
and I'll figure something out. This program is probably in continual WIP anyway.


The game was not found on HowLongToBeat.

- Most likely, the game name isn't exactly what it should be. If the program searches for "[Game] Definitive Edition" 
it won't find "[Game]". This is an unfortunate quirk of HLTB's own search API. You will need to go into Excel and edit
the name of the game if the cell in column B specifies it was not found in HLTB's database. If the hours are set to 0,
HLTB correctly found the game, but simply does not have data on it.


Other reasons these may occur:

- The name is in another language
- The program is searching for non-game software such as Borderless Gaming
- Your internet is having problems
- You may be having rate limits

---

My Excel file was corrupted.
- Open Excel, go to File -> Open -> Browse. Single-click on SteamGames.xlsx, and navigate to the Open button on the
bottom right. Click the down arrow next to Open, click Open and Repair, Extract Data, Convert to Values. Finally,
go back to File -> Save As, and save the file as SteamGames.xlsx. You may have to change the drop down below the
file name from .xls to .xlsx. After that, you can rerun the export script.