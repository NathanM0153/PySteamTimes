Welcome to PySteamTimes! This program is meant to take all your games from Steam,
put them in Excel, and then search HowLongToBeat.com for all of your games and put in 
the times. I've written this README as if the user has zero experience with command
line instructions. If that's you, don't worry! This shouldn't be too involved.
If you have ideas for features, please contact me on Discord at TM#7221.


Libraries needed:
- aiohttp
- requests
- fake_useragent
- openpyxl
- steamfront
- howlongtobeatpy

Install these in Windows Powershell using the command:
pip install [name of library]

If for whatever reason that doesn't work, use this instead:
python -m pip install [name of library]


HOW TO RUN:

OPTION 1:
- Open Windows Powershell (or Linux terminal, if using Linux)
- Navigate to your installation folder using this command
cd [file path]

You can get your file path by clicking on the bar to the left of the search bar in File Explorer.
For example:
cd C:\Users\Admin\Documents\PySteamTimes

Once in the folder, run the following command:
python importSteam.py

This program will take several minutes to complete.

After completing the steps specified by the program, run this command:
python exportHLTB.py


OPTION 2:

Windows:
- Right-click on importSteam.py, click "Open with", then choose another app.
- Navigate to your installation of Python, and check "Always use this app to open .py files"
- By double clicking on any py file, it should automatically open and run the file on Command Prompt.

Linux: 
- Right-click on the .py file and select "Properties". 
- In the "Open With" tab, select Python from the list of applications.
- Click on "Set as default".


This program will take several minutes to complete.

After completing the steps specified by the program, double click on exportHLTB.py.


OPTION 3:
- Open IDLE, included with your installation of Python
- Open importSteam.py under File
- Run the script using Run -> Run Module

This program will take several minutes to complete.

After completing the steps specified by the program, open exportHLTB.py and run it.





TROUBLESHOOTING

The Steam app ID was not found.

- 99% of the time this happens it's due to a game either being no longer purchaseable or being an 
accessory to another game, like a Friend's Pass. If you want to be absolutely sure, go to the links 
specified and find the games. If it's many more than 1 or 2% of your library, contact me and I'll 
figure something out. This program is probably in continual WIP anyway.


Many Steam App IDs in a row are not found.

- You ran into a problem with Steam's request rate limit. Wait for 3-5 minutes and re-run the script.


The game was not found on HowLongToBeat.

- Most likely, the game name isn't exactly what it should be. If the program searches for "[Game] Definitive Edition"
it won't find "[Game]". This is an unfortunate quirk of HLTB's own search API. You will need to go into Excel and edit
the name of the game if the cell in the time column specifies it was not found in HLTB's database. It is also possible
that HLTB does not have any record of the game's existence whatsoever. If the hours are set to 0, HLTB correctly found
the game in its database, but simply does not have any completion data on it.

Other reasons this may occur:

- The name is in another language
- The program is searching for non-game software such as Borderless Gaming or a soundtrack
- Your internet is having problems

---

My Excel file was corrupted.
- Open Excel, go to File -> Open -> Browse. Single-click on SteamGames.xlsx, and navigate to the Open button on the
bottom right. Click the down arrow next to Open, click Open and Repair, Extract Data, Convert to Values. Finally,
go back to File -> Save As, and save the file. After that, you can rerun the export script.

My error isn't listed here.
- Please try running the script again. Sometimes there are weird errors that I can't reproduce and disappear on the
next time I try running it. If it persists after 2 or 3 tries, contact me.


If none of these answer your question, you are free to contact me on Discord at TM#7221 and I will help.