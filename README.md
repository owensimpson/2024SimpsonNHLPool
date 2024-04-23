# 2024SimpsonNHLPool
Contains the Python Script for a personal project meant for an NHL Playoff Pool. The script scrapes the web for an Excel file, uses pandas to extract relevant information, then uses gspread to upload data to a Google Spreadsheet shared
by the participants of the pool. So, running the program will update the spreadsheet with the most up-to-date information relating to the points of the players chosen. Then, using Task Scheduler, I scheduled the program to run once a day at 2 am
so in the morning, the pool would be up to date with no manual data entry. 

The script is comprised of 3 main steps:

1. Using Selenium ChromeDriver, the script will open the NHL stats webpage on NHL.com. This page is dynamically loaded using JavaScript, meaning that I can't simply import the static data into Google Sheets. By first setting the download path,
  then Using developer tools, I can set the driver to download the Excel spreadsheet that is downloadable from the page using the CSS Selector.

2. Next, the spreadsheet will be analyzed and data will be extracted from it using pandas. Pandas extracts 30 player names (10 for each participant) along with their current points in the playoffs. This data is added to dictionaries for each player in the format
   {'Player name': Points}. Since the NHL stats page (and the Excel sheet by association) only contains the top 100 players on the first page, a while loop is set to continuously repeat the process of opening the next page on the website (using a different URL),
   download the spreadsheet for those 100 players, analyze and extract the data, and add it to dictionaries IF the length of the dictionaries are not ALL at 10 (this means all players have been acquired). Since every file we download has the same name, I
   needed to configure the logic for pandas to always open the most recently downloaded file. Once the dictionaries are complete, the next step begins.

4. Finally, using the dictionaries gathered for each participant, I set up a Google Cloud project to make use of the Google Sheets API. The scope and credentials of the API needed to be configured, as well as authorizing access to the spreadsheet.
  By defining the rows and columns where each player's data will be located, as well as coding the logic for Python to find the player's name, then placing the point values on the cell to the right, the script will input the player data into the correct positions.
  At this point, the script is finished.

Spreadsheet before running the program:
![image](https://github.com/owensimpson/2024SimpsonNHLPool/assets/167917725/09c14ea6-d73e-42d0-a338-e6be99d6821c)

Spreadsheet after running the program:
![image](https://github.com/owensimpson/2024SimpsonNHLPool/assets/167917725/1d15b897-b370-4fd2-bbf8-833a76eee4ad)
