from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials


# This script opens up the NHL website playoff stats page, uses Selenium to automatically download the spreadsheet on the page,
# uses pandas to read through the spreadsheet and find the current points of each of our players, the uses gspread to input the
# data into a Google Sheet where we can view the current point values and standings.


# URL List (for when we need to switch pages, the page should not go past 5)
url = ['https://www.nhl.com/stats/skaters?reportType=season&seasonFrom=20232024&seasonTo=20232024&gameType=3&sort=points,goals,assists&page=0&pageSize=100',
       'https://www.nhl.com/stats/skaters?reportType=season&seasonFrom=20232024&seasonTo=20232024&gameType=3&sort=points,goals,assists&page=1&pageSize=100',
       'https://www.nhl.com/stats/skaters?reportType=season&seasonFrom=20232024&seasonTo=20232024&gameType=3&sort=points,goals,assists&page=2&pageSize=100']

# Players Chosen
O_players = ['Connor McDavid', 'Nikita Kucherov', 'Mikko Rantanen', 'Elias Pettersson', 'Brayden Point', 'Zach Hyman', 'Jason Robertson', 'Chris Kreider', 'Auston Matthews', 'Brad Marchand']
M_players = ['Nathan MacKinnon', 'Leon Draisaitl', 'David Pastrnak', 'Matthew Tkachuk', 'Roope Hintz', 'Mika Zibanejad', 'J.T. Miller', 'Steven Stamkos', 'Vincent Trocheck', 'Adrian Kempe']
G_players = ['Sebastian Aho', 'Jake Guentzel', 'Adam Fox', 'Cale Makar', 'Sam Reinhart', 'Aleksander Barkov', 'Evan Bouchard', 'Quinn Hughes', 'Mark Stone', 'Jack Eichel']

# Initialize Chrome Webdriver
service = Service(r"C:\Users\16139\OneDrive\Desktop\2024 NHL Playoff Pool\chromedriver-win64\chromedriver-win64\chromedriver.exe")
driver = webdriver.Chrome()
driver.maximize_window()

# Set Download Path to a Folder in the working directory
params = {'behavior': 'allow', 'downloadPath': r"C:\Users\16139\OneDrive\Desktop\2024 NHL Playoff Pool\AutomatedDownloads"}
driver.execute_cdp_cmd('Page.setDownloadBehavior', params)

i = 0
# Navigate to the page
driver.get(url[i])
time.sleep(3)  # Wait for the page to load

# Function to download statistics from NHL website
def download_stats():
    export_button = driver.find_element(By.CSS_SELECTOR, '#season-tabpanel > h4 > a > svg > path')
    export_button.click()
    time.sleep(10)  # Wait for the file to download
    driver.switch_to.window(driver.window_handles[0])

# Function to Extract Player Points
def extract_player_points(participant_choices, returnable_dict):

    # Find Newest Excel File
    files = os.listdir(r"C:\Users\16139\OneDrive\Desktop\2024 NHL Playoff Pool\AutomatedDownloads")
    newest_file = max(files, key=lambda x: os.path.getmtime(os.path.join(r"C:\Users\16139\OneDrive\Desktop\2024 NHL Playoff Pool\AutomatedDownloads", x)))
    file_path = os.path.join(r"C:\Users\16139\OneDrive\Desktop\2024 NHL Playoff Pool\AutomatedDownloads", newest_file)

    df = pd.read_excel(file_path)

    # Extract player names and points
    for index, row in df.iterrows():
        player_name = row['Player']
        points = row['P']

        # Check if the player is in the participant choices list
        if player_name in participant_choices:
            returnable_dict[player_name] = points

    return returnable_dict

# First Runthrough
download_stats()

O_Player_Points = {}
M_Player_Points = {}
G_Player_Points = {}

O_Player_Points = extract_player_points(O_players, O_Player_Points)
M_Player_Points = extract_player_points(M_players, M_Player_Points)
G_Player_Points = extract_player_points(G_players, G_Player_Points)

# set i to 1 (this will be increased every time it goes through to go to the next page
i = 1
# reiterate through the next pages of the website, download the new file and keep going
while True:

    driver.get(url[i])
    time.sleep(3)
    download_stats()

    O_Player_Points = extract_player_points(O_players, O_Player_Points)
    M_Player_Points = extract_player_points(M_players, M_Player_Points)
    G_Player_Points = extract_player_points(G_players, G_Player_Points)

    if len(O_Player_Points) == 10 and len(M_Player_Points) == 10 and len(G_Player_Points) == 10:
        break
    if i == 2:
        break
    i += 1

# Exit the Webdriver
driver.quit()

# From this point, the code automatically exports the data into the Google Sheets we are using for the pool
# Define the dictionaries
player_points = {'O': O_Player_Points, 'M': M_Player_Points, 'G': G_Player_Points}

# Define the scope and credentials
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(r"C:\Users\16139\OneDrive\Desktop\2024 NHL Playoff Pool\credentials.json.json", scope)

# Authorize access
gc = gspread.authorize(credentials)

# Open the Google Sheets spreadsheet
sheet = gc.open("Simpson NHL Playoff Pool 2024").sheet1

# Define the starting row for each participant
starting_rows = {'O': 7, 'M': 7, 'G': 7}

# Define the columns for each participant
columns = {'O': 'B', 'M': 'D', 'G': 'F'}

# Write data to specific cells
for participant, points_dict in player_points.items():
    row = starting_rows[participant]
    column = columns[participant]
    for player, points in points_dict.items():
        # Find the cell where the player name is located
        cell = sheet.find(player)
        # Update the cell next to the player name with the points
        sheet.update_cell(cell.row, cell.col + 1, points)

























