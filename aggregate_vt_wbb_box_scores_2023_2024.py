from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, WebDriverException
import time
from io import StringIO
import os

# Path to the Edge WebDriver executable
webdriver_path = 'C:/Users/Ameen/Downloads/edgedriver_win64/msedgedriver.exe'

# Initialize the WebDriver with EdgeService
service = EdgeService(executable_path=webdriver_path)
driver = webdriver.Edge(service=service)

# URL of the 2023-2024 Virginia Tech Women's Basketball season page
# season_url = 'https://hokiesports.com/sports/womens-basketball/schedule/2023-24'
# URL of the 2022-2023 Virginia Tech Women's Basketball season page
season_url = 'https://hokiesports.com/sports/womens-basketball/schedule/2022-23'
driver.get(season_url)

# Function to get box score links
def get_box_score_links():
    return driver.find_elements(By.XPATH, "//a[span[text()='Box Score']]")

# Wait for the box score links to be loaded
WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.XPATH, "//a[span[text()='Box Score']]")))

# Get all box score links
box_score_links = get_box_score_links()

# Function to extract team names and date
def extract_team_names_and_date(soup):
    date_meta = soup.find('meta', {'name': 'description'})
    date_text = date_meta['content'] if date_meta else 'Unknown Date'
    game_date = date_text.split('on ')[-1] if 'on ' in date_text else 'Unknown Date'
    team_headers = soup.find_all('h3')
    if len(team_headers) >= 2:
        team1_name = team_headers[1].text.strip()
        team2_name = team_headers[2].text.strip()
        team1_parts = team1_name.split(" ")
        team1_parts.pop()
        team1_name = " ".join(team1_parts)
        team2_parts = team2_name.split(" ")
        team2_parts.pop()
        team2_name = " ".join(team2_parts)
    else:
        team1_name = 'Team 1'
        team2_name = 'Team 2'
    return team1_name, team2_name, game_date

# Create the folder if it doesn't exist
folder = "2022-2023_Virginia_Tech_Womens_Basketball_Box_Scores"
if not os.path.exists(folder):
    os.makedirs(folder)

# Initialize lists to accumulate season data
season_game_summaries = []
season_player_stats = []
season_team_summaries = []
season_technical_fouls = []

# Loop through each box score link
for i in range(len(box_score_links)):
    try:
        # Re-fetch the box score links to avoid stale element exception
        box_score_links = get_box_score_links()
        link = box_score_links[i]
        
        # Get the box score URL
        box_score_url = link.get_attribute('href')
        
        # Open the box score page
        driver.get(box_score_url)
        
        # Wait for the tables to be loaded
        WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.TAG_NAME, 'table')))
        
        # Get the page source
        page_source = driver.page_source
        
        # Parse with BeautifulSoup
        soup = BeautifulSoup(page_source, 'html.parser')
        
        # Extract team names and date
        team1_name, team2_name, game_date = extract_team_names_and_date(soup)
        
        # Find all tables on the box score page
        tables = soup.find_all('table')
        
        # Extract data from tables
        if len(tables) >= 1:
            game_summary_df = pd.read_html(StringIO(str(tables[0])))[0]
            game_summary_df['Game Date'] = game_date
            if team1_name == "Virginia Tech" or team2_name == "Virginia Tech":
                season_game_summaries.append(game_summary_df[game_summary_df.columns[:7]])
        
        if len(tables) >= 2:
            player_stats_df = pd.read_html(StringIO(str(tables[1])))[0]
            player_stats_df['Game Date'] = game_date
            player_stats_df['Team'] = team1_name
            if team1_name == "Virginia Tech":
                season_player_stats.append(player_stats_df)
        
        if len(tables) >= 3:
            team_summary_df = pd.read_html(StringIO(str(tables[2])))[0]
            team_summary_df['Game Date'] = game_date
            team_summary_df['Team'] = team1_name
            if team1_name == "Virginia Tech":
                season_team_summaries.append(team_summary_df)
        
        if len(tables) >= 4:
            technical_fouls_df = pd.read_html(StringIO(str(tables[3])))[0]
            technical_fouls_df['Game Date'] = game_date
            technical_fouls_df['Team'] = team1_name
            if team1_name == "Virginia Tech":
                season_technical_fouls.append(technical_fouls_df)
        
        if len(tables) >= 5:
            player_stats_df2 = pd.read_html(StringIO(str(tables[4])))[0]
            player_stats_df2['Game Date'] = game_date
            player_stats_df2['Team'] = team2_name
            if team2_name == "Virginia Tech":
                season_player_stats.append(player_stats_df2)
        
        if len(tables) >= 6:
            team_summary_df2 = pd.read_html(StringIO(str(tables[5])))[0]
            team_summary_df2['Game Date'] = game_date
            team_summary_df2['Team'] = team2_name
            if team2_name == "Virginia Tech":
                season_team_summaries.append(team_summary_df2)
        
        if len(tables) >= 7:
            technical_fouls_df2 = pd.read_html(StringIO(str(tables[6])))[0]
            technical_fouls_df2['Game Date'] = game_date
            technical_fouls_df2['Team'] = team2_name
            if team2_name == "Virginia Tech":
                season_technical_fouls.append(technical_fouls_df2)
        
        # Go back to the season page
        driver.get(season_url)
        
        # Wait for the box score links to be loaded again
        WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.XPATH, "//a[span[text()='Box Score']]")))
    
    except (StaleElementReferenceException, TimeoutException, WebDriverException, ConnectionResetError) as e:
        print(f"Encountered exception: {e}. Retrying...")
        driver.get(season_url)
        WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.XPATH, "//a[span[text()='Box Score']]")))
        box_score_links = get_box_score_links()
        i -= 1  # Decrement the index to retry the current link
    
    # Add a small delay to avoid rapid navigation which might lead to errors
    time.sleep(1)

# Combine season data into DataFrames
season_game_summaries_df = pd.concat(season_game_summaries, ignore_index=True) if season_game_summaries else pd.DataFrame()
season_player_stats_df = pd.concat(season_player_stats, ignore_index=True) if season_player_stats else pd.DataFrame()
season_team_summaries_df = pd.concat(season_team_summaries, ignore_index=True) if season_team_summaries else pd.DataFrame()
season_technical_fouls_df = pd.concat(season_technical_fouls, ignore_index=True) if season_technical_fouls else pd.DataFrame()

# Export season data to Excel
season_filename = f'{folder}/2022-2023_Virginia_Tech_Womens_Basketball_Season_Summary.xlsx'
with pd.ExcelWriter(season_filename, engine='xlsxwriter') as writer:
    season_game_summaries_df.to_excel(writer, sheet_name='Game Summary', index=False)
    season_player_stats_df.to_excel(writer, sheet_name='Player Statistics', index=False)
    season_team_summaries_df.to_excel(writer, sheet_name='Team Summary', index=False)
    season_technical_fouls_df.to_excel(writer, sheet_name='Technical Fouls', index=False)

print(f"Season data successfully exported to {season_filename}")

# Close the browser
driver.quit()

print("All box scores successfully exported.")
