import pandas as pd

# Load the Excel file
file_path = '2023-2024_Virginia_Tech_Womens_Basketball_Season_Summary.xlsx'
data = pd.ExcelFile(file_path)

# Define the cleaned file path
cleaned_file_path = '2023-2024_Virginia_Tech_Womens_Basketball_Season_Summary_Cleaned.xlsx'

# Load each sheet into a DataFrame
game_summary_df = data.parse('Game Summary')
player_stats_df = data.parse('Player Statistics')
team_summary_df = data.parse('Team Summary')
technical_fouls_df = data.parse('Technical Fouls')

# Split team summary into raw stats and percentages
team_summary_raw = team_summary_df[team_summary_df['FG'].str.contains('-')].copy()
team_summary_percentage = team_summary_df[~team_summary_df['FG'].str.contains('-')].copy()

# Function to split columns with dashes into separate columns
def split_columns(df, columns):
    for col in columns:
        if col in df.columns:
            print(f"Splitting column: {col}")
            print(df[col].head())  # Debug: print first few rows of the column before splitting
            if df[col].str.contains('-').any():
                df[[f'{col}M', f'{col}A']] = df[col].str.split('-', expand=True).astype(int)
                print(f"Column {col} split into {col}M and {col}A")
                print(df[[f'{col}M', f'{col}A']].head())  # Debug: print first few rows after splitting
            df.drop(columns=col, inplace=True)
        else:
            print(f"Column {col} not found in DataFrame.")

# Apply split_columns function
split_columns(team_summary_raw, ['FG', '3PT', 'FT'])
split_columns(player_stats_df, ['FG', '3PT', 'FT'])

# Ensure all expected columns are present in the percentage DataFrame
for col in ['FGM', 'FGA', '3PTM', '3PTA', 'FTM', 'FTA']:
    if col not in team_summary_percentage.columns:
        team_summary_percentage[col] = None

# Convert percentage values to floats
team_summary_percentage[['FGM', 'FGA', '3PTM', '3PTA', 'FTM', 'FTA']] = team_summary_percentage[['FGM', 'FGA', '3PTM', '3PTA', 'FTM', 'FTA']].replace('', pd.NA).astype(float) / 100

# Split quarters into separate sheets
quarters = ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr']
raw_sheets = {}
percentage_sheets = {}

for quarter in quarters:
    raw_sheets[quarter] = team_summary_raw[team_summary_raw['Team Summary'] == quarter.replace('Qtr', 'Quarter')].drop(columns=['Team Summary'])
    percentage_sheets[quarter] = team_summary_percentage[team_summary_percentage['Team Summary'] == quarter.replace('Qtr', 'Quarter')].drop(columns=['Team Summary', 'FGM', 'FGA', '3PTM', '3PTA', 'FTM', 'FTA'])

# Calculate total percentage for each game
total_percentage = team_summary_percentage.groupby(['Game Date', 'Team']).sum().reset_index()
for col in ['FGM', 'FGA', '3PTM', '3PTA', 'FTM', 'FTA']:
    if col in total_percentage.columns:
        total_percentage[col] = total_percentage[col].replace('', pd.NA).astype(float)

# Clean player statistics by splitting FG, 3PT, and FT into made and attempted
split_columns(player_stats_df, ['FG', '3PT', 'FT'])
player_stats_df.drop(columns=['GS'], inplace=True)

# Split ORB-DRB into separate columns
if 'ORB-DRB' in player_stats_df.columns:
    player_stats_df[['ORB', 'DRB']] = player_stats_df['ORB-DRB'].str.split('-', expand=True).astype(int)
    player_stats_df.drop(columns=['ORB-DRB'], inplace=True)

# Ensure all expected columns are present and fill missing values with 0
for col in ['FGM', 'FGA', '3PTM', '3PTA', 'FTM', 'FTA']:
    if col not in player_stats_df.columns:
        player_stats_df[col] = 0

# Replace NaNs with 0 in player stats
player_stats_df.fillna(0, inplace=True)

# Rename '##' column to 'Num'
player_stats_df.rename(columns={'##': 'Num'}, inplace=True)

# Rearrange columns
player_stats_df = player_stats_df[['Num', 'Player', 'MIN', 'FGM', 'FGA', '3PTM', '3PTA', 'FTM', 'FTA', 'ORB', 'DRB', 'REB', 'PF', 'A', 'TO', 'BLK', 'STL', 'PTS', 'Game Date', 'Team']]

# Rename for easier typing
tfdf = technical_fouls_df

# Define the columns for the clean DataFrame
columns = ["Game Date", "Team", "Technical Fouls", "Points in the Paint", "Points off Turnovers", "Second Chance Points", "Fast Break Points", "Bench Points", "Scores Tied", "Lead Changed"]
clean_df = pd.DataFrame(columns=columns)

for index in range(0, len(tfdf), 3):
    # Extract game date and team name (common for all three rows)
    game_date = tfdf.loc[index, "Game Date"]
    team = tfdf.loc[index, "Team"]
    
    # Initialize a dictionary to hold the data for this game
    game_data = {"Game Date": game_date, "Team": team}

    # Extract and parse data from the three rows
    technical_fouls = tfdf.loc[index, 0].split(":")[1].strip()
    if technical_fouls.lower() == "none":
        technical_fouls = 0
    else:
        try:
            technical_fouls = int(technical_fouls)
        except ValueError:
            technical_fouls = 0  # Handle non-integer values gracefully

    points_in_the_paint = tfdf.loc[index + 1, 0].split(":")[1].strip()
    points_off_turnovers = tfdf.loc[index + 2, 0].split(":")[1].strip()
    second_chance_points = tfdf.loc[index, 1].split(":")[1].strip()
    fast_break_points = tfdf.loc[index + 1, 1].split(":")[1].strip()
    bench_points = tfdf.loc[index + 2, 1].split(":")[1].strip()
    scores_tied = tfdf.loc[index, 2].split(":")[1].replace(" time(s)", "").strip()
    lead_changed = tfdf.loc[index + 1, 2].split(":")[1].replace(" time(s)", "").strip()

    # Add extracted data to the dictionary
    game_data.update({
        "Technical Fouls": technical_fouls,
        "Points in the Paint": points_in_the_paint,
        "Points off Turnovers": points_off_turnovers,
        "Second Chance Points": second_chance_points,
        "Fast Break Points": fast_break_points,
        "Bench Points": bench_points,
        "Scores Tied": scores_tied,
        "Lead Changed": lead_changed
    })

    # Convert game_data dictionary to DataFrame for this iteration
    game_data_df = pd.DataFrame([game_data])

    # Use pd.concat() to append the game_data_df to clean_df
    clean_df = pd.concat([clean_df, game_data_df], ignore_index=True)

# Export the cleaned data to a new Excel file
with pd.ExcelWriter(cleaned_file_path, engine='xlsxwriter') as writer:
    game_summary_df.to_excel(writer, sheet_name='Game Summary', index=False)
    player_stats_df.to_excel(writer, sheet_name='Player Statistics', index=False)
    
    for quarter in quarters:
        raw_sheets[quarter].to_excel(writer, sheet_name=f'{quarter} Raw', index=False)
        percentage_sheets[quarter].to_excel(writer, sheet_name=f'{quarter} Pct', index=False)
    
    clean_df.to_excel(writer, sheet_name='Team Stats', index=False)

print(f"Cleaned data successfully exported to {cleaned_file_path}")
