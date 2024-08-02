import pandas as pd

# Load the Excel file
file_path = '2023-2024_Virginia_Tech_Womens_Basketball_Season_Summary_Cleaned.xlsx'
data = pd.ExcelFile(file_path)

# Load the Player Statistics sheet
player_stats_df = data.parse('Player Statistics')

# Find the rows where the Totals are located
totals_rows = player_stats_df[player_stats_df['Player'] == 'Totals']

# Extract the Total Rebounds
total_rebounds = totals_rows[['REB']].copy()
total_rebounds.columns = ['Total REB']

# Extract the Total Assists
total_assists = totals_rows[['A']].copy()
total_assists.columns = ['Total AST']

# Extract the Total Turnovers
total_turnovers = totals_rows[['TO']].copy()
total_turnovers.columns = ['Total TO']

# Extract the Total Blocks
total_blocks = totals_rows[['BLK']].copy()
total_blocks.columns = ['Total BLK']

# Extract the Total Steals
total_steals = totals_rows[['STL']].copy()
total_steals.columns = ['Total STL']

# Combine the extracted totals into a single DataFrame
totals_df = pd.concat([total_rebounds, total_assists, total_turnovers, total_blocks, total_steals], axis=1)

# Save the combined DataFrame to a new Excel file
output_file_path = '2023-2024_Virginia_Tech_Womens_Basketball_Season_Totals.xlsx'

totals_df.to_excel(output_file_path, index=False)

print(f"Total rebounds successfully exported to {output_file_path}")
