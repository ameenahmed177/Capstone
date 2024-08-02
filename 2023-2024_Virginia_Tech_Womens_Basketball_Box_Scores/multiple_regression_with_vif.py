import pandas as pd
import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor

# Load your data from the Excel file
file_path = '2023-2024_Virginia_Tech_Womens_Basketball_Season_Summary_Cleaned.xlsx'
sheet_name = 'Game Summary Modified'
data = pd.read_excel(file_path, sheet_name=sheet_name)

# Define the independent variables in the specified order
X = data[['Total FGM', 'Total 3PTM', 'Total FTM', 'Total REB', 'Total AST', 'Total TO', 'Total BLK', 'Total STL', 
          'Points in the Paint', 'Points off Turnovers', 'Second Chance Points', 'Fast Break Points', 'Bench Points']]

# Add a constant to the independent variables
X = sm.add_constant(X)

# Calculate VIF for each predictor
vif_data = pd.DataFrame()
vif_data["feature"] = X.columns
vif_data["VIF"] = [variance_inflation_factor(X.values, i) for i in range(len(X.columns))]

print(vif_data)