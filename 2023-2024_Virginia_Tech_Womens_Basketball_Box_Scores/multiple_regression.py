import pandas as pd
import statsmodels.api as sm

# Load your data from the Excel file
file_path = '2023-2024_Virginia_Tech_Womens_Basketball_Season_Summary_Cleaned.xlsx'
sheet_name = 'Game Summary Modified'
data = pd.read_excel(file_path, sheet_name=sheet_name)

# Define the independent variables in the specified order
X = data[['Total FGM', 'Total 3PTM', 'Total FTM', 'Total REB', 'Total AST', 'Total TO', 'Total BLK', 'Total STL', 
          'Points in the Paint', 'Points off Turnovers', 'Second Chance Points', 'Fast Break Points', 'Bench Points']]

# Define the dependent variable (Win or Loss)
y = data['Win or Loss (Numeric)']

# Add a constant to the independent variables
X = sm.add_constant(X)

# Perform the regression using statsmodels
model = sm.OLS(y, X).fit()

# Extract the summary table from the regression results
summary = model.summary2().tables[1]

# Convert the summary table to a DataFrame
summary_df = pd.DataFrame(summary)

# Save the DataFrame to a CSV file
summary_df.to_csv('regression_results.csv')

print("Regression results have been saved to 'regression_results.csv'")