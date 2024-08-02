import pandas as pd
import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor

# Load your data from the Excel file
file_path = '2023-2024_Virginia_Tech_Womens_Basketball_Season_Summary_Cleaned.xlsx'
sheet_name = 'Game Summary Modified'
data = pd.read_excel(file_path, sheet_name=sheet_name)

# Define the independent variables including 'Total TO'
X = data[['Total 3PTM', 'Total FTM', 'Total REB', 'Total AST', 'Total TO']]

# Define the dependent variable (Win or Loss)
y = data['Win or Loss (Numeric)']

# Add a constant to the independent variables
X = sm.add_constant(X)

# Calculate VIF for each predictor
vif_data = pd.DataFrame()
vif_data["feature"] = X.columns
vif_data["VIF"] = [variance_inflation_factor(X.values, i) for i in range(len(X.columns))]

# Print VIF data
print(vif_data)

# Fit the regression model
model = sm.OLS(y, X).fit()

# Extract the summary table as a DataFrame
summary_df = model.summary2().tables[1]

# Export the summary table to a CSV file
csv_file_name = 'regression_results_with_total_to.csv'
summary_df.to_csv(csv_file_name)

# Print confirmation message
print(f"CSV file '{csv_file_name}' has been exported successfully.")