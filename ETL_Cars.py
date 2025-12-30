import pandas as pd

# Step 1: Extract - Load the original Excel file
input_file = 'Car_sales_modified - Copy.xlsx'  # Replace with your actual file path if needed
df = pd.read_excel(input_file, sheet_name='Sheet1')

print(f"Extraction complete: {df.shape[0]} rows, {df.shape[1]} columns loaded.")
print("Original columns:", df.columns.tolist())
print("\nUnique Product Categories:", df['Product_Category'].unique())
print("Unique Sub Categories:", df['Sub_Category'].unique())

# Step 2: Transform

# 2.1: Data Cleaning
# Convert Date column to proper datetime (assuming the first date-related column is 'Date')
date_cols = ['Date', 'Day', 'Month', 'Year']  # Common in this dataset
if 'Date' in df.columns:
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

# Remove duplicates if any
df.drop_duplicates(inplace=True)

# Handle missing values (fill or drop as appropriate; here filling numerics with 0, categoricals with 'Unknown')
df.fillna({
    'Customer_Age': df['Customer_Age'].median(),
    'Order_Quantity': 0,
    'Profit': 0,
    'Revenue': 0,
    # Categorical
    'Age_Group': 'Unknown',
    'Customer_Gender': 'Unknown',
    'Country': 'Unknown',
    'State': 'Unknown',
}, inplace=True)

# 2.2: Feature Engineering
# Calculate Profit Margin if not already useful
if 'Profit' in df.columns and 'Revenue' in df.columns:
    df['Profit_Margin_%'] = (df['Profit'] / df['Revenue'] * 100).round(2)

# Create a full date string or extract year-month for aggregation later
df['Year_Month'] = df['Date'].dt.strftime('%Y-%m')

# 2.3: Rename columns for consistency (optional, make them cleaner)
df.rename(columns={
    'Customer_Age': 'Customer_Age',
    'Age_Group': 'Age_Group',
    'Customer_Gender': 'Gender',
    'Product_Category': 'Category',
    'Sub_Category': 'Sub_Category',
    'Order_Quantity': 'Quantity',
    'Unit_Cost': 'Unit_Cost',
    'Unit_Price': 'Unit_Price',
    'Profit': 'Profit',
    'Cost': 'Total_Cost',
    'Revenue': 'Total_Revenue'
}, inplace=True)

# 2.4: Filter or enrich as needed (example: focus on key countries or categories)
# df = df[df['Country'] == 'United States']  # Uncomment if you want to filter

print("\nTransformation complete:")
print("New columns added: Profit_Margin_%, Year_Month")
print("Duplicates removed and missing values handled.")
print(f"Rows after cleaning: {df.shape[0]}")

# Step 3: Load - Save the transformed data to a new Excel file (and optionally CSV)
output_excel = 'Cars_sales_ETL_transformed.xlsx'
output_csv = 'Cars_sales_ETL_transformed.csv'

df.to_excel(output_excel, sheet_name='Transformed_Data', index=False)
df.to_csv(output_csv, index=False)

print(f"\nLoad complete: Transformed data saved to '{output_excel}' and '{output_csv}'")
print("\nSample of transformed data:")
print(df.head().to_string(index=False))
print("\nSummary statistics:")
print(df.describe(include='all'))