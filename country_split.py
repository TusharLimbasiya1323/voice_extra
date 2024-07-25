import pandas as pd
import os
from datetime import datetime

# Read the Excel file into a DataFrame
df = pd.read_excel('/home/user/Documents/Audit_Rules/AWS DFS/DFS_Worakble_lusha_list2.xlsx')

# Group by 'Country' column
grouped = df.groupby('Country')

# Get today's date in YYYY-MM-DD format
today_date = datetime.today().strftime('%d-%m-%Y Extra')

# Define the directory path where you want to save the files
parent_dir = "/home/user/Desktop/AWS APJ DFS/"
directory = today_date + '/'
path = os.path.join(parent_dir, directory)

# Create the directory if it doesn't exist
os.makedirs(path, exist_ok=True)

# Iterate through each group and save as a separate Excel file
for country, group in grouped:
    # Define the output file name
    output_file_name = f'DFS_{country}_country.xlsx'
    
    # Create a new Excel writer object
    with pd.ExcelWriter(os.path.join(path, output_file_name), engine='xlsxwriter') as writer:
        # Write the group (subset of the DataFrame) to the Excel file
        group.to_excel(writer, sheet_name='Sheet1', index=False)

print("Excel files created successfully.")
