import pandas as pd
import os
import re
# Directory containing Excel files
dir = "/home/anst7699/Documents/Academia/Thesis/searches"
# List to store DataFrames from each Excel file
dfs = []
# Iterate over each file in the directory
for filename in os.listdir(dir):
    if filename.endswith('.xls') or filename.endswith('.xlsx'):
        # Construct the full file path
        file_path = os.path.join(dir, filename)
        # Read the Excel file into a DataFrame and append it to the list
        df = pd.read_excel(file_path)
        dfs.append(df)
# Concatenate all DataFrames in the list into one DataFrame
dataframe = pd.concat(dfs, ignore_index=True)
#dataframe.to_excel("searches_raw_concat.xlsx")
# Now combined_df contains all data from the Excel files
# Remove duplicate rows from the combined DataFrame
articles_df = dataframe.drop_duplicates()
articles_df.reset_index(drop=True, inplace=True)
articles_df.to_excel("searches.xlsx")
print(articles_df)
