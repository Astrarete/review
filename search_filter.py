import pandas as pd
import os
import re

articles_df = pd.read_excel("searches.xlsx")
def extract_words(row):
    # Function to clean and split text into words
    def clean_and_split(text):
        # Replace punctuation with a space and convert to lowercase
        cleaned_text = re.sub(r'[^\w\s]', ' ', text.lower())
        # Split text into words
        return cleaned_text.split()
    # Clean and split words from each specified column
    title_words = clean_and_split(str(row['Article Title']))
    author_keywords_words = clean_and_split(str(row['Author Keywords']))
    keywords_plus_words = clean_and_split(str(row['Keywords Plus']))
    abstract_words = clean_and_split(str(row['Abstract']))
    # Combine all words into a single list
    all_words = title_words + author_keywords_words + keywords_plus_words + abstract_words
    # Remove any empty strings and strip whitespace from the list
    all_words = [word.strip() for word in all_words if word.strip()]
    return all_words
# Apply the extract_words function to each row and store the result in a temporary variable
temp_words = articles_df.apply(lambda row: extract_words(row), axis=1)
# Create a copy of the DataFrame
articles_df_copy = articles_df.copy()
# Assign the temporary variable to the 'All Words' column of the copied DataFrame using .loc[]
articles_df_copy.loc[:, 'All Words'] = temp_words

# Assign the modified DataFrame back to the original DataFrame
df = articles_df_copy
#df.to_excel("searches_all_words.xlsx")### for some reason, extracting into a table and reading it breaks the code
#df = pd.read_excel("searches_all_words.xlsx")

nutrients = {
    "N" : ["n", "nitrogen", "nitrate", "nitrous", "ammonium", "ammonia", "NH", "urea"],
    "P" : ["p", "phosphorus", "fosforus", "phosphate", "fosfate"],
    "K" : ["k", "potassium"],
    "S" : ["s", "sulphur", "sulfur", "sulphate", "sulfate", "sulfide", "sulphide", "sulphuric", "sulfuric"],
    "Ca" : ["ca", "calcium", "carbonate"],
    "Mg" : ["mg", "magnesium"],
    "Cu" : ["cu", "copper"],
    "Zn" : ["zn", "zinc"],
    "Fe" : ["fe", "iron"],
    "B" : ["b", "boron"],
    "Mn" : ["mn", "manganese", "manganate"],
    "Mo" : ["mo", "molybdenum"],
    "Ni" : ["ni", "nickel"],
    "Al" : ["al", "aluminium", "aluminum", "aluminate", "aluminosilicate"]
}
interactions = {
    "availability +": ["solubility", "mobilise", "mobilize", "mobilisation", "mobilization", "solute", "dissolve", "mineralization", "mineralisation", "mineralise", "mineralize", "solubilization", "solubilisation", "leaching", "desorption", "dissolution", "release", "extraction"], #"availability", "bioavailability",
    "availability -": ["fixation", "immobilisation", "immobilization", "adsorption", "binding", "sorption", "precipitation", "absorption", "retention", "sorbing", "association"],
    "pH": ["ph", "acidity", "protons", "acid", "acidic"],
    "SOM": ["c", "carbon", "organic", "humus", "humic"]
}
# Initialize a dictionary to store tags for each article
article_tags = {}
### Tag assignment loop
for index, article_keywords in enumerate(df['All Words']):
    article_tags[index] = set()
    for nutrient, keywords in nutrients.items():
        if any(keyword in article_keywords for keyword in keywords):
            article_tags[index].add(nutrient)
    for interaction, keywords in interactions.items():
        if any(keyword in article_keywords for keyword in keywords):
            article_tags[index].add(interaction)

# Add tags to the DataFrame
for index, tags in article_tags.items():
    df.at[index, 'tags'] = ', '.join(tags)  # Convert set to string for DataFrame
# Create a separate DataFrame for dropped articles
dropped_df = df[(df['tags'] == '') | (df['tags'] == 'None')]
# Drop articles from the original DataFrame if there are no tags or "None"
df = df.drop(dropped_df.index)
dropped_df.to_excel("searches_dropped.xlsx")
df.to_excel("searches_clean.xlsx")

###nutrient analysis BASIC
# Create a matrix for nutrients
matrix = [[f'{nutrient1}-{nutrient2}' for nutrient1 in nutrients.keys()] for nutrient2 in nutrients.keys()]
# Create a pandas DataFrame from the matrix
matrix_nutrients = pd.DataFrame(matrix, index=nutrients.keys(), columns=nutrients.keys())
# Print the DataFrame
#print(matrix_nutrients)
# Loop through each row and column in the matrix_df
for nutrient_row, nutrient_row_values in matrix_nutrients.iterrows():
    for nutrient_column, _ in nutrient_row_values.items():
        # Get the keywords for the current nutrients
        nutrient_row_keywords = nutrients.get(nutrient_row, [])
        nutrient_column_keywords = nutrients.get(nutrient_column, [])
        # Initialize the counter for the current cell to zero
        cell_count = 0
        # Iterate over each article's list of words
        for article_keywords in df['All Words']:
            # Check if any keyword from the first nutrient is present in the article's keywords
            if any(keyword in article_keywords for keyword in nutrient_row_keywords):
                # Check if any keyword from the second nutrient is present in the article's keywords
                if any(keyword in article_keywords for keyword in nutrient_column_keywords):
                    # Increment the counter for the current cell
                    cell_count += 1
        # Write the count to the corresponding cell in the matrix DataFrame
        matrix_nutrients.at[nutrient_row, nutrient_column] = cell_count
# Print the updated first DataFrame with counts
print(matrix_nutrients)
matrix_nutrients.to_excel('matrix_nutrients.xlsx', index=True)

###nutrient analysis WITH INTERACTIONS
# Initialize an empty dictionary to store matrices for each interaction
interaction_matrices = {}

# Loop through each interaction type and its keywords
for interaction, keywords in interactions.items():
    # Create a matrix for the current interaction
    matrix = [[f'{nutrient1}-{nutrient2}' for nutrient1 in nutrients.keys()] for nutrient2 in nutrients.keys()]
    # Create a pandas DataFrame from the matrix
    matrix_df = pd.DataFrame(matrix, index=nutrients.keys(), columns=nutrients.keys())

    # Loop through each row and column in the matrix DataFrame
    for nutrient_row, nutrient_row_values in matrix_df.iterrows():
        for nutrient_column, _ in nutrient_row_values.items():
            # Get the keywords for the current nutrients
            nutrient_row_keywords = nutrients.get(nutrient_row, [])
            nutrient_column_keywords = nutrients.get(nutrient_column, [])

            # Initialize the counter for the current cell to zero
            cell_count = 0

            # Iterate over each article's list of words
            for article_keywords in df['All Words']:
                # Check if any keyword from the first nutrient is present in the article's keywords
                if any(keyword in article_keywords for keyword in nutrient_row_keywords):
                    # Check if any keyword from the second nutrient is present in the article's keywords
                    if any(keyword in article_keywords for keyword in nutrient_column_keywords):
                        # Check if any keyword from the interaction is present in the article's keywords
                        if any(keyword in article_keywords for keyword in keywords):
                            # Increment the counter for the current cell
                            cell_count += 1

            # Write the count to the corresponding cell in the matrix DataFrame
            matrix_df.at[nutrient_row, nutrient_column] = cell_count

    # Store the matrix DataFrame for the current interaction in the dictionary
    interaction_matrices[interaction] = matrix_df

# Save each interaction matrix to Excel
for interaction, matrix_df in interaction_matrices.items():
    matrix_df.to_excel(f'matrix_{interaction}.xlsx', index=True)


###### OLD factor and interaction analysis

###availability-nutrient analysis
# Create a matrix for factors and nutrients
matrix = [[f'{nutrient}-{interaction}' for nutrient in nutrients.keys()] for interaction in interactions.keys()]
# Create a pandas DataFrame from the matrix
matrix_interactions = pd.DataFrame(matrix, index=interactions.keys(), columns=nutrients.keys())
# Print the DataFrame
#print(matrix_factors)

# Loop through each row and column in the matrix_df
for interaction, _ in interactions.items():
    for nutrient, _ in nutrients.items():
        # Get the keywords for the current interaction and nutrient
        interaction_keywords = interactions.get(interaction, [])  # Corrected variable name
        nutrient_keywords = nutrients.get(nutrient, [])
        # Initialize the counter for the current cell to zero
        cell_count = 0
        # Iterate over each article's list of words
        for index, article_keywords in enumerate(df['All Words']):
            # Check if at least one interaction keyword and at least one nutrient keyword are present in the article's keywords
            if any(keyword in article_keywords for keyword in interaction_keywords):
                if any(keyword in article_keywords for keyword in nutrient_keywords):
                    # Increment the counter for the current cell
                    cell_count += 1
        # Write the count to the corresponding cell in the matrix DataFrame
        matrix_interactions.at[interaction, nutrient] = cell_count  # Corrected index

# Print the updated matrix DataFrame
print(matrix_interactions)
matrix_interactions.to_excel('matrix_interactions.xlsx', index=True)

"""
###factor-nutrient analysis
# Create a matrix for factors and nutrients
matrix = [[f'{nutrient}-{factor}' for nutrient in nutrients.keys()] for factor in factors.keys()]
# Create a pandas DataFrame from the matrix
matrix_factors = pd.DataFrame(matrix, index=factors.keys(), columns=nutrients.keys())
# Print the DataFrame
#print(matrix_factors)
# Loop through each row and column in the matrix_df
for factor, _ in factors.items():
    for nutrient, _ in nutrients.items():
        # Get the keywords for the current factor and nutrient
        factor_keywords = factors.get(factor, [])
        nutrient_keywords = nutrients.get(nutrient, [])
        # Initialize the counter for the current cell to zero
        cell_count = 0
        # Iterate over each article's list of words
        for index, article_keywords in enumerate(df['All Words']):
            # Check if at least one factor keyword and at least one nutrient keyword are present in the article's keywords
            if any(keyword in article_keywords for keyword in factor_keywords):
                if any(keyword in article_keywords for keyword in nutrient_keywords):
                    # Increment the counter for the current cell
                    cell_count += 1
        # Write the count to the corresponding cell in the matrix DataFrame
        matrix_factors.at[factor, nutrient] = cell_count
# Print the updated matrix DataFrame
print(matrix_factors)
matrix_factors.to_excel('matrix_factors.xlsx', index=True)
"""