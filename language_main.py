import os
import pandas as pd
from filedialog import select_excel_file

# Get the path to the desktop
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
# Create the full path for the SQL file
file_path = os.path.join(desktop, 'language.sql')


df_input = select_excel_file()
# Mysql language syntaxs saved in .sql 
df_language = pd.read_excel(df_input, sheet_name='Language_list')
language_first_column = df_language.iloc[:, 0]

# Check if the column is not empty
if not language_first_column.empty:
    # Define the SQL file path for variants
    file_path_variant = os.path.join(desktop, 'variant.sql')

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(
        "INSERT INTO `userve`.`language_lists` (`id`, `ref_id`, `name`, `form_slug`, `language`, `created_at`, `updated_at`)  VALUES\n")
    for i, item in enumerate(language_first_column):
    
        f.write(f"{item}")  
        
        if i == len(language_first_column) - 1:
            f.write(";\n")
        else:
             f.write(",\n")
