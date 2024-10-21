import pandas as pd
from filedialog import select_excel_file

df_input = select_excel_file()
# Mysql language syntaxs saved in .sql 
df_language = pd.read_excel(df_input, sheet_name='Language_list')
language_first_column = df_language.iloc[:, 0]

with open('language.sql', 'w', encoding='utf-8') as f:
    f.write("INSERT INTO `userve`.`language_lists` (`id`, `ref_id`, `name`, `form_slug`, `language`, `created_at`, `updated_at`)  VALUES\n")
    for i, item in enumerate(language_first_column):
    
        f.write(f"{item}")  
        
        if i == len(language_first_column) - 1:
            f.write(";\n")
        else:
             f.write(",\n")
