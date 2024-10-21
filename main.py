import pandas as pd
from openpyxl import load_workbook
from Sheet_Item_Data import Group_data, Categories_data, MenuItems_data, CatalogyToItem_data
from Sheet_Modify_Data import ModifyItem_data, ModifyCategories_data
from filedialog import select_excel_file, save_excel_file

# Read the data from 'china.xlsx'
df_input = select_excel_file()
# print(df_input)
df_item = pd.read_excel(df_input, sheet_name='item')
df_modi = pd.read_excel(df_input, sheet_name='modify')

# Process data using custom functions
processed_data = {
    'Group_data': Group_data(df_item),
    'Categories_data': Categories_data(df_item),
    'MenuItems_data': MenuItems_data(df_item),
    'ModifyCategories_data': ModifyCategories_data(df_modi),
    'ModifyItem_data': ModifyItem_data(df_modi),
    'CatalogyToItem_data': CatalogyToItem_data(df_item)
}

# Write processed data to a new Excel file 'china_final.xlsx'
# print(df_output)
with pd.ExcelWriter(save_excel_file(), engine='openpyxl') as writer:
    for sheet_name, table_data in processed_data.items():
        # Write each DataFrame to a separate sheet
        table_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Open the workbook and select the sheet to delete the first row
        workbook = writer.book
        worksheet = workbook[sheet_name]
        
        # Delete the first row (header)
        worksheet.delete_rows(1)
        
# Mysql variants syntaxs saved in .sql 
df_variant = pd.read_excel(df_input, sheet_name='menu_item_variants')
variant_first_column = df_variant.iloc[:, 0]

with open('variant.sql', 'w', encoding='utf-8') as f:
    f.write("INSERT INTO `userve`.`menu_item_variants` (`id`, `item_id`, `name`, `price`, `extra_price`, `sort`, `created_at`, `updated_at`)  VALUES\n")
    for i, item in enumerate(variant_first_column):
    
        f.write(f"{item}")  
        
        if i == len(variant_first_column) - 1:
            f.write(";\n")
        else:
            f.write(",\n")

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
