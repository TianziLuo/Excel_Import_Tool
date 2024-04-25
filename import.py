import pandas as pd
from openpyxl import load_workbook
from Sheet_Item_Data import Group_data, Categories_data, MenuItems_data, CatalogyToItem_data
from Sheet_Modify_Data import ModifyItem_data, ModifyCategories_data
from filedialog import select_excel_file

# Read the data from 'china.xlsx'
df_input = select_excel_file()
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
with pd.ExcelWriter('D:/china_final.xlsx', engine='openpyxl') as writer:
    for sheet_name, table_data in processed_data.items():
        # Write each DataFrame to a separate sheet
        table_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Open the workbook and select the sheet to delete the first row
        workbook = writer.book
        worksheet = workbook[sheet_name]
        
        # Delete the first row (header)
        worksheet.delete_rows(1)
        
