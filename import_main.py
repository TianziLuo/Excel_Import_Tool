import os
import pandas as pd
import subprocess
import mysql.connector
from mysql.connector import Error
from Sheet_Item_Data import Group_data, Categories_data, MenuItems_data, CatalogyToItem_data
from Sheet_Modify_Data import ModifyItem_data, ModifyCategories_data
from filedialog import select_excel_file, save_excel_file


## step 1
# Get the path to the desktop
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
# Create the full path for the SQL file
file_path_necessary = os.path.join(desktop, 'necessary.sql')

# Read the data from 'china.xlsx'
df_input = select_excel_file()
df_item = pd.read_excel(df_input, sheet_name='item')
df_modi = pd.read_excel(df_input, sheet_name='modify')
df_variant = pd.read_excel(df_input, sheet_name='menu_item_variants')



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

#necessary sql syntax  
tax = input("Enter tax rate: ")

with open(file_path_necessary, 'w', encoding='utf-8') as f:
    f.write(
            f"UPDATE `userve`.`setting_tax_lists` SET `rate` = '{tax}' WHERE (`id` = '1');\n"
            "UPDATE `userve`.`menu_items` SET `is_discountable` = '1' WHERE (`id` >= '1'); \n"
            )
    
#variant sql if available
# Extract the first column of the 'menu_item_variants' DataFrame
variant_first_column = df_variant.iloc[:, 0]

# Check if the column is not empty
if not variant_first_column.empty:
    # Define the SQL file path for variants
    file_path_variant = os.path.join(desktop, 'variant.sql')
    
    # Write SQL to the file on the desktop
    with open(file_path_variant, 'w', encoding='utf-8') as f:
        f.write(
            "INSERT INTO `userve`.`menu_item_variants` (`id`, `item_id`, `name`, `price`, `extra_price`, `sort`, `created_at`, `updated_at`) VALUES\n"
        )
        for i, item in enumerate(variant_first_column):
            f.write(f"{item}")
            
            if i == len(variant_first_column) - 1:
                f.write(";\n")
            else:
                f.write(",\n")


## step 2
program_path = r"D:\uServePro menu input 1.6\MenuXlsImport.exe"

try:
    # Use subprocess to open the program
    subprocess.Popen(program_path, shell=True)
    print(f"Successfully opened the program: {program_path}")
except FileNotFoundError:
    # Handle the case where the program is not found
    print(f"Program not found: {program_path}")
except Exception as e:
    # Handle any other exceptions
    print(f"An error occurred: {e}")


mysql_workbench_path = r"C:\Program Files\MySQL\MySQL Workbench 8.0 CE\MySQLWorkbench.exe"

if os.path.exists(mysql_workbench_path):
    os.startfile(mysql_workbench_path)
    try:
        connection = mysql.connector.connect(
            host="127.0.0.1",
            port=3308,
            user="root",
            password="123456"
        )
        if connection.is_connected():
            print("Connection to MySQL database was successful!")
    except Error as e:
        print(f"Error while connecting to MySQL: {e}")
    finally:
        if 'connection' in locals() and connection.is_connected():
            connection.close()
            print("MySQL connection is closed.")
else:
    print(f"Path not found: {mysql_workbench_path}. Please verify the path and try again.")

