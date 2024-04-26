import pandas as pd
## Group_data
def Group_data(df):
    unique_Menu_Name_En = df['Menu_En'].dropna().unique()
    unique_Menu_Name_Cn = df['Menu_Cn'].dropna().unique()
    unique_Menu_id = df['Menu_id'].dropna().unique()
    Group_data =  [list(unique_Menu_Name_En),
                list(unique_Menu_id),
                list(unique_Menu_Name_En),
                list(unique_Menu_Name_Cn)]

    none_row = [None] * len(unique_Menu_Name_En)
    one_row = [1] * len(unique_Menu_Name_En) 

    for _ in range(2):
        Group_data.insert(2, none_row.copy())
    Group_data.insert(4, one_row.copy())

    Group_data = pd.DataFrame(list(zip(*Group_data)))
    return Group_data

## Categories_data
def Categories_data(df):
    filter_condition = df['Cate_En'].notna()
    filtered_data = df[filter_condition]

    Categories_data = filtered_data.iloc[:, [0,4,3,4,5]]

    columns_to_insert = {
        'FontColor': pd.Series(['#FFFFFF'] * len(filtered_data), index=filtered_data.index),
        'BackgroundColor': pd.Series(['#1FA3DC'] * len(filtered_data), index=filtered_data.index),
        'OneCol': pd.Series([1] * len(filtered_data), index=filtered_data.index)
        }

    insert_index =3
    for col_name, col_data in columns_to_insert.items():
        Categories_data.insert(insert_index, col_name, col_data)
        insert_index += 1
    return Categories_data

## MenuItems_data
def MenuItems_data(df):
    MenuItems_data = df.iloc[:, [7,9,0,6,6,7,8]]
    columns_to_insert = {
        'FontColor': pd.Series(['#FFFFFF'] * len(df), index=df.index),
        'BackgroundColor': pd.Series(['#355d6e'] * len(df), index=df.index),
        'OneCol1': pd.Series([1] * len(df), index=df.index),
        'OneCol2': pd.Series([1] * len(df), index=df.index)
    }

    insert_index = 4
    for col_name, col_data in columns_to_insert.items():
        MenuItems_data.insert(insert_index, col_name, col_data)
        insert_index += 1

    return MenuItems_data

## CategoryToItem_data
def CatalogyToItem_data(df):
    CatalogyToItem_data = df.iloc[:, [3,6,6]]
    return CatalogyToItem_data 