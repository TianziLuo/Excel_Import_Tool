import pandas as pd
# ModifyCategories_data
def ModifyCategories_data(df):
    filter_condition = df['Modi_En'].notna()
    filtered_data = df[filter_condition]

    ModifyCategories_data = filtered_data.iloc[:, [1,0,1,2]]

    columns_to_insert = {
        'FontColor': pd.Series(['#FFFFFF'] * len(filtered_data), index=filtered_data.index),
        'BackgroundColor': pd.Series(['#1FA3DC'] * len(filtered_data), index=filtered_data.index),
        'OneCol': pd.Series([1] * len(filtered_data), index=filtered_data.index)
        }

    insert_index =2
    for col_name, col_data in columns_to_insert.items():
        ModifyCategories_data.insert(insert_index, col_name, col_data)
        insert_index += 1
    return ModifyCategories_data

# ModifyItem_data
def ModifyItem_data(df):
    ModifyItem_data = df.iloc[:, [0,4,6,3,4,5]]
    columns_to_insert = {
        'FontColor': pd.Series(['#FFFFFF'] * len(df), index=df.index),
        'BackgroundColor': pd.Series(['#1FA3DC'] * len(df), index=df.index),
        'OneCol1': pd.Series([1] * len(df), index=df.index),
        }

    insert_index = 4
    for col_name, col_data in columns_to_insert.items():
        ModifyItem_data.insert(insert_index, col_name, col_data)
        insert_index += 1
    return ModifyItem_data












