import pandas as pd

def csv_creator(df_dict, path = None):
    '''
    Thake in a dictionary and unpack it. Each key should have 1 dataframe associated with it
    '''
    if path == None:
        path = 'C:\\Users\\tsingleton\\Desktop\\CALSTART\\2. Sus Fleets\\4. Version Control Github\\SF-Github-Repo\\Code\\CSV Files\\'
    for key, value in df_dict.items():
        value.to_csv(path+key+'_CSV_Table'+'.csv')