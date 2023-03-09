import os
import pandas as pd

class Enrollment_Form:
    
    def __init__(self):
        self.path = 'C:\\Users\\tsingleton\\Desktop\\CALSTART\\2. Sus Fleets\\4. Version Control Github\\SF-Github-Repo\\Report\\Fleet Enrollment Forms'
        
    def create_dict(self): 
        '''
        read in each file at a time.
        open up file and read in specific sheets. a dictionary of dfs will automatically be created per file.
        '''
        #list of sheet names that contain the info in a more database friendly format
        sheets = ['Database User', 'Database ZEV', 'Database BYD', 'Database RYD', 'Database Fuel']
        #dict copmprehension that stores a dictionary for each file in self.path. Each value will be a dict that contains a series of dfs
        #one for each ot the sheets listed in the sheets list.
        df_dict = {file.split(' Sustainable')[0]: pd.read_excel(self.path+'\\'+file, sheet_name = sheets) for file in os.listdir(self.path)}
        
        return df_dict
    
    def merge_df(self, df_dict):
        '''
        take in dictionary of dfs
        merge them into one df per file
        '''
        new_dict = {}
        for key, value in df_dict.items():
            #final_df = pd.DataFrame()
            new_dict[key] = pd.concat([df for df in df_dict[key].values()])
            
        return new_dict
            

