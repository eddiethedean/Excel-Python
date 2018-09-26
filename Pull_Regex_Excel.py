
#easy file exlporer popup
from tkinter import Tk
from tkinter.filedialog import askopenfilename
def get_excel_path():
    Tk().withdraw() #prevents full GUI from appearing
    return askopenfilename(filetypes=[('Excel Files','*.csv;*.xlsx;*.xlsm;*.xltx;*.xltm')],
           title='Select your Excel File and Click Open')
           
import os
import pandas as pd
from openpyxl import load_workbook
def upload_spreadsheet(path, active_sheet_only=True):
    '''Returns pandas dataframe. Returns empty dataframe is fails'''
    #check if file exists
    if not os.path.isfile(path):
        return pd.DataFrame()
        
    #check file type
    period_index = path.rfind('.')
    file_type = path[period_index:].lower()
    
    if file_type=='.csv':
        df = pd.read_csv(path)
    elif file_type in ['.xlsx','.xlsm','.xltx','.xltm']:
        wb = load_workbook(path)
        #convert sheets to pandas dataframe
        if active_sheet_only:
            df = pd.DataFrame(wb.active.values)
        else:
            #combine all sheets into one dataframe
            frames = [pd.DataFrame(sheet.values) for sheet in wb.worksheets]
            df = pd.concat(frames)
    else:
        return pd.Dataframe()
            
    wb.close()
    return df
        
def combine_dataframe(dataframe):
    '''Converts pandas dataframe into one list'''
    return [cell for cell in [dataframe[i] for i in dataframe]]
    

import re
def text_search(text, regex, dedupe=True):
    re_matches = re.findall(regex, str(text))
    if re_matches:
        #remove all but first match if tuples
        re_matches = [y[0] if type(y)==tuple else y for y in re_matches]
        if dedupe:
            return list(set(re_matches))
        return list(re_matches)
    return []
    

import subprocess
def copy_text(text):
    subprocess.run(['clip.exe'], input=text.strip().encode('ascii',
'ignore'), check=True)
    
    
def main():
    file_path = get_excel_path()
    if file_path == '':return
    df = upload_spreadsheet(file_path, active_sheet_only=False)
    #convert to a list then join into string
    text = ' '.join([str(x) for x in combine_dataframe(df)])
    pattern = r'([0-9]{9,14})'
    matches = text_search(text, pattern)
    copy_text('\n'.join(matches))
    print(len(matches), 'matches added to your clipboard.')
    
if __name__=='__main__':
    main()
