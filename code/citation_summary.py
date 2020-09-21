'''
UPDATE2.0
problem face:   fromat excel change overtime.
                add comment
                add tutorial how to

In main file name contain 
1) code contain all code (.py)
2) resouce 
3) citation summary (program create)
4) click.bat (UI)
5) how to use.txt
'''

import pandas as pd
import glob
import json
import numpy as np
from datetime import datetime
from xlrd import open_workbook

#Check incoming file and duplicate.
def check_file_in_folder(folder_name, file_format,raw_file):
    used_files = glob.glob(f'../{folder_name}/{file_format}') #check resource file that user input.
    list_result = glob.glob(raw_file) #duplicate file named raw data to use in this program.
    return used_files, list_result

#Check that the last columns is Q4 or not. Inorder to Sum all that year.
def check_last_coloum(result):
    result_column_list = list(result.columns)
    if result_column_list[-2][-2:] == 'Q4':
        result_column_list[-1] = f'Total {result_column_list[-2][0:4]}'
        result.columns = result_column_list
        return result

#Check log file to find last updated.
def check_last_file_update(js,used_files,folder_name):
    keys_js = [i for i in js.keys()]
    last_key = keys_js[-1]
    js[last_key][1]
    index = used_files.index(f'../{folder_name}\\{js[last_key][1]}')+1
    return index

#Drop first columns (index column) cause some have index, some not.
def check_file(df):
    if df.iloc[:,0].dtype == np.float64 :
         df.drop(df.columns[0], axis=1, inplace = True)
    else :
        pass
    return df

#Get input from final result file.
def read_file_result(raw_file):
    result = pd.read_excel(raw_file)
    result.drop(result.columns[0], axis=1, inplace = True)
    return result
    
#Get last file to make final result.
def read_file_new(used_files,skiprows_number,js,folder_name):
    df = pd.read_excel(used_files[check_last_file_update(js,used_files,folder_name)], skiprows=skiprows_number) #skipheader
    df.drop(df.tail(2).index, inplace = True) #drop last 2 row which is document tail. 
    check_file(df)
    return df

#writing excel.
def write_excel(df,raw_file,desktop_path):
    excel_data = df.to_excel (desktop_path, header=True) #citation summary
    excel_use = df.to_excel (raw_file, header=True) #file to use
    return excel_data, excel_use

#if there are to final citaion file run this part.
def first_run(folder_name,file_format,skiprows_number1,skiprows_number2,used_files, list_result,raw_file,desktop_path):
    '''
    gatering all data in first time if file names "citation_summary" doesn't exist.
    '''
    df1 = pd.read_excel(used_files[0], skiprows = skiprows_number1) #read excel oldest file possible.
    df2 = pd.read_excel(used_files[1], skiprows = skiprows_number2) #read more present excel file.
    df1.drop(df1.tail(2).index, inplace = True) # drop last 2 row which is document tail. 
    df2.drop(df2.tail(2).index, inplace = True) # drop last 2 row which is document tail. 
    check_file(df1) # drop first column if index popup in excel. 
    check_file(df2) # drop first column if index popup in excel. 
    result = pd.merge(df1, df2, how = 'outer', on = ['Title', 'Authors', 'Year', 'Scopus Source title'])
    result.fillna(0,inplace=True) #there are some document that system stop tracking, so it show up as N/A.
    col_name1 = used_files[0][len(folder_name)+4:-5] #use file name as column name (ex 2020-Q1) for older file.
    col_name2 = used_files[1][len(folder_name)+4:-5] #use file name as column name (ex 2020-Q1) for more present file.
    result[col_name2] = result['Citations_y'] - result['Citations_x'] #find the difference of 2 continus year.
    result[col_name2] = np.where(result[col_name2] < 0, 0, result[col_name2]) #if the differenciate is less that 0 then 0.
    result.drop(columns = ['Volume_x', 'Issue_x', 'Pages_x', 'Citations_y'], inplace = True) #drop unsatisfy column.
    result.rename(columns = {'Citations_x' : col_name1,  'Volume_y': 'Volume', 'Issue_y':'Issue', 'Pages_y':'Pages'}, inplace = True) #rename columns.
    result = result[['Title', 'Authors', 'Year', 'Scopus Source title','Volume','Issue','Pages',col_name1,col_name2]] #to sort column.
    result['Total'] = result[col_name1] + result[col_name2] #citation number of that Q.
    check_last_coloum(result) #check if Q4 or not.
    return write_excel(result,raw_file,desktop_path)

#if it is end of Q. run this part to make total of year summation.
def finished_year(folder_name,file_format,skiprows_number,used_files, list_result,js,raw_file,desktop_path):
    result = pd.merge(read_file_result(raw_file), read_file_new(used_files,skiprows_number,js,folder_name), how = 'outer', on = ['Title', 'Year', 'Scopus Source title','Volume', 'Issue', 'Pages'])
    result.drop(columns = ['Authors_y'], inplace = True)
    result.rename(columns= {'Authors_x' : 'Authors'}, inplace = True)
    result.fillna(0,inplace=True)
    col_name = used_files[check_last_file_update(js,used_files,folder_name)][len(folder_name)+4:-5]
    result[col_name] = result['Citations'] - result.iloc[:,-2]
    result[col_name] = np.where(result[col_name] < 0, 0, result[col_name])
    result.drop(columns = ['Citations'], inplace = True)
    result['Total'] = result.iloc[:, 7:].sum(axis = 1) #make new column from sumation during year (4Q)
    check_last_coloum(result)
    return write_excel(result,raw_file,desktop_path)

#if it is not end of Q. run this part to continue.
def not_finished_year(folder_name,file_format,skiprows_number,used_files, list_result,js,raw_file,desktop_path):
    result = pd.merge(read_file_result(raw_file), read_file_new(used_files,skiprows_number,js,folder_name), how = 'outer', on = ['Title', 'Year', 'Scopus Source title','Volume', 'Issue', 'Pages'])
    result.drop(columns = ['Authors_y'], inplace = True)
    result.rename(columns= {'Authors_x' : 'Authors'}, inplace = True)
    result.fillna(0,inplace=True)
    col_name = used_files[check_last_file_update(js,used_files,folder_name)][len(folder_name)+4:-5]
    result[col_name] = result['Citations'] - result['Total']
    result[col_name] = np.where(result[col_name] < 0, 0, result[col_name])
    result.drop(columns = ['Total'], inplace = True)
    result.drop(columns = ['Citations'], inplace = True)
    result['Total'] = result.iloc[:, 7:].sum(axis = 1)
    check_last_coloum(result)
    return write_excel(result,raw_file,desktop_path)

#read log file to get last file update infomation.
def read_json_stamp(log_filname):
    with open("log.json", "r") as jsonFile:
        js = json.load(jsonFile)
    return js

#update/create log file using json.
def write_json_stamp(js):
    with open("log.json", "w") as jsonFile:
        json.dump(js, jsonFile)

#find amount of row to skip by locate where Title is.
def find_skiprows_number(file_name):
    book = open_workbook(file_name)
    for sheet in book.sheets():
        for rowidx in range(sheet.nrows):
            row = sheet.row(rowidx)
            for colidx, cell in enumerate(row):
                if cell.value == "Title" :
                    return rowidx

#run all code
def click():
    #find folder that match criteria.
    folder_name = 'resource'
    file_format = '20*Q[1-4].xlsx'
    raw_file = 'raw_data.xlsx'
    log_filname = 'log.json'
    desktop_path = r'../citations_summary.xlsx' 
  
    #check if it is result file or not.
    used_files, list_result = check_file_in_folder(folder_name, file_format,raw_file)
    dt = str(datetime.now())
    try:
        js = read_json_stamp(log_filname)
        js_keys = [i for i in js.keys()]
        last_js_key = js_keys[-1]
    except Exception:
        pass
    
    #if rawdata have dummy file. if not run firstrun part to create it.
    if len(list_result) == 0:
        js = {}
        skiprows_number1 = find_skiprows_number(used_files[0])
        skiprows_number2 = find_skiprows_number(used_files[1])
        first_run(folder_name,file_format,skiprows_number1,skiprows_number2,used_files, list_result,raw_file,desktop_path)
        js[0] = [dt, used_files[1][len(folder_name)+4:]]
        write_json_stamp(js)

    #make total columns.
    elif js[last_js_key][1][-7:-5] == 'Q4':
        skiprows_number = find_skiprows_number(used_files[check_last_file_update(js,used_files,folder_name)])
        finished_year(folder_name,file_format,skiprows_number,used_files, list_result,js,raw_file,desktop_path)
        js[len(js)] = [dt, used_files[check_last_file_update(js,used_files,folder_name)][len(folder_name)+4:]]
        write_json_stamp(js)

    #check file in result that it documented (log) if yes create summary form last update (no more calculation).
    elif used_files[-1][len(folder_name)+4:] == js[last_js_key][1]:
        result = read_file_result(raw_file)
        result.to_excel (desktop_path, header=True)

    #run not finish year code to calculate diff without create total columns.
    else:
        skiprows_number = find_skiprows_number(used_files[check_last_file_update(js,used_files,folder_name)])
        not_finished_year(folder_name,file_format,skiprows_number,used_files, list_result,js,raw_file,desktop_path)
        js[len(js)] = [dt, used_files[check_last_file_update(js,used_files,folder_name)][len(folder_name)+4:]]
        write_json_stamp(js)
            
if __name__ == "__main__":
    click()