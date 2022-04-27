import csv

import openpyxl
import pandas as pd
import os
import time

from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import decimal


def format_number(num):
    try:
        dec = decimal.Decimal(num)
    except:
        return num

    val = dec
    return val

def read_csv(filename):
    path = os.path.dirname(__file__)
    print('File name :    ', os.path.basename(__file__))
    print('Directory Name:     ', path)

    # titles_df = pd.read_csv(r'C:\Users\ACER\PycharmProjects\diff_calculator\data\header\header.CSV', header=0)
    # print(titles_df.shape)

    if filename == 'ori':
        pathFile = os.path.join(path, "data","ori","CLSSCHED.CSV")

        try:
            modified = os.path.getmtime(pathFile)
            year, month, day, hour = time.localtime(modified)[:-5]
            file_date = str(str(year)+'_'+str(month)+'_'+str(day))
            print ("Original Date:",file_date)
            print('pathFile is:',pathFile)
        except:
            file_date = time.ctime(os.path.getmtime(pathFile))
    elif filename == 'new':
        pathFile = os.path.join(path, "data","changed","CLSSCHED.CSV")
        try:
            modified = os.path.getmtime(pathFile)
            year, month, day, hour = time.localtime(modified)[:-5]
            file_date = str(str(year)+'_'+str(month)+'_'+str(day))
            print ("New Date:",file_date)
        except:
            file_date = time.ctime(os.path.getmtime(pathFile))
        print('pathFile is:',pathFile)
        print(pathFile)

    df = pd.read_csv(pathFile, header=None)
    # df1.columns = df.columns
    # print(df.iloc[0, 66])
    # print(df)

    return df, file_date


if __name__ == '__main__':

    df_original_file, ori_date = read_csv('ori')
    df_modified_file, new_date = read_csv('new')

    # filtering out the dfs for CHEM
    df_original_file = df_original_file.loc[df_original_file.iloc[:,14] == "CHEM"]
    df_modified_file = df_modified_file.loc[df_modified_file.iloc[:,14] == "CHEM"]

    #print('CHEMMMKMMMMM')
    #print(df_original_file.shape[0])
    #print(df_original_file.shape[1])

    for i in range(df_original_file.shape[0]):
        for j in range(df_original_file.shape[1]):
            df_original_file.iloc[i,j] = format_number(df_original_file.iloc[i,j])

    for i in range(df_modified_file.shape[0]):
        for j in range(df_modified_file.shape[1]):
            df_modified_file.iloc[i,j] = format_number(df_modified_file.iloc[i,j])

    print('lenght of columns in new is', len(df_modified_file.columns))

    filtered_columns = [0, 4, 7, 15, 16, 17, 18, 19, 21, 23, 25, 27, 34, 35, 36, 38, 51, 53, 57, 66, 67, 81, 82, 83, 90,
                        91,
                        97, 110, 112]

    """ori_column = df_original_file.shape[1]
    ori_rows = df_original_file.shape[0]

    print('orginal data : columns ' + str(ori_column) + ' rows' + str(ori_rows))

    mod_column = df_modified_file.shape[1]
    mod_rows = df_modified_file.shape[0]

    print('mod data : columns ' + str(mod_column) + ' rows' + str(mod_rows))

    #df_changes_list = pd.DataFrame(columns=['rows','columns'])

    if ori_column == mod_column and ori_rows == mod_rows:
        for i in range(0, ori_column - 1 ):
            for j in range(0, ori_rows-1 ):
                if df_original_file.iloc[j,i] != df_modified_file.iloc[j,i]:
                    df_changes_list.append()"""

    # ne = (df_modified_file != df_original_file).any(1)
    # print(ne)

    df_filtered_ori = df_original_file.iloc[:]
    df_filtered_mod = df_modified_file.iloc[:]
    if len(df_modified_file.columns) > 140: #avoids applying filters to short files
        print(len(df_modified_file.columns),'len(df_modified_file.columns)')
        df_filtered_ori = df_original_file.iloc[:, filtered_columns]
        df_filtered_mod = df_modified_file.iloc[:, filtered_columns]

    # output a gives all the differences and new rows

    df_changes = pd.concat([df_filtered_ori, df_filtered_mod]).drop_duplicates(keep=False)

    df_changes.to_csv(f'data/diffs/{ori_date}_{new_date}_diff_out.csv', sep=',', index=False, header=False)
    #df_changes.to_csv('diff_out.csv', sep=',', index=False, header=False)

    #df_changes_arranged = pd.read_csv('diff_out.csv', header=None)
    df_changes_arranged = df_changes

    df_changes_arranged = df_changes_arranged.sort_values(
        by=[df_changes_arranged.columns[0], df_changes_arranged.columns[7]], ascending=True)
    df_changes_arranged.to_csv(f'data/diffs/{ori_date}_{new_date}_diff_out_arranged.csv', sep=',', index=False, header=False)

    print(df_changes_arranged.shape)
    print(df_changes_arranged.shape[0])
    print(df_changes_arranged.shape[1])

    j = 0
    i = 0
    c = 0
    rows_count = df_changes_arranged.shape[0]
    cols_count = df_changes_arranged.shape[1]
    print('rows_count:' + str(rows_count))

    changed_id = []
    unique_id = []
    red_color = 'background-color: red'
    green_color = 'background-color: lightgreen'

    for i in range(rows_count - 1):
        if ((df_changes_arranged.iloc[i, 0] == df_changes_arranged.iloc[i + 1, 0]) and (
                df_changes_arranged.iloc[i, 7] == df_changes_arranged.iloc[i + 1, 7])) :
            print("same rows detected" + str(i))
            j = 0

            for c in range(cols_count):
                if df_changes_arranged.iloc[i, c] != df_changes_arranged.iloc[i + 1, c]:
                    print('--- diferent column' + str(c))
                    changed_id.append([i, c])
                    # changed_id.append([i + 1, c])

        else:
            j = j + 1
            if j >= 2:
                print('unique row detected' + str(i))
                unique_id.append([i+1, 1])
                j = 0

            elif i == rows_count - 2:
                unique_id.append([i + 2, 1])
                j = 0

            else:
                print("different rows detected" + str(i))

    print(changed_id)
    print(unique_id)

    #df_changes_arranged.to_excel('out_excel.xlsx', index=False, header=False)
    df_filtered_ori.to_csv('short_form_out.csv', index=False, header=False)

    '''changed_id_excel = []
    unique_id_excel = []
    tempString = ''

    for z in range(len(changed_id)):
        temp_x = colnumbertocolname(changed_id[z][1])
        temp_y = changed_id[z][0]
        tempString = temp_x + str(temp_y)
        changed_id_excel.append(tempString)

    print(changed_id_excel)'''

    #wb = openpyxl.load_workbook('out_excel.xlsx')
    wb = Workbook()
    ws = wb.active

    for r in dataframe_to_rows(df_changes_arranged, index=False, header=False):
        ws.append(r)

    fill_pattern_modified_content = PatternFill(patternType='solid', fgColor='ffaba6')
    fill_pattern_modified_content_darker = PatternFill(patternType='solid', fgColor='f03629')
    fill_pattern_unique_content = PatternFill(patternType='solid', fgColor='61bdff')

    for i in range(len(changed_id)):
        ws.cell(changed_id[i][0] + 1, changed_id[i][1] + 1).fill = fill_pattern_modified_content
        ws.cell(changed_id[i][0] + 2, changed_id[i][1] + 1).fill = fill_pattern_modified_content_darker

    for t in range(len(unique_id)):
        for p in range(cols_count):
            ws.cell(unique_id[t][0], p + 1).fill = fill_pattern_unique_content

    # ws.cell(1,2).fill = fill_pattern_modified_content
    # ws['V1'].fill = fill_pattern_modified_content

    wb.save("colored_output.xlsx")
