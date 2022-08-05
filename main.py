import csv

from openpyxl import load_workbook, Workbook
import pandas as pd
import os
import time

from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import decimal
from pathlib import Path
from datetime import date
import glob


def format_number(num):
    try:
        dec = decimal.Decimal(num)
    except:
        return num

    val = dec
    return val

def latest_edited_file_glob(pathToDir):
    pathTest = os.path.dirname(__file__)

    if pathToDir == 'ori':
        list_of_files = glob.glob(
            os.path.join(pathTest, "data", "ori", '*'))  # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getctime)
        print("glob latest_file:" + latest_file)
    elif pathToDir == 'changed':
        list_of_files = glob.glob(
            os.path.join(pathTest, "data", "changed", '*'))  # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getctime)
        print("glob latest_file:" + latest_file)
    else:
        list_of_files = glob.glob(
            os.path.join(pathTest, "color_diffs", "Spring 2022", '*'))  # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getctime)
        print("glob latest_file:" + latest_file)

    return latest_file


def read_csv(filename):
    #pathTest = os.path.dirname(__file__)
    #latest_edited_file = max([f for f in os.scandir(os.path.join(pathTest, "diffs", "Spring 2022"))], key=lambda x: x.stat().st_mtime).name
    #print(latest_edited_file)
    latest_edited_file_glob('test') # to get date

    path = os.path.dirname(__file__)
    print('File name :    ', os.path.basename(__file__))
    print('Directory Name:     ', path)

    # titles_df = pd.read_csv(r'C:\Users\ACER\PycharmProjects\diff_calculator\data\header\header.CSV', header=0)
    # print(titles_df.shape)

    if filename == 'ori':
        #latest_edited_file = max([f for f in os.scandir(os.path.join(path, "data", "ori"))], key=lambda x: x.stat().st_mtime).name
        #pathFile = os.path.join(path, "data", "ori", latest_edited_file)

        pathFile = latest_edited_file_glob('ori')
        #print(latest_edited_file)

        try:
            modified = os.path.getmtime(pathFile)
            year, month, day, hour = time.localtime(modified)[:-5]
            file_date = str(str(year) + '_' + str(month) + '_' + str(day))
            print("Original Date:", file_date)
            print('pathFile is:', pathFile)
        except:
            file_date = time.ctime(os.path.getmtime(pathFile))

    elif filename == 'new':

        #latest_edited_file = max([f for f in os.scandir(os.path.join(path, "data", "changed"))], key=lambda x: x.stat().st_mtime).name
        #pathFile = os.path.join(path, "data", "changed", latest_edited_file)

        pathFile = latest_edited_file_glob('changed')

        #print(latest_edited_file)

        try:
            modified = os.path.getmtime(pathFile)
            year, month, day, hour = time.localtime(modified)[:-5]
            file_date = str(str(year)+'_'+str(month)+'_'+str(day))
            print ("New Date:", file_date)
            file_date = str(str(year) + '_' + str(month) + '_' + str(day))
            print("New Date:", file_date)
        except:
            file_date = time.ctime(os.path.getmtime(pathFile))
        print('pathFile is:', pathFile)
        print(pathFile)

    #df = pd.read_csv(pathFile, header=None)
    df = pd.read_csv(pathFile, header=None)
    # df1.columns = df.columns
    # print(df.iloc[0, 66])
    # print(df)

    return df, file_date


def duplicate_entry_merger(df):
    # df.to_excel('diff_out_original.xlsx')
    df = df.applymap(str)
    df = df.groupby([df.columns[0], df.columns[7]]).agg(" ; ".join).reset_index()
    # df.to_excel('diff_out.xlsx')
    return df


if __name__ == '__main__':

    today = date.today()

    df_original_file, ori_date = read_csv('ori')
    df_modified_file, new_date = read_csv('new')

    # filtering out the dfs for CHEM
    df_original_file = df_original_file.loc[df_original_file.iloc[:, 14] == "CHEM"]
    df_modified_file = df_modified_file.loc[df_modified_file.iloc[:, 14] == "CHEM"]

    for i in range(df_original_file.shape[0]):
        for j in range(df_original_file.shape[1]):
            df_original_file.iloc[i, j] = format_number(df_original_file.iloc[i, j])

    for i in range(df_modified_file.shape[0]):
        for j in range(df_modified_file.shape[1]):
            df_modified_file.iloc[i, j] = format_number(df_modified_file.iloc[i, j])

    filtered_columns = [0, 4, 7, 15, 16, 17, 18, 19, 21, 23, 25, 27, 34, 35, 36, 38, 51, 53, 57, 66, 67, 81, 82, 83, 90,
                        91,
                        97, 110, 112]

    df_filtered_ori = df_original_file.iloc[:]
    df_filtered_mod = df_modified_file.iloc[:]
    if len(df_modified_file.columns) > 140:  # avoids applying filters to short files
        # print(len(df_modified_file.columns),'len(df_modified_file.columns)')
        df_filtered_ori = df_original_file.iloc[:, filtered_columns]
        df_filtered_mod = df_modified_file.iloc[:, filtered_columns]

    term = df_filtered_mod.iloc[0, 1]
    print('term' + term)

    Path("color_diffs/" + term).mkdir(parents=True, exist_ok=True)

    # ----------------
    df_filtered_ori = duplicate_entry_merger(df_filtered_ori)
    df_filtered_mod = duplicate_entry_merger(df_filtered_mod)
    # ----------------

    df_changes = pd.concat([df_filtered_ori.assign(type='original'), df_filtered_mod.assign(type='modified')])
    df_changes = df_changes.drop_duplicates(keep=False, subset=df_changes.columns.difference(['type']))

    df_changes.to_csv(f'data/diffs/{ori_date}_{new_date}_diff_out.csv', sep=',', index=False, header=False)
    # df_changes.to_csv('diff_out.csv', sep=',', index=False, header=False)

    # df_changes_arranged = pd.read_csv('diff_out.csv', header=None)
    df_changes_arranged = df_changes

    df_changes_arranged = df_changes_arranged.sort_values(
        by=[df_changes_arranged.columns[0], df_changes_arranged.columns[7]], ascending=True)
    df_changes_arranged.to_csv(f'data/diffs/{ori_date}_{new_date}_diff_out_arranged.csv', sep=',', index=False,
                               header=False)

    # print(df_changes_arranged.shape)
    # print(df_changes_arranged.shape[0])
    # print(df_changes_arranged.shape[1])

    j = 0
    i = 0
    c = 0
    rows_count = df_changes_arranged.shape[0]
    cols_count = df_changes_arranged.shape[1]
    # print('rows_count:' + str(rows_count))

    changed_id = []
    unique_id = []

    unique_id_from_ori = []
    unique_id_from_mod = []

    red_color = 'background-color: red'
    green_color = 'background-color: lightgreen'

    for i in range(rows_count - 1):
        if ((df_changes_arranged.iloc[i, 0] == df_changes_arranged.iloc[i + 1, 0]) and (
                df_changes_arranged.iloc[i, 7] == df_changes_arranged.iloc[i + 1, 7])):
            # print("same rows detected" + str(i))
            j = 0

            for c in range(cols_count):
                if c < 29:
                    if df_changes_arranged.iloc[i, c] != df_changes_arranged.iloc[i + 1, c]:
                        # print('--- different column' + str(c))
                        changed_id.append([i, c])
                        # changed_id.append([i + 1, c])

        else:
            j = j + 1
            if j >= 2:
                # print('unique row detected' + str(i))
                unique_id.append([i + 1, 1])
                j = 0

            elif i == rows_count - 2:
                unique_id.append([i + 2, 1])
                j = 0

            else:
                print("different rows detected" + str(i))

    wb = Workbook()
    ws = wb.active

    for r in dataframe_to_rows(df_changes_arranged, index=False, header=False):
        ws.append(r)

    fill_pattern_modified_content = PatternFill(patternType='solid', fgColor='ffaba6')
    fill_pattern_modified_content_darker = PatternFill(patternType='solid', fgColor='f03629')
    fill_pattern_unique_content = PatternFill(patternType='solid', fgColor='61bdff')

    fill_pattern_from_original = PatternFill(patternType='solid', fgColor='FFA500')
    fill_pattern_from_modified = PatternFill(patternType='solid', fgColor='66CDAA')

    # print('data/////')
    # print(ws.cell(1,30).value)

    for l in range(len(df_changes_arranged)):
        if ws.cell(l + 1, 30).value == "original":
            # print("original")
            for p in range(cols_count):
                ws.cell(l + 1, p + 1).fill = fill_pattern_from_original
        elif ws.cell(l + 1, 30).value == "modified":
            # print("modified")
            for p in range(cols_count):
                ws.cell(l + 1, p + 1).fill = fill_pattern_from_modified

    for i in range(len(changed_id)):
        ws.cell(changed_id[i][0] + 1, changed_id[i][1] + 1).fill = fill_pattern_modified_content
        ws.cell(changed_id[i][0] + 2, changed_id[i][1] + 1).fill = fill_pattern_modified_content_darker

    wb.save("color_diffs/" + term + "/" + today.isoformat() + "_output.xlsx")
