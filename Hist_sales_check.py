# -*- coding: utf-8 -*-
from enum import unique
from heapq import merge
import os.path
import select
from numpy import dtype
import pandas as pd
import os
import re
import tkinter as tk
from tkinter import filedialog
import glob
import pywintypes
import xlwings as xw
import datetime
import shutil, tempfile
from pathlib import Path

def select_file(title):
    root = tk.Tk() #This creates a root window object essentially the main GUI window.Even though you won’t display it, it’s required to initialize the GUI environment.
    root.withdraw() #This hides the root window so it doesn’t pop up awkwardly.
    file_path = filedialog.askopenfilename(filetypes=[("Excel Binary Workbook", "*.xlsb")], title=title)
    return file_path

def openas_temporary_local(original_path: str) -> Path:
    original = Path(original_path)
    tempdir = Path(tempfile.gettempdir())
    temp_copy = tempdir / original.name
    shutil.copy2(original, temp_copy)
    return temp_copy

def normalize_columns(df):
    df = df.rename(columns=lambda col: re.sub(r'[^a-zA-Z0-9 ]', '', col).strip().lower())
    df = df.rename(columns=lambda col: re.sub(r' {2,}', ' ', col))
    return df

def unprotect_sheet(workbook, sheet_name):
    target_sheet = workbook.sheets[sheet_name]
    if target_sheet.api.ProtectContents:
        target_sheet.api.Unprotect(Password="bigmoney")

def main():
    timestamp = datetime.datetime.now().strftime("_%Y-%m-%d")
    try:
        #Select WB2, WB4, WB5 and WB6 files from file browswer
        WB2_file = select_file("Select WB2 File")
        WB4_file = select_file("Select WB4 File")
        WB5_file = select_file("Select WB5 File")
        WB6_file = select_file("Select WB6 File")

        WB5_version = input(f"\nWhat version is WB5, old/new: ")

        #Open excel file on background
        app = xw.App(visible=False)
        
        #Open WB2 on temporary copy of the file to avoid autosave in Onedrive
        wb2 = app.books.open(openas_temporary_local(WB2_file))
        
        #Unprotect sheets
        unprotect_sheet(wb2, "2023")
        unprotect_sheet(wb2, "2024")
        unprotect_sheet(wb2, "2025")

        #Get the 3 year historical sales from WB2 and save as dataframe
        hist_sales_2023 = wb2.sheets['2023'].range("A3").options(pd.DataFrame, header=1, index=False, expand='table').value
        hist_sales_2024 = wb2.sheets['2024'].range("A3").options(pd.DataFrame, header=1, index=False, expand='table').value
        hist_sales_2025 = wb2.sheets['2025'].range("A3").options(pd.DataFrame, header=1, index=False, expand='table').value
        print(f"\ndone saving hist sales to dataframes")

        #Get the Last closed month in channel setup tab
        last_closed_month = str(wb2.sheets['Channel Setup'].range("K17").value)
        print(f"\nLast Closed Month: {last_closed_month}")

        #Get the Master list in masterlist tab as dataframe
        masterlist = wb2.sheets['Master List'].range("C1:E202").options(pd.DataFrame, header=1, index=False).value
        masterlist = masterlist.drop(index=0).reset_index(drop=True)

        #masterlist['Parent ASIN', 'Child ASIN', 'Title'] column names in masterlist dataframe


        #Remove non AlphaNumeric characters in the headers to avoid UTF-8 errors
        hist_sales_2023 = normalize_columns(hist_sales_2023)
        hist_sales_2024 = normalize_columns(hist_sales_2024)
        hist_sales_2025 = normalize_columns(hist_sales_2025)
        #print(f"\nDone normalizing hist sales dataframes")

        wb2.close()

        #Combine all hist sales data frames
        hist_sales_all = pd.concat([hist_sales_2023, hist_sales_2024, hist_sales_2025], ignore_index=True)
        #print(f"\nDone combining all hist sales dataframes")

        #Pivot historical sales by date to get the total units sold per month
        hist_sales_all_grouped = hist_sales_all.groupby("date", as_index=False)["total units ordered"].sum()
        #print(f"\nDone grouping hist sales by date with total sales")
        #print(f"\n{hist_sales_all_grouped}")

        #Get the unique asins with sales from wb2
        unique_asins_sold = hist_sales_all.groupby("child asin", as_index=False)["total units ordered"].sum()
        unique_asins_sold = unique_asins_sold[unique_asins_sold["total units ordered"] > 0]
        
        #Determine if there are any unique asins with sales that are not in masterlist
        not_in_masterlist = set(unique_asins_sold["child asin"]) - set(masterlist["Child ASIN"])
        print(f"\nThese ASINs have sales but not in masterlist: {not_in_masterlist}")
      
        if not_in_masterlist:
            new_asins = unique_asins_sold[unique_asins_sold["child asin"].isin(not_in_masterlist)]
            save_to = os.path.dirname(WB2_file)
            save_to = os.path.join(save_to, "new_asins.csv")
            new_asins.to_csv(save_to, index=False)
            continue_checking = input(f"\nContinue checking?\n1. Yes\n2. No\nWhat is your choice? ")

            if continue_checking != "1":
                wb2.close()
                app.quit()
                exit()
            elif continue_checking == "1":
                print(f"Continue checking WB4-WB6")

        
        #Open WB4 on temporary copy of the file to avoid autosave in Onedrive
        wb4 = app.books.open(openas_temporary_local(WB4_file))
        wb4.api.Unprotect(Password="bigmoney")

        wb4_months = wb4.sheets['Hist Sales by ASIN'].range("G7:AP7").value
        wb4_sales = wb4.sheets['Hist Sales by ASIN'].range("G209:AP209").value
        wb4_hist_sales = pd.DataFrame({
            "date": wb4_months,
            "total units ordered": wb4_sales})
        print(f"\nDone WB4")
        wb4.close()

        
        #Open WB5 on temporary copy of the file to avoid autosave in Onedrive
        app = xw.App(visible=True, add_book=False)
        wb5 = app.books.open(openas_temporary_local(WB5_file))
        print(f"\nWB5 opened")
        wb5.api.Unprotect(Password="bigmoney")
        print(f"\nWB5 unprotected")

        #skip freeze pane on excel
        window = wb5.app.api.Windows(wb5.name)
        window.SplitRow = 0
        window.SplitColumn = 0
        window.FreezePanes = False
            
        print(f"\nWB5 freezpanes cleared")
        
        #There is difference in range to get if WB5 version is old
        
        if WB5_version == "new":
            month_range = "H7:AQ7"
            value_range = "H209:AQ209"
        elif WB5_version == "old":
            month_range = "G7:AP7"
            value_range = "G209:AP209"

        wb5_months = wb5.sheets['Hist Sales by ASIN'].range(month_range).value
        wb5_sales = wb5.sheets['Hist Sales by ASIN'].range(value_range).value
        wb5_hist_sales = pd.DataFrame({
            "date": wb5_months,
            "total units ordered": wb5_sales})
        print(f"\nDone WB5")
        wb5.close()


        #Open WB6 on temporary copy of the file to avoid autosave in Onedrive
        app = xw.App(visible=False, add_book=False)
        wb6 = app.books.open(openas_temporary_local(WB6_file))
        wb6.api.Unprotect(Password="bigmoney")

        #Check if Hist Sales or Hist Sales by ASIN is the tab in WB6 or not
        sheet_name_old = "Hist Sales"
        sheet_name_new = "Hist Sales by ASIN"
        if sheet_name_old in [sheet.name for sheet in wb6.sheets]:
            sheet_name = sheet_name_old
            month_range = "G5:AP5"
            value_range = "G207:AP207"
        elif sheet_name_new in [sheet.name for sheet in wb6.sheets]:
            sheet_name = sheet_name_new
            month_range = "H7:AQ7"
            value_range = "H209:AQ209"
        
        print(f"\nDetected the sheet details: {sheet_name}, {month_range}, {value_range}")

        wb6_months = wb6.sheets[sheet_name].range(month_range).value
        wb6_sales = wb6.sheets[sheet_name].range(value_range).value
        wb6_hist_sales = pd.DataFrame({
            "date": wb6_months,
            "total units ordered": wb6_sales})
        print(f"\nDone WB6")
        wb6.close()


        wb2_hist_sales = hist_sales_all_grouped.rename(columns={"total units ordered": "wb2 sales"})
        wb4_hist_sales = wb4_hist_sales.rename(columns={"total units ordered": "wb4 sales"})
        wb5_hist_sales = wb5_hist_sales.rename(columns={"total units ordered": "wb5 sales"})
        wb6_hist_sales = wb6_hist_sales.rename(columns={"total units ordered": "wb6 sales"})

        #print data types of the data frames
        '''
        print(f"\nWB2\n{wb2_hist_sales.dtypes}")
        print(f"\nWB4\n{wb4_hist_sales.dtypes}")
        print(f"\nWB5\n{wb5_hist_sales.dtypes}")
        print(f"\nWB6\n{wb6_hist_sales.dtypes}")
        

        #print the data frames
        print(f"\nWB2\n{wb2_hist_sales}")
        print(f"\nWB4\n{wb4_hist_sales}")
        print(f"\nWB5\n{wb5_hist_sales}")
        print(f"\nWB6\n{wb6_hist_sales}")
        '''

        #Merge hist_sales of wb2, wb4, wb5 and wb6
        print(f"\nMerging the dataframes")
        merged_hist_sales = (
            wb2_hist_sales.merge(wb4_hist_sales, on="date", how="outer")
                .merge(wb5_hist_sales, on="date", how="outer")
                .merge(wb6_hist_sales, on="date", how="outer")
        )
        print(f"\nDone Merging")

        #Add difference column for all hist_sales
        print(f"Adding columns for the differences in sales")
        merged_hist_sales["wb2 - wb4"] = merged_hist_sales["wb2 sales"] - merged_hist_sales["wb4 sales"]
        merged_hist_sales["wb2 - wb5"] = merged_hist_sales["wb2 sales"] - merged_hist_sales["wb5 sales"]
        merged_hist_sales["wb2 - wb6"] = merged_hist_sales["wb2 sales"] - merged_hist_sales["wb6 sales"]
        
        #print(f"\n{merged_hist_sales}")

        #Change column data type to int
        merged_hist_sales["wb2 - wb4"] = merged_hist_sales["wb2 - wb4"].round().astype("int")
        merged_hist_sales["wb2 - wb5"] = merged_hist_sales["wb2 - wb5"].round().astype("int")
        merged_hist_sales["wb2 - wb6"] = merged_hist_sales["wb2 - wb6"].round().astype("int")

        '''
        merged_wb2_wb4['difference'] = merged_wb2_wb4["total units ordered_wb2"] - merged_wb2_wb4["total units ordered_wb4"]
        print(merged_wb2_wb4)
        '''
        print(f"\n{merged_hist_sales}")

        #save the result into a csv file
        csv_file_output = os.path.dirname(WB2_file)
        csv_filename = f"hist_sales_checking_{timestamp}.csv"
        csv_file_output = os.path.join(csv_file_output, csv_filename)
        merged_hist_sales.to_csv(csv_file_output, index=False)
        app.quit()
        

    except Exception as e:
        print(f"Failed due to error: {e}")
        wb2.close()
        wb4.close()
        wb5.close()
        wb6.close()
        app.quit()
main()
'''
        #
        #
        #
        #
'''