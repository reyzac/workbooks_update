import os.path
import select
from sqlite3 import TimestampFromTicks
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
import time
import ctypes

def normalize_columns(df):
    df = df.rename(columns=lambda col: re.sub(r'[^a-zA-Z0-9 ]', '', col).strip().lower())
    df = df.rename(columns=lambda col: re.sub(r' {2,}', ' ', col))
    return df

def remove_currency(df, column_name1, column_name2):
    df[column_name1] = df[column_name1].str.replace(r'[^\d\.]', '', regex=True)
    df[column_name1] = pd.to_numeric(df[column_name1], errors='coerce')
    
    df[column_name2] = df[column_name2].str.replace(r'[^\d\.]', '', regex=True)
    df[column_name2] = pd.to_numeric(df[column_name2], errors='coerce')
    return df

def select_folder():
    #Force focus into python window
    ctypes.windll.user32.SetForegroundWindow(ctypes.windll.kernel32.GetConsoleWindow())

    root = tk.Tk() #This creates a root window object—essentially the main GUI window.Even though you won’t display it, it’s required to initialize the GUI environment.
    root.withdraw() #This hides the root window so it doesn’t pop up awkwardly.
    root.attributes('-topmost', True)  # Force dialog to front
    folder_path = filedialog.askdirectory(parent=root)
    return folder_path

def select_file(title):
    #Force focus into python window
    ctypes.windll.user32.SetForegroundWindow(ctypes.windll.kernel32.GetConsoleWindow())

    root = tk.Tk() #This creates a root window object—essentially the main GUI window.Even though you won’t display it, it’s required to initialize the GUI environment.
    root.withdraw() #This hides the root window so it doesn’t pop up awkwardly.
    root.attributes('-topmost', True)  # Force dialog to front
    file_path = filedialog.askopenfilename(filetypes=[("Excel Binary Workbook", "*.xlsb")], title=title, parent=root)
    return file_path

def list_all_sheets(wb, message):
    print("Available sheets: ")
    for i, sheet in enumerate(wb.sheets):
        print(f"{i}: {sheet.name}")
    print(f"\n{message} \n>")
    sheet_index = int(input("\nEnter the # of sheet to select: "))
    return sheet_index

def generate_new_filename(filePath, timestamp):
    new_filename = os.path.basename(filePath)
    new_filename = new_filename.replace('.xlsb', timestamp + '.xlsb')
    filefolder = os.path.dirname(filePath)
    output_folder = os.path.join(filefolder, "output")
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"\nCreating output folder")
    else:
        print(f"\nSaving file, please wait.....")

    new_filename = os.path.join(output_folder, new_filename)
    return new_filename

def open_astemporary_local(original_path: str) -> Path:
    original = Path(original_path)
    temp_dir = Path(tempfile.gettempdir())
    temp_copy = temp_dir / original.name
    shutil.copy2(original, temp_copy)
    return temp_copy

def select_workbooks():
    #print(f"\n Select WB2 file")
    WB2_file = select_file("Select WB2 file that is UPDATED or from OUTPUT folder!!")

    #print(f"\n Select WB4 file")
    WB4_file = select_file("Select WB4 file")

    #print(f"\n Select WB5 file")
    WB5_file = select_file("Select WB5 file")

    #print(f"\n Select WB6 file")
    WB6_file = select_file("Select WB6 file")
    return WB2_file, WB4_file, WB5_file, WB6_file

def main():
    timestamp = datetime.datetime.now().strftime("_%H%M%S_%m.%d.%Y")

    what_to_do = input(f"\nSelect what to do:\n1. Update WB2\n2. Just Rollover WB2\nWhat is your choice? ")

    wb2_monthly_update = "2025_09_US"
    wb2_date = "09/25/2025"
    wb2_export = "Export to Customer Forecasting"
    wb4_import = "Import Product Sales (WB2)"
    decision = input(f"\nSelect workbook:\n4. WB4\n5. WB5\nWhat workbook to paste in WB6? : ")
    WB4_file = ""

    if decision == "5":
        wb4_export = "Export to SV (WB5)"
        wb5_import = "Everything for Import"
        wb5_export = "Export to Financial Package"
        wb6_import = "Import from Valuation Workbook "
    elif decision == "4":
        wb4_export = "Export Values to WB5&6"
        wb5_import = "Import Values from WB4"
        wb6_import = "Import Values from WB4"

    if what_to_do == "1":
        print(f"\nSelect the folder where the csv files are saved")
        time.sleep(0.5)
        csv_folder = select_folder()
        csv_files = glob.glob(os.path.join(csv_folder, "*.csv"))
        #print(csv_files)
        #input("Press any key to continue")
        df_csv_files = [] #Initialize an emtpy list to store the dataFrames for each csv files

        for f in csv_files:
            df = pd.read_csv(f, encoding="utf-8")
            print("\n" + f + "\n")
            country = input("Input what amazon seller central country: ")

            #always forgot US is .com and not .us, so this is the solution
            if country == ".us":
                country = ".com"

            df['channel'] = country
            print(df.head())
            df_csv_files.append(df)
            #print(f)
            #print(df.dtypes)

        #print(type(df_csv_files))

        df_csv_files = [normalize_columns(df) for df in df_csv_files]

    
        #print(df_csv_files)
    
        combined = pd.concat(df_csv_files, ignore_index=True)
        remove_currency(combined, "ordered product sales", "ordered product sales b2b")
    
        column_names = ["parent asin","child asin","title","sessions total","session percentage total","page views total","page views percentage total","featured offer buy box percentage","units ordered","units ordered b2b","unit session percentage","unit session percentage b2b","ordered product sales","ordered product sales b2b","total order items","total order items b2b"]
    
        #remove columns that are not in the list
        combined_clean = combined[column_names]
        combined_channel = pd.DataFrame()
        combined_channel['channel'] = combined['channel']
        print(combined_channel.head())

        #Output a csv file for the combined
        combined_clean_filename = f"hist_sales_{wb2_monthly_update}.csv"
        combined_file_path = os.path.join(csv_folder, "output", combined_clean_filename)
        combined_clean.to_csv(combined_file_path, index=False)

        
    
        #Call the select workbooks function
        WB2_file, WB4_file, WB5_file, WB6_file = select_workbooks()
        WB2_folder = os.path.dirname(WB2_file)

        print(WB2_file)

        if os.path.exists(WB2_file):
            print("file exists:", WB2_file)
        else:
            print("file not found")

        try:
            app = xw.App(visible=False) 
            wb2 = app.books.open(open_astemporary_local(WB2_file))
            print("\n Workbook opened successfully!!!! \n")
        except Exception as e:
            print("\nFailed to open workbook: \n", e)

        try:
            message = "IN WB2: what sheet to paste the monthly update? "
            print(f"{message} {wb2_monthly_update}\n>")
            #target_sheet = wb2.sheets[list_all_sheets(wb2, message)]
            target_sheet = wb2.sheets[wb2_monthly_update]
            target_sheet.range('A2').value = combined_clean.values.tolist()
            target_sheet.range('AB2').value = combined_channel.values.tolist()
            #wb2.sheets['Channel Setup'].range('K17').value = input("Enter date (ex. 01/25/2025): ")
            wb2.sheets['Channel Setup'].range('K17').value = wb2_date
        except Exception as e:
            print("Sheet or range access failed:", e) 
        
        #Generate a new filename for WB2 with timestamp at the end 
        WB2_file_new = generate_new_filename(WB2_file, timestamp)

        #Save as new file updated WB2
        try:
            wb2.save(WB2_file_new)
            print(f"\n WB2 file saved in {WB2_file_new}")
            wb2.close()
        except Exception as e:
            print(f"\n Unable to save WB2 file: {e}")

        what_to_do = input(f"\nProceed to roll over?\n1. No\n2. Yes\nWhat is your choice? ")

    if what_to_do == "2":
        #Call the select workbooks function
        if WB4_file:
            print(f"\nWorkbook files already selected")
        else:
            WB2_file, WB4_file, WB5_file, WB6_file = select_workbooks()
            WB2_folder = os.path.dirname(WB2_file)

        try:
            app = xw.App(visible=False) 
            wb2 = app.books.open(open_astemporary_local(WB2_file))
            print(f"\nProceeding to roll over to WB4-6")
            message = "IN WB2: what sheet is for export? "
            print(f"\n{message} {wb2_export}\n>")
            #source_sheet = wb2.sheets[list_all_sheets(wb2, message)]
            source_sheet = wb2.sheets[wb2_export]

        except Exception as e:
            print(f"Error on latest update: {e}")
        try:
            
            wb4 = app.books.open(open_astemporary_local(WB4_file))
        except pywintypes.com_error as e:
            print(f"\nWARNING: ERROR OCCURED but continue: {e}")

        try:
            message = "IN WB4: what sheet is import?"
            print(f"\n{message} {wb4_import}")
            #target_sheet = wb4.sheets[list_all_sheets(wb4, message)]
            target_sheet = wb4.sheets[wb4_import]
            target_sheet.range('A1').value = source_sheet.used_range.value
            wb2.close()
        except Exception as e:
            print("\nWB4 sheet or range access failed: ", e)

        

        

        #Generate a new filename for WB4 with timestamp at the end 
        WB4_file_new = generate_new_filename(WB4_file, timestamp)

        #Save as new file updated WB4
        try:
            wb4.save(WB4_file_new)
            print(f"\n WB4 file saved in {WB4_file_new}")
        except Exception as e:
            print(f"\n Unable to save WB4 file: {e}")

        

        #Open WB5 file
        try:
            if decision == "5":
                app = xw.App(visible=True, add_book=False)
                wb5 = app.books.open(open_astemporary_local(WB5_file))
                print(f"\nWB5 opened")
                wb5.api.Unprotect(Password="bigmoney")
                print(f"\nWB5 unprotected")

                #skip freeze pane on excel
                window = wb5.app.api.Windows(wb5.name)
                window.SplitRow = 0
                window.SplitColumn = 0
                window.FreezePanes = False
            
                print(f"\nWB5 freezpanes cleared")
            else:
                wb5 = app.books.open(open_astemporary_local(WB5_file))
        except Exception as e:
            print(f"\n Unable to open WB5 file: {e}")

        #Copy WB4 export sheet to WB5/WB6
        message = "IN WB4: what sheet is for export? "
        print(f"\n{message} {wb4_export}\n>")
        #source_sheet = wb4.sheets[list_all_sheets(wb4, message)]
        source_sheet = wb4.sheets[wb4_export]

        #Paste values of WB4 export to WB5  
        try:
            message = "IN WB5: what sheet is import?"
            print(f"\n{message} {wb5_import}")
            target_sheet = wb5.sheets[list_all_sheets(wb5, message)]
            target_sheet.range('A1').value = source_sheet.used_range.value
        except Exception as e:
            print(f"\nWB5 sheet or range access failed: {e}\n")
            wb5.close()
            app.quit()
            print(f"\nClosing workbooks...Exiting app..")
            exit()

    
        #Save as new file updated WB5
        try:
            WB5_file_new = generate_new_filename(WB5_file, timestamp)
            wb5.save(WB5_file_new)
            print(f"\n WB5 file saved in {WB5_file_new}")
        except Exception as e:
            print(f"\n Unable to save WB5 file: {e}")
    
        
        if decision == '5':
            wb4.close()
            message = "IN WB5: what sheet is for export? "
            print(f"\n{message} {wb5_export}\n>")
            #source_sheet = wb5.sheets[list_all_sheets(wb5, message)]
            source_sheet = wb5.sheets[wb5_export]
        elif decision == '4':
            wb5.close()
        else:
            print(f"\nWrong choice, default is that WB4 will be pasted to WB6")

        #Open WB6 file
        try:
            
            wb6 = app.books.open(open_astemporary_local(WB6_file))
            if decision == "5":
                app = xw.App(visible=True, add_book=False)
                wb5 = app.books.open(open_astemporary_local(WB5_file))
                print(f"\nWB5 opened")
                wb5.api.Unprotect(Password="bigmoney")
                print(f"\nWB5 unprotected")

                #skip freeze pane on excel
                window = wb5.app.api.Windows(wb5.name)
                window.SplitRow = 0
                window.SplitColumn = 0
                window.FreezePanes = False
            
                print(f"\nWB5 freezpanes cleared")
        except Exception as e:
            print(f"\n Unable to open WB6 file: {e}\n")
            wb6.close()
            app.quit()
            print(f"\nClosing workbooks...Exiting app..")
            exit()

        #Paste values of WB4 export to WB6
        try:
            message = "IN WB6: what sheet is import?"
            print(f"\n{message} {wb6_import}")
            target_sheet = wb6.sheets[list_all_sheets(wb6, message)]
            target_sheet.range('A1').value = source_sheet.used_range.value
        except Exception as e:
            print(f"\nWB6 sheet or range access failed: {e}\n")
            wb6.close()
            app.quit()
            print(f"\nClosing workbooks...Exiting app..")
            exit()

        #Save as new file updated WB6
        try:
            WB6_file_new = generate_new_filename(WB6_file, timestamp)
            wb6.save(WB6_file_new)
            print(f"\n WB6 file saved in {WB6_file_new}")
            wb6.close()
        except Exception as e:
            print(f"\n Unable to save WB6 file: {e}")

        if decision == '4':
            wb4.close()
        elif decision == '5':
            wb5.close()
    
        print(f"\nROLLOVER Done. You are so very very very brightest of the world, hello world!\n")
        app.quit()

    elif what_to_do == "1":
        print(f"\nWB2 ONLY update DONE.\n")
        wb2.close()
        app.quit()
    
    else: 
        print(f"invalid choice")

    

main()