# -*- coding: utf-8 -*-
"""
Spyder Editor

Written by Alec McKay, September, 2019

Script for combining tables in Excel or csv workbooks.

How to Use:
    Run program. When prompted, select multiple files 
    with tables that you wish to append together. The 
    tables will be appended onto eachother in one 
    workbook named appendedBook.xlsx, in the same 
    folder as the files you appended. A rightmost
    column will be appended with the name of the 
    file that the row was taken from. A worksheet 
    "Directory" will also be created that lists the 
    names of all files that were appended

Requirements:
    1. Data to be appended must be in the first worksheet
    2. All data must have the same headers
    3. All workbooks must be in the same folder.
    4. Workbook named appendedBook.xlsx must not be open
    5. If there is a workbook names appendedBook.xlsx in 
       the folder you're running it, it will be 
       overridden
"""

from os import listdir
from tkinter import filedialog
from tkinter import Tk
#from tkinter import *
from tkinter import messagebox
import pandas as pd
import xlwings as xw
import sys


#interface for opening files
def open_files():
    
    #initiate UI
    root = Tk()   
    root.lift()
    
    root.filename = filedialog.askopenfilenames(initialdir="/", title="Select file")
    root.destroy()
    
    if root.filename == '':
        sys.exit()
    
    return root.filename

#interface for displaying message
def display_message(messageTitle, message):
    
    #initiate UI
    root = Tk()
    
    #hide root window
    root.withdraw()
    
    #create message box to display message
    messagebox.showinfo(messageTitle, message)
    root.destroy()

#appends files together
def append_books(filenames):
       
    #loop through all files in directory. If file is of .xlsx type, doesn't begin with '~' 
    #and isn't 'appendedBook.xlsx', then transform into a dataframe and append onto the 
    #empty dataframe that was created.
    appendBook = pd.DataFrame()
        
    directory = filenames[0][0: filenames[0].rfind('/')]
    
    file_short_names = []
    
    for file in filenames:
        
        if(file[-3:len(file)] == 'xls' or file[-4:len(file)] == 'xlsx'):
            df = pd.read_excel(file,sheet_name=0, index=False, index_label=False)
            
        elif(file[-3:len(file)] == 'csv'):
            df = pd.read_csv(file)
            
        else:
            display_message("Error", "only select files of the type xlsx, xls or csv.")
            sys.exit("Wrong type of file appended")
        
        file_short_name = file[file.rfind('/') + 1:]
        file_short_names.append(file_short_name)
        df['source_name'] = [file_short_name] * len(df.index)
        appendBook = appendBook.append(df)
        
    try:
        #write dataframe into file
        appendBook.to_excel(directory + '\\appendedBook.xlsx')

    except PermissionError:
        display_message("Error", "Cannot have workbook called appendedBook.xlsx open while running macro. Please close and try again")
        sys.exit("appendedBook.xlsx was open")
        
    #open aggregated file and bring to focus
    appendedBook = xw.Book(directory + '\\appendedBook.xlsx')
    
    directory = appendedBook.sheets.add("directory")
    
    i = 1
    
    for name in file_short_names:
        directory.range('A' + str(i)).value = name
        i += 1
   
    appendedBook.sheets[1].activate()
    xw.apps.active.activate(steal_focus=True)
   
def main():
    append_books(open_files())
    
if __name__ == '__main__':
    main()