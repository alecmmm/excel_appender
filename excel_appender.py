# -*- coding: utf-8 -*-
"""
Spyder Editor

Written by Alec McKay, September, 2019

"""

from tkinter import filedialog
from tkinter import Tk
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

#creates a file, appending an integer to its name if a file with that name
#is already open
def create_file(directory, file_name, df):
    i = 1
    
    output_name = file_name
    
    while True:
        
        try:
            df.to_excel(directory + output_name + ".xlsx")
            return output_name
            
        except PermissionError:
            i+= 1
            output_name = file_name + str(i)

#appends files together
def append_books(filenames):
    #loop through all files in directory. If file is of .xlsx type, doesn't begin with '~' 
    #and isn't 'appended_book.xlsx', then transform into a dataframe and append onto the 
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
    
    output_name = create_file(directory, "\\appended_book", appendBook)
    
    #open aggregated file and bring to focus
    appendedBook = xw.Book(directory + output_name + '.xlsx')
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
