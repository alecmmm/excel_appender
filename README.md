# excel_appender
Combines tables in Excel workbooks and/or csv files

## How to Use
Best run from Anaconda environment. Run program. When prompted, select multiple files that you wish to append together (in Tkinter interface, you can select multiple files by holding CTRL and left clicking files). The tables will be appended onto each other in oneworkbook named appendedBook.xlsx, in the same folder as the files you appended. A rightmost column will be added, named "source_name". A worksheet named "directory" will be created containing the names of all files that were appended.

## Pictures

## Requirements
1. When appending Excel files, data must be in the first worksheet of the workbook
2. All data must have the same headers
3. All files must be in the same folder.
4. To avoid name conflicts, no workbook named appendedBook.xlsx can be open
5. If there is a workbook already named appendedBook.xlsx in 
   the folder you're running it, it will be 
   overridden
6. Must have TkInter, pandas and xlwings packages
