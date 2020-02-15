# excel_appender
Combines tables in Excel workbooks and CSVs

How to Use:
    Run program. When prompted, select multiple files 
    with tables that you wish to append together. The 
    tables will be appended onto each other in one 
    workbook named appendedBook.xlsx, in the same 
    folder as the files you appended.

Requirements:
    1. Data to be appended must be in the first worksheet
    2. All data must have the same headers
    3. All workbooks must be in the same folder.
    4. Workbook named appendedBook.xlsx must not be open
    5. If there is a workbook names appendedBook.xlsx in 
       the folder you're running it, it will be 
       overridden
    6. Must have TkInter, pandas and xlwings
