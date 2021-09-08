# ITP Week 3 Day 2 Lecture

#-------------THE OS MODULE----------------
#Review--> Files have a name and a path.
#Review--> The root folder is the lowest folder.
#Review--> In a file path, the folders and filename are separated by backslashes on #Windows and forward slashes on Linux and Mac.
#Review--> The current working directory is the folder that any relative paths are relative to.
#Use the os.path.join() function to combine folders with the correct slash.

#Begin by importing the os module to the top of your file, like so:
import os

from openpyxl.workbook import workbook

os.getcwd() #will return the current working directory.
os.chdir() #will change the current working directory.

#Review--> Absolute paths begin with the root folder, relative paths do not.
#The . folder represents "this folder", the .. folder represents "the parent folder".

os.path.abspath() #returns an absolute path form of the path passed to it.
os.path.relpath() #returns the relative path between two paths passed to it.
os.makedirs() #can make folders.
os.path.getsize() #returns a file's size.
os.listdir() #returns a list of strings of filenames.
os.path.exists() #returns True if the filename passed to it exists.
os.path.isfile() #and 
os.path.isdir() #return True if they were passed a filename or file path.

#
#-----------------BACK TO EXCEL-------------
#

import openpyxl
my_workbook = openpyxl.Workbook()
#Review--> Determine the names of the sheets in the Excel file using the .get_sheet_names() function imported from openpyxl
my_workbook.get_sheet_names() # Result -->  ['Sheet']
#Notice the correllation between lines 37 and 39 ^^^
my_sheet = my_workbook.get_sheet_by_name('Sheet')

#-------INSERTING ROWS & COLUMNS--------

#Much of the following info is pulled directly from the OpenPyXL official documentation

#The following openpyxl methods will allow you to insert & delete rows and columns 
openpyxl.worksheet.worksheet.Worksheet.insert_rows()
openpyxl.worksheet.worksheet.Worksheet.insert_cols()
openpyxl.worksheet.worksheet.Worksheet.delete_rows()
openpyxl.worksheet.worksheet.Worksheet.delete_cols()

#Getting a feel for the values impacted by these methods may take a couple tries :)
#For instance, the default is one row or column. To insert a row at 7 (before the existing row 7):
my_sheet.insert_rows(7)

#Deleting Rows & Columns
my_sheet.delete_cols(6, 3)

#You can also move ranges of cells within a worksheet:
my_sheet.move_range("D4:F10, rows=1, cols=2")
#This will move the cells in the range D4:F10 up one row, and right two columns. The cells will overwrite any existing cells.



#Review--> Access a cell value from the Excel sheet
my_sheet['A1'].value  # Result is 'None' on a black Excel document

#Review--> Change the value of a cell
my_sheet['A1'] = 37
my_sheet['A2'] = 'Pears'

#Because we know that when we create a blank Excel document, the starting value is 'None', we can clear cell values by setting them to 'None' 



#Write Edit Delete Excel
#Write new sheets and a file


#RECAP:
#You can view and modify a sheet's name with its "title" member variable.
#Changing a cell's value is done using the square brackets, just like changing a value in a list or dictionary.
#Changes you make to the workbook object can be saved with the save() method.