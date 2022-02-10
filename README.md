
# AutoXL
AutoXL is a python program to automate excel or spreadsheet to perform a task which earlier consumed considerable time doing manually. But with AutoXL it can be done in a matter of seconds. 
Apart from automating and fixing the xlsx file, a bar chart is also added to display the corrected column as a reference.



## Screenshots

![App Screenshot](https://github.com/BikramdeepSingh/AutoXL/blob/main/images/xlsx%20file%20after%20automation.PNG)


## How code works
* The program takes an excel file name with its extension as the input, along with the name which is to be added for the updating column.
* Built in openpyxl is a python library which is used to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.
* Workbook/Spreadsheet is then loaded into the program and the sheet which is to be updated is passed as an argument to instance of the openpyxl.
* Rows and cells are accessed and required updates are made.
* Built in classes of openpyxl i.e. BarChart and Reference are used to acces the values of updated rows and column.
* Values are then passed to instance of BarChart and is then addded to the sheet.
* Aftermost, Excel file is then saved.

## Authors

- [@bikramdeepsingh](https://github.com/BikramdeepSingh)

