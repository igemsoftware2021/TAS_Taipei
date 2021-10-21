Installations: 
    Install Python 3
    Install OpenPyXL (pip install openpyxl) in any Python environment

Excel Files: 
    Only the first sheet in each Excel file is of interest "daily"
    Scroll to the far right on "daily" to see the supply-demand for each week for each blood type

Code: 
    "OneYearExcelSheet.py", reads and edits and saves one Excel file once
        Change the year in the filename at the TOP AND BOTTOM of to change the Excel file editted
    "10000Simulations.py", reads one Excel file 16 * 10,000 times and DOES NOT edit/save
        Change the year in the filename at the TOP AND BOTTOM of to change the Excel file used