# sangoma_compare.py
# Highlight all assemblies we are testing
import openpyxl
import re                                                                                     # Regular Expressions library
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

# Settings and variables
wb1_filepath = 'Mara Transfer List_MCO Update.xlsx'
wb2_filepath = 'Master Test Matrix_CCD_2019.xlsx'

wb1_ws = {"Transfer list" : 2}
wb2_ws = {"ProdData_TW_SW_MTP": 1, "Test Categories Mapping": 1}

# Flag to highlight all matches even after finding it the first instance
find_all_match = True

def open_files(wb1, wb2):
    '''
    Opens workbook1 and workbook2 file paths and assigns to workbook1 and workbook2
    Select sheets and assigns to workbook1_ws and workbook2_ws
    '''
    wb1_path = wb1
    wb2_path = wb2

    workbook1 = openpyxl.load_workbook(wb1_path)
    workbook2 = openpyxl.load_workbook(wb2_path)

    #workbook1_ws = workbook1[ws1]                  # Opens the worksheet of workbook1
    #workbook2_ws = workbook2[ws2]                  # Opens the worksheet of workbook2

    return workbook1, workbook2

def test(org_sheet, new_sheet):
    print(org_sheet['A3'].value, new_sheet['B5'].value)

def stats(ws):
    '''
    Pass in the worksheet and will return the max column and row of sheet.
    '''
    max_row = ws.max_row
    max_col = ws.max_column

    print("There are " + str(max_row) + " rows and " + str(max_col) + " columns in " + str(ws) + ".") 

def compare(ws1, ws1_column, ws2, ws2_column, find_all_match):
    '''
    Takes in worksheet 1, worksheet 1 column and compares against workseet 2, with worksheet 2 column.
    Worksheet 2 will be highlighted with the matches in purple and saved as a new file.
    Can change to highlight worksheet 1, just need to comment out the line.
    '''

    # ws1 or worksheet1 max rows and columns
    max_ws1_row = ws1.max_row
    max_ws1_col = ws1.max_column

    # ws2 or worksheet2 max rows and columns
    max_ws2_row = ws2.max_row
    max_ws2_col = ws2.max_column
    
    not_found = False                                                # Flag for finding value

    for r in range(1, max_ws1_row + 1):                              # Loop through rows, start at row 1, and max row + 1 to include last row
        ws1_value = ws1.cell(r, ws1_column).value

        for ws2_r in range(1, max_ws2_row + 1):                     # max row + 1 to include the last line
            if ws1_value == ws2.cell(ws2_r, ws2_column).value:
                #ws1.cell(r, ws1_column).fill = PatternFill(patternType='solid', fill_type='solid', fgColor='ca19ec')   # Highlight cell to purple
                ws2.cell(ws2_r, ws2_column).fill = PatternFill(patternType='solid', fill_type='solid', fgColor='ca19ec') # Highlight cell to purple
                not_found = False
                #print(ws1_value)                                # Prints out found value for debugging
                if not find_all_match:                           # If user wants to find all matches even after finding the first instance
                    break                                        # Found value in wb2 sheet, break out, don't break out to highlight duplicates as well
            elif ws1_value == None:
                not_found = False
                break
            else:
                not_found = True
                                   
        if not_found:
            print(str(ws1_value) + " not found.")

if __name__ == '__main__':
    wb1, wb2 = open_files(wb1_filepath, wb2_filepath)

    for wb1_ws_key in wb1_ws:                                                                       # Loop through each worksheet in workbook 1 (usually just 1)
        print(str(wb1_ws_key))
        for wb2_ws_key in wb2_ws:                                                                   # Loop through each worksheet in workbook 2
            '''
            Feeds in workseet 1 key (worksheet title) and worksheet 1 value (worksheet column) and
            feeds in workseet 2 key (worksheet title) and worksheet 2 value (worksheet column) into 
            compare function so it can compare worksheet 1 with multiple workseets of workbook 2.
            '''
            compare(wb1[wb1_ws_key], wb1_ws[wb1_ws_key], wb2[wb2_ws_key], wb2_ws[wb2_ws_key], find_all_match)      
    #test(ws1, ws2)
    #print(ws1['B4'].value)
    #stats(ws1)
    #stats(ws2)
    #wb1.save('wb1.xlsx')
    
    wb2.save('new.xlsx')