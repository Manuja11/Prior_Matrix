## import library xlwings
import xlwings as xw

## open the main excel workbook (BEST Prioritization Matrix) placed in sharepoint
master_wb = xw.Book(r"C:\Users\u130628\Straumann Group\BEST Internal Team - Miscellaneous Documents\BEST Prioritization Matrix V2.xlsx")

## create a new variable to list sheets within main excel workbook
master_sheets = master_wb.sheets

## create a new list variable to get column A data from row 5 (from sheet 2)
master_sheet_col = master_sheets[1].range('A:A')[4:].value

## open excel workbook with EPIC data downloaded from sharepoint
epicdata_wb = xw.Book(r"C:\Users\u130628\Straumann Group\BEST Internal Team - Miscellaneous Documents\EPIC_Data.xlsx")

## create a new list variable to get data from all three columns from row 2
epicdata_raw = epicdata_wb.sheets[1].range('A2').expand().value

## create another variable to list all values from EPIC data sheet that are not present in main workbook sheet 2 column A
temp_data = [i[:] for i in epicdata_raw if i[0] not in master_sheet_col[:]]

## add new values to main workbook
for i in temp_data:
    new_row = master_sheets[1].range('A5').end('down').row + 1
    master_wb.sheets[1].range(new_row,1).value = i

## save the main workbook
master_wb.save()

## close the main workbook
master_wb.close()

## close the EPIC data workbook
epicdata_wb.close()
