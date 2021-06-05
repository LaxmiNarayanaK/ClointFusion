import ClointFusion as cf
import os
import shutil

cf.OFF_semi_automatic_mode()
#CREARING DIRECTORY NAME CREATED DOCS IF NOT PRESENT
if not os.path.exists('created docs'):
    os.makedirs('created docs')
#COPYING THE Excel.xlsx FILE FROM ORIGINAL DOCS TO CREATED DOCS
if not os.path.exists('C:/Users/Anu/clointfusion/created docs/Excel.xlsx'):
    shutil.copy('C:/Users/Anu/clointfusion/original docs/Excel.xlsx', 'C:/Users/Anu/clointfusion/created docs/Excel.xlsx')

#1. Get all the header columns.
print(f"header columns are {cf.excel_get_all_header_columns('C:/Users/Anu/clointfusion/created docs/Excel.xlsx')}")


#2. Get Row &amp; Column count.
row_col_count=cf.excel_get_row_column_count('C:/Users/Anu/clointfusion/created docs/Excel.xlsx')
print(f"No. of rows are {row_col_count[0]}")
print(f"No of columns are {row_col_count[1]}")

#3. Get all sheet names in the ‘Excel.xlsx’.

print(f"sheet names are {cf.excel_get_all_sheet_names('C:/Users/Anu/clointfusion/created docs/Excel.xlsx')}")

#4. Remove the duplicate data w.r.t ‘ID’ column.

print(f"No. of unique ids found after removing duplicates is {cf.excel_remove_duplicates('C:/Users/Anu/clointfusion/created docs/Excel.xlsx')}")

#5. Sort the data w.r.t ‘OrderDate’ column.
print(f"Data is sorted wrt Order date {cf.excel_sort_columns('C:/Users/Anu/clointfusion/created docs/Excel.xlsx')}")

# 6. Store the following data in a python dictionary and insert the data at the last
# row respectively.
# a. ID: 1027
# b. OrderDate: 4/14/2020
# c. Region: East
# d. Rep: Jones
# e. Item: Binder
# f. Units: 60
# g. UnitCost: 4.99
# h. Total: 449.1

Dict = {"ID ": 1027, "OrderDate": "4/14/2020", "Region": "East", "Rep": "Jones", "Item": "Binder", "Units": 60,
        "UnitCost": 4.99, "Total": 499.1}
header_columns =cf.excel_get_all_header_columns('C:/Users/Anu/clointfusion/created docs/Excel.xlsx')

for i in range(0, row_col_count[1]):
    cf.excel_set_single_cell(excel_path='C:/Users/Anu/clointfusion/created docs/Excel.xlsx', columnName=header_columns[i], cellNumber=43,setText=Dict[header_columns[i]])

# 7. Split the excel on row count ‘12’. (This will create multiple excel files named
# ‘Split’. Check the function arguments for further customization).

cf.excel_split_the_file_on_row_count(excel_path='C:/Users/Anu/clointfusion/created docs/Excel.xlsx', sheet_name="Split", rowSplitLimit=12,outputFolderPath="C:/Users/Anu/clointfusion/created docs")

#8.Create a python dictionary named ‘data’ such that it stores the ‘ID’ and ‘Units’
# of each row data in the excel file.

data = {1001: 95,
        1002: 50,
        1003: 36,
        1004: 27,
        1005: 56,
        1006: 60,
        1007: 75,
        1008: 90,
        1009: 32,
        1010: 60,
        1011: 90,
        1012: 29,
        1013: 81,
        1014: 35,
        1015: 2,
        1016: 16,
        1017: 28,
        1018: 64,
        1019: 15,
        1020: 96,
        1021: 67,
        1022: 74,
        1023: 46,
        1024: 87,
        1025: 4,
        1026: 7,
        1027: 50,
        1028: 66,
        1029: 96,
        1030: 53,
        1031: 80,
        1032: 5,
        1033: 62,
        1034: 55,
        1035: 42,
        1036: 3,
        1037: 7,
        1038: 76,
        1039: 57,
        1040: 14,
        1041: 11,
        1042: 94,
        1043: 28,
        }