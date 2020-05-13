#User Defined Exception
class Object_Failed_Exception(Exception):
    def _init_(self, message):
        self.message = message


#Checking Status
def Status_Checking(Rows,Columns):
    flag = True
    while (sheet.cell_value(Rows+1,Columns+1) == 'Fail' or sheet.cell_value(Rows+1,Columns+1) == 'Pass'):
        status = sheet.cell_value(Rows+1,Columns+1)
        if status == 'Fail':
            list.append(sheet.cell_value(Rows+1,Columns))
            flag = False
        Rows +=1
    return flag


#Raising Exception
def Raising_Exception(flag):
    try:
        if flag == False:
            raise Object_Failed_Exception(list)
    finally:
        print("Finished")   


#Reading File
import xlrd
import glob
import os
import pandas as pd 
list_of_files = glob.glob(os.path.join(r'C:\Users\Kunj\Desktop\Final\*.csv'))
latest_file = max(list_of_files, key=os.path.getmtime)
print(latest_file)
df = pd.read_csv(latest_file, sep='\t', header=None)
writer = pd.ExcelWriter('test.xlsx')
df.to_excel(writer, index=False)
writer.save()
list_of_files = glob.glob(os.path.join(r'C:\Users\Kunj\Desktop\Final\*.xlsx'))
latest_file = max(list_of_files, key=os.path.getmtime)
print(latest_file)
workbook = xlrd.open_workbook(latest_file)
sheet = workbook.sheet_by_index(0)
list = []


#Scanning the Sheet
for rows in range(sheet.nrows):
    for columns in range(sheet.ncols):
        names = sheet.cell_value(rows, columns)
        if names == 'Project Name':
            rProject = rows
            cProject = columns

        elif names == 'Object ':
            rObject = rows
            cObject = columns


#Calling Function to check the Status for Project and Object
flag_1 = Status_Checking(rProject,cProject)
flag_2 = Status_Checking(rObject,cObject)


#Removing Excel File
os.remove('test.xlsx')


#Calling Function to check whether exception can be raise or not
Raising_Exception(flag_1)
Raising_Exception(flag_2)




                
