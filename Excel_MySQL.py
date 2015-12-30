import xlrd
xlf =xlrd.open_workbook('MF.xls')

#Return list of sheets in Excel file:
def Sheets_Name(x):
    sheets_name=[sheet.name for sheet in x.sheets()]
    for i,item in enumerate(sheets_name):
        print (i,item)

#Getting ordered list of chosen sheets'columns:
def Sheet_Columns(n):
    Sheet=xlf.sheet_by_index(n)
    columns=[Sheet.cell_value(0,col)for col in range(Sheet.ncols)]
    for i,item in enumerate(columns):
        print (i,item)

# Testing string columns for max length in order to create VARCHAR type:
def Column_length(n,col):
    Sheet=xlf.sheet_by_index(n)
    length_test=[Sheet.cell_value(row,col) for row in range (Sheet.nrows)]
    var=max([len(i)for i in length_test])
    print (var)

# Testing column raw data for type integrity:
def Column_Type(n,col):
    Sheet=xlf.sheet_by_index(n)
    length_test=[Sheet.cell_value(row,col) for row in range (Sheet.nrows)]
    t=[type(i) for i in length_test]
    for i, item  in enumerate(t):
        print (i,item)

# Testing first raw of each column for general data type:
def General_Type(n,row):
    Sheet=xlf.sheet_by_index(n)
    type_test=[type(Sheet.cell_value(row,c)) for c in range(Sheet.ncols)]
    for i, item in enumerate(type_test):
        print (i,item)

#Sheets_Name()
#Sheet_Columns()
#Column_length()
#General_Type()
#Column_Type()

#Establishing MySQL Connection:
import pymysql

conn = pymysql.connect(host='localhost',user='root',password='root',database='master_file',charset='utf8')
cur = conn.cursor()

sheets_name=[sheet.name for sheet in xlf.sheets()]
Sheet=xlf.sheet_by_index(7)
columns=[Sheet.cell_value(0,col)for col in range(Sheet.ncols)]

cur.execute('USE master_file')

#Creating MYSQL Table:

#cur.execute('CREATE TABLE %s (%s INT, %s VARCHAR(4),%s INT, %s DATE, %s INT)'
#%(sheets_name[7],columns[0],columns[1],columns[2],columns[3],columns[4]))


#Populating created table with data:
import datetime

for row in range (1,Sheet.nrows):
    values=int(Sheet.cell_value(row,0))
    values1=Sheet.cell_value(row,1)
    values2=int(Sheet.cell_value(row,2))
    values4=int(Sheet.cell_value(row,4))

#Changing date format:
    exceltime=int(Sheet.cell_value(row,3))
    time_tuple=xlrd.xldate_as_tuple(exceltime,0)
    dt=datetime.datetime(*time_tuple)
    values3=str(dt)[0:10]

    cur.execute('INSERT INTO %s (%s,%s,%s,%s,%s) VALUES (%i,%r,%i,%r,%i);'
                %(sheets_name[7],columns[0],columns[1],columns[2],columns[3],columns[4],values,values1,values2,values3,values4))

conn.commit()

cur.close()
conn.close()
