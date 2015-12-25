import xlrd
xlf =xlrd.open_workbook('MF.xls')

#Return list of sheets in Excel file:
def Sheets_Name(x):
    sheets_name=[sheet.name for sheet in x.sheets()]
    print (sheets_name)

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


#Establishing MySQL Connection:
import pymysql

conn = pymysql.connect(host='localhost',user='root',password='root',database='master_file',charset='utf8')
cur = conn.cursor()

#cur.execute('USE master_file')

#cur.execute('CREATE TABLE %s (%s INT, %s VARCHAR(40),%s VARCHAR(5),%s INT, %s FLOAT(10,2),%s INT, %s DATE, %s DATE, %s VARCHAR(50),%s VARCHAR(10))'
      #%(sheets_name[5],columns1[0],columns1[4],columns1[5],columns1[6],columns1[7],columns1[8],columns1[2],columns1[12],columns1[13],columns1[1]))




#Populating created table with data:
import datetime
Sheet=xlf.sheet_by_index()

"""for row in range (1,Sheet.nrows):
    values=int(Sheet.cell_value(row,0))
    values4=Sheet.cell_value(row,4)
    values5=Sheet.cell_value(row,5)
    values6=int(Sheet.cell_value(row,6))
    values7=Sheet.cell_value(row,7)
    values8=int(Sheet.cell_value(row,8))
    values13=Sheet.cell_value(row,13)
    values1=Sheet.cell_value(row,1)
    values14=values13[0:5]

#Changing date format:
    exceltime2=int(Sheet.cell_value(row,2))
    time_tuple=xlrd.xldate_as_tuple(exceltime2,0)
    dt2=datetime.datetime(*time_tuple)
    values2=str(dt2)[0:10]

    exceltime12=int(Sheet.cell_value(row,12))
    time_tuple1=xlrd.xldate_as_tuple(exceltime12,0)
    dt12=datetime.datetime(*time_tuple1)
    values12=str(dt12)[0:10]


    #cur.execute('INSERT INTO %s (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) VALUES (%i,%r,%r,%i,%.2f,%i,%r,%r,%r,%r,%r);'
                #%(sheets_name[5],columns1[0],columns1[4],columns1[5],columns1[6],columns1[7],columns1[8],
    #columns1[2],columns1[12],'Vendor_Code',columns1[13],columns1[1],values,values4,values5,values6, values7,values8,values2,values12, values14,values13,values1))"""

#conn.commit()
