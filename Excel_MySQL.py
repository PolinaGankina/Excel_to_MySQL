import xlrd
mf =xlrd.open_workbook('MF.xls')
sheets_name=[sheet.name for sheet in mf.sheets()]

#Getting ordered list of chosen sheets'columns:
Base_SAP=mf.sheet_by_index(2)
columns=[Base_SAP.cell_value(0,col)for col in range(Base_SAP.ncols)]
#for i,item in enumerate(columns):
    #print (i, item)

Costs=mf.sheet_by_index(5)
columns1=[Costs.cell_value(0,col)for col in range(Costs.ncols)]
#for i,item in enumerate(columns1):
    #print (i, item)

# Testing string columns for max length in order to create VARCHAR type:
length_test=[Costs.cell_value(row,12) for row in range (Costs.nrows)]
#var=max([len(i)for i in length_test])
#print (var)

# Testing column raw data for type integrity:
t=[type(i) for i in length_test]
#for i, m  in enumerate(t):
        #print (i,m)

# Testing first raw of each column for general data type:
type_test=[type(Costs.cell_value(1,c)) for c in range(Costs.ncols)]
#for i, item in enumerate(type_test):
    #print (i,item)


#MySQL Connection:
import pymysql
conn = pymysql.connect(host='localhost',user='root',password='root',database='master_file',charset='utf8')
cur = conn.cursor()

cur.execute('USE master_file')

#cur.execute('CREATE TABLE %s (%s INT, %s VARCHAR(40),%s VARCHAR(5),%s INT, %s FLOAT(10,2),%s INT, %s DATE, %s DATE, %s VARCHAR(50),%s VARCHAR(10))'
      #%(sheets_name[5],columns1[0],columns1[4],columns1[5],columns1[6],columns1[7],columns1[8],columns1[2],columns1[12],columns1[13],columns1[1]))


import datetime

#Populating created table with data:
for row in range (1,Costs.nrows):
    values=int(Costs.cell_value(row,0))
    values4=Costs.cell_value(row,4)
    values5=Costs.cell_value(row,5)
    values6=int(Costs.cell_value(row,6))
    values7=Costs.cell_value(row,7)
    values8=int(Costs.cell_value(row,8))
    values13=Costs.cell_value(row,13)
    values1=Costs.cell_value(row,1)
    values14=values13[0:5]

#Changing date format:
    exceltime2=int(Costs.cell_value(row,2))
    time_tuple=xlrd.xldate_as_tuple(exceltime2,0)
    dt2=datetime.datetime(*time_tuple)
    values2=str(dt2)[0:10]

    exceltime12=int(Costs.cell_value(row,12))
    time_tuple1=xlrd.xldate_as_tuple(exceltime12,0)
    dt12=datetime.datetime(*time_tuple1)
    values12=str(dt12)[0:10]


    cur.execute('INSERT INTO %s (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) VALUES (%i,%r,%r,%i,%.2f,%i,%r,%r,%r,%r,%r);'
                %(sheets_name[5],columns1[0],columns1[4],columns1[5],columns1[6],columns1[7],columns1[8],
    columns1[2],columns1[12],'Vendor_Code',columns1[13],columns1[1],values,values4,values5,values6, values7,values8,values2,values12, values14,values13,values1))

#conn.commit()
