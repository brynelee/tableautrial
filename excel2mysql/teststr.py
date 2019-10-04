import xlrd
import pymysql

str1 = 'hello world time'
str2 = str1.replace(' ', '_')
str3 = str1.replace(' ', '-')
print(str2)
print(str3)

excelName = './excel2mysql/test1.xls'
excel = xlrd.open_workbook(excelName)
sheet = excel.sheet_by_index(0)
row_number = sheet.nrows
column_number = sheet.ncols

field_list = sheet.row_values(0)

print(row_number)
print(column_number)
print(field_list)

# define function for field nomalization
def normalizingDBFields(fieldList):
    for i in range(len(fieldList)):
        fieldList[i] = fieldList[i].replace(' ', '_').replace('-', '_').lower()

normalizingDBFields(field_list)
print(field_list)

# processing
data_list = []
for i in range(1,row_number):
    data_list.append(sheet.row_values(i))

table = 'testTableName'
create_sql = "create table {}(".format(table)

for field in field_list[:-1]:
    create_sql += "{} varchar(50),".format(field)
create_sql += "{} varchar(50))".format(field_list[-1])

print('='*50)
print('table create SQL is: ', create_sql)
print('='*50)

for data in data_list:
    new_data = ["'{}'".format(pymysql.escape_string(str(i))) for i in data]
    insert_sql = "insert into {} values({})".format(\
        table,','.join(new_data))
    print('data insert SQL are: ', insert_sql)

