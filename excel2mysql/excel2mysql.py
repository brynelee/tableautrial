import xlrd
import pymysql

def getConn(database='SuperstoreSampleDB'):
    args = dict(
        host='localhost',
        user='root',
        passwd='time4@FUN',
        db=database,
        charset='utf8'
    )
    conn = pymysql.connect(**args)
    return conn


# define function for field nomalization
def normalizingDBFields(fieldList):
    for i in range(len(fieldList)):
        fieldList[i] = fieldList[i].replace(' ', '_').replace('-', '_').lower()
   

def excel2mysql(excelName,database='SuperstoreSampleDB',table='test2'):
    #下面代码作用：获取到excel中的字段和数据
    excel = xlrd.open_workbook(excelName)
    sheet = excel.sheet_by_index(0)
    row_number = sheet.nrows
    column_number = sheet.ncols

    # prepare field list
    field_list = sheet.row_values(0)
    normalizingDBFields(field_list)

    data_list = []
    for i in range(1,row_number):
        data_list.append(sheet.row_values(i))

    #下面代码作用：根据字段创建表，根据数据执行插入语句
    conn = getConn(database)
    cursor = conn.cursor()
    drop_sql = "drop table if exists {}".format(table)
    cursor.execute(drop_sql)
    create_sql = "create table {}(".format(table)

    for field in field_list[:-1]:
        create_sql += "{} varchar(200),".format(field.replace(' ', '_'))
    create_sql += "{} varchar(200))".format(field_list[-1].replace(' ', '_'))
    print('='*100)
    print(create_sql)
    print('='*100)
    cursor.execute(create_sql)

    for data in data_list:
        new_data = ["'{}'".format(pymysql.escape_string(str(i))) for i in data]
        insert_sql = "insert into {} values({})".format(\
            table,','.join(new_data))
        print('*'*100)
        print(insert_sql)
        print('*'*100)
        cursor.execute(insert_sql)

    conn.commit()
    conn.close()

if __name__ == '__main__':
    excel2mysql("./excel2mysql/test1.xls")