import pymysql
import xlwt

relative_path = './excel2mysql/'

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

def mysql2excel(database='SuperstoreSampleDB',table='test2',excelResult = ''):
    conn = getConn(database)
    cursor = conn.cursor()
    cursor.execute("select * from {}".format(table))
    data_list = cursor.fetchall()
    excel = xlwt.Workbook()
    sheet = excel.add_sheet("sheet1")
    row_number = len(data_list)
    column_number = len(cursor.description)

    print('*'*100)
    print('Got ', row_number, ' data records out of the database.')
    print('Got ', column_number, ' data fields from the database.')
    print('*'*100)

    for i in range(column_number):
        sheet.write(0,i,cursor.description[i][0])
    for i in range(row_number):
        for j in range(column_number):
            sheet.write(i+1,j,data_list[i][j])
        print('*'*100)
        print('write a record: ', data_list[i])

    excelName = "mysql_{}_{}.xls".format(database,table)
    if excelResult != '':
        excelName = excelResult
    excelName = relative_path + excelName
    excel.save(excelName)

if __name__ == "__main__":
    mysql2excel("SuperstoreSampleDB","test2")