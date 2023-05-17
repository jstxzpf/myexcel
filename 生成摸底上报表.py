import json
import os
import pyodbc
import xlrd
import xlwt

# 读取mysql.json文件中的参数
with open("mysql.json", "r") as f:
    mysql_params = json.load(f)

# 连接到SQL Server数据库
conn = pyodbc.connect("DRIVER={ODBC Driver 18 for SQL Server};SERVER=%s;DATABASE=%s;ENCRYPT=yes;TrustServerCertificate=yes;UID=%s;PWD=%s"
                      % (mysql_params["server"], 
                         mysql_params["database"], mysql_params["user"], mysql_params["password"]))
# 读取泰兴市2023村代码表格
cursor = conn.cursor()

query = "SELECT 乡, 村, 调查村编码 FROM 泰兴市2023村代码"
cursor.execute(query)
results = cursor.fetchall()

# 遍历数据，创建以[乡]字段内容为名的文件夹，并写入数据到xls文件
for row in results:
    xiang = row[0]
    cun = row[1]
    code = row[2]

    # 创建以[乡]字段内容为名的文件夹
    folder_path = '/home/zpf/cqmd/{}'.format(xiang)
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    # 读取[摸底上报表]中符合条件的数据
    query = """
        SELECT [自然村+户编码], [村民小组（自然村）名称], [养殖户名称],
                      [是否法人单位（1或0）], [生猪存栏量（头）], [其中能繁母猪（头）],
                      [生猪出栏量（头）], [牛存栏量（头）], [其中：奶牛存栏（头）],
                      [其中：牦牛存栏（头）], [牛出栏量（头）], [其中：牦牛出栏（头）],
                      [羊存栏量（只）], [其中：绵羊存栏（只）], [羊出栏量（只）],
         [其中：绵羊出栏（只）], [家禽存栏量（只）], [其中蛋禽存栏（只）], [家禽出栏量（只）], [是否代养户（1或0）]
        FROM [摸底上报表]
        WHERE [调查村编码] =\'%s\'
    """%(code)
    print(query)
    
    cursor.execute(query)
    data = cursor.fetchall()
    book = xlrd.open_workbook('/home/zpf/test.xls')
    sheet = book.sheet_by_index(0)
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet=workbook.add_sheet('sheet1')
    for i in range(sheet.nrows):
        for j in range(sheet.ncols):
            worksheet.write(i, j, sheet.cell_value(i, j))
            
    # 从B4单元格开始写入数据
    start_row, start_col = 2, 1
    for i, row in enumerate(data):
        for j, value in enumerate(row):
            worksheet.write(start_row + i, start_col + j, value)

    # 写入数据到xls文件
    file_name = '{}.xls'.format(code[-6:]+cun)
    file_path = os.path.join(folder_path, file_name)
    # 写入“***子表开始”、“单位负责人”和“方丽梅”
    worksheet.write(len(data) + 2, 0, '***子表开始')
    worksheet.write(len(data) + 2, 1, '单位负责人')
    worksheet.write(len(data) + 2, 2, '方丽梅')
    worksheet.write(len(data) + 3, 1, '填表人')
    worksheet.write(len(data) + 3, 2, '翟旭敏')
    worksheet.write(len(data) + 4, 1, '省')
    worksheet.write(len(data) + 4, 2, '江苏省')
    worksheet.write(len(data) + 5, 1, '县')
    worksheet.write(len(data) + 5, 2, '泰兴市')
    worksheet.write(len(data) + 6, 1, '乡')
    worksheet.write(len(data) + 6, 2, xiang)
    worksheet.write(len(data) + 7, 1, '市')
    worksheet.write(len(data) + 7, 2, '泰州市')
    worksheet.write(len(data) + 8, 1, '调查对象地区级别')
    worksheet.write(len(data) + 8, 2, '1')
        
    workbook.save(file_path)

# 关闭数据库连接
cursor.close()
conn.close()