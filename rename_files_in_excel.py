"""
第一步获取目录下所有的文件名
第二部创建一个excel写入所有的文件名
第三步提醒用户编辑生成的excel，重命名文件
第四步按照用户编辑后的excel重命名文件
"""


import os
import xlwt
import xlrd
from xlwt import Workbook

p = os.listdir()

"""
创建excel
"""
w = Workbook(encoding='utf-8')

ws = w.add_sheet('1')
ws.write(0, 0, "源文件名称")
ws.write(0, 1, "修改后的文件名称")
first_col = ws.col(0)
first_col.width = 10000
second_col = ws.col(1)
second_col.width = 10000
j = 1
for i in range(0, len(p)):
    if p[i] == 'rename_files_in_excel.py':
        continue
    ws.write(j, 0, p[i])
    j = j + 1

w.save('重命名.xls')
input("请填充本目录下生成的【重命名.xls】文件中的修改后文件名称列表,完成后，请按回车键")

book = xlrd.open_workbook('重命名.xls')
sheet1 = book.sheet_by_index(0)


for i in range(1, len(p)):
    src = str(sheet1.cell(i, 0).value)
    dst = str(sheet1.cell(i, 1).value)
    src = str(src)
    dst = str(dst)
    print("将"+src+"重命名为" + dst)
choose1 = input("确认重命名？确认请按回车键，取消请关闭窗口")

for i in range(1, len(p)):
    src = str(sheet1.cell(i, 0).value)
    dst = str(sheet1.cell(i, 1).value)
    src = str(src)
    dst = str(dst)
    os.rename(src, dst)

os.remove('重命名.xls')






