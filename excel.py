#!user/bin/env python3

path=input('输入EXCEL路径:')

resultname=input('输入形成的文件路径:')

import xlrd

source=xlrd.open_workbook(path)

sheet = source.sheet_by_index(0)

E= sheet.nrows    # sheet的行数

f = open(resultname, 'w+')

print ('truncate table T_CS_BBXMTX;', file=f)

for A in range (1,E):

 a= str(int(sheet.cell(A, 0).value)) #报表ID
 b= str(int(sheet.cell(A, 1).value)) #报表格式ID
 c= sheet.cell(A, 4).value      #报表项目名称
 d= str(int(sheet.cell(A, 2).value)) #行号
 e= str(int(sheet.cell(A, 3).value)) #列号
 g= str(int(sheet.cell(A, 6).value)) #指标ID
 h= str(int(sheet.cell(A, 8).value)) #计量单位ID

 print ('INSERT INTO `T_CS_BBXMTX` VALUES ('+a+','+b+',\''+c+'\','+d+','+e+','+g+','+h+',2);', file=f)

