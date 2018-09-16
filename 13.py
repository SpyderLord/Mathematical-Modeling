import xlwt
import xlrd
# import xlutils.copy

#创建一个新的文件夹
workbook=xlwt.Workbook(encoding='utf-8',style_compression=0)
sheet=workbook.add_sheet('test',cell_overwrite_ok=True)
sheet.write(0,0,"value")

#读取原始数据
data=xlrd.open_workbook(r'./选择-3.xlsx')
table=data.sheets()[0]
ncols=table.ncols
nrows=table.nrows
cell2=table.cell_value(1,13)
# print(cell2)
# print("row%d"%nrows)
for i in range(nrows-1):
    cell1=table.cell_value(i+1,13)
    cell2=table.cell_value(i+1,14)
    if cell1==0:
        sheet.write(i+1,13,0)
    else:
        sheet.write(i+1,13,1)
    if cell2==0:
        sheet.write(i+1,14,0)
    else:
        sheet.write(i+1,14,1)

workbook.save(r'./444.xlsx')
