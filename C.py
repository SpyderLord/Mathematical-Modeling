import xlwt
import xlrd
# import xlutils.copy

#创建一个新的文件夹
workbook=xlwt.Workbook(encoding='utf-8',style_compression=0)
sheet=workbook.add_sheet('test',cell_overwrite_ok=True)
sheet.write(0,0,"value")

#读取原始数据
data=xlrd.open_workbook(r'./4.xlsx')
table=data.sheets()[0]
ncols=table.ncols
nrows=table.nrows
cell2=table.cell_value(1,2)
# print(type(cell2))
print("row%d"%ncols)
for i in range(nrows-1):
    cell=table.cell_value(i+1,1)
    if cell==0:
        sheet.write(i+1,0,0)
    else:
        cell=table.cell_value(i+1,2)
        if type(cell)==str:
            sheet.write(i+1,0,1)
        else:
            sheet.write(i+1,0,0)

workbook.save(r'./222-2.xlsx')