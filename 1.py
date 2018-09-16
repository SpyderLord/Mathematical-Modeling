import xlrd
import xlutils.copy
import xlwt

# 创建一个新的文件夹
workbook=xlwt.Workbook(encoding='utf-8',style_compression=0)
sheet=workbook.add_sheet('test',cell_overwrite_ok=True)
sheet.write(0,0,"name")

#读取原始文件夹
data=xlrd.open_workbook(r'./选择-2.xlsx')
table=data.sheets()[0]
ncols=table.ncols
nrows=table.nrows

fileName='./num.txt'
print("列数等于%d 行数等于%d"%(ncols,nrows))

for i in range(nrows-1):
    with open(fileName, "w")as f:
        #进行每行的搜索
        cell_1=table.cell_value(i+1,1)
        if cell_1==0:
            f.write('0')

            # f.write('\n')

# workbook.save(r'./2.xlsx')

