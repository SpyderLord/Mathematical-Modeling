import xlwt
import xlrd
import argparse #从控制台中进行参数的解析

#本文件实现的是将excel中的空格转换成为数字0，如果是数字则保持不变，生成的数列保存在新的文件中

def blank2zero(loadpath,index,savepath):
    '''

    :param loadpath:读取文件的路径
    :param index: excel中需要转换的文件的列index
    :param savepath: 文件保存路径
    :return:
    '''
    #创建一个新的表格
    workbook=xlwt.Workbook(encoding='utf-8',style_compression=0)
    sheet=workbook.add_sheet('1',cell_overwrite_ok=True)
    sheet.write(0,0,'value')

    #读取原始数据
    data=xlrd.open_workbook(loadpath)
    table=data.sheets()[0]
    ncols=table.ncols
    nrows=table.nrows

    for i in range(nrows):
        cell=table.cell_value(i+1,index)
        if type(cell)==str:
            sheet.write(i,index,1)
        else:
            sheet.write(i,index,cell)
    workbook.save(savepath)

if __name__=='__main__':
    # parser=argparse.ArgumentParser()
    # parser.add_argument('loadpath',type=str)
    # parser.add_argument('index',type=int)
    # parser.add_argument('savepath',type=str)
    # args=parser.parse_args()
    # blank2zero(args.loadpath,args.index,args.savepath)
    blank2zero('./60001-x.xlsx',10,'K2.xlsx')