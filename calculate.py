import xlwt
book=xlwt.Workbook(encoding='utf-8',style_compression=0)
sheet=book.add_sheet('aa',cell_overwrite_ok=True)
sheet.write(0,0,"English")
sheet.write(1,0,2)
txt1='中文名字'
sheet.write(0,1,txt1)
txt2="spyder"
sheet.write(1,1,txt2)
book.save(r'./calculate.xlsx')
