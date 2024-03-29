import xlrd

class GetData:
    filename = '' #xls文件路径
    workbook = None  #工作薄
    sheetNames = []  #表名列表
    sheet = None
    rows = 0  #行数
    cols = 0  #列数

    def __init__(self):
        pass

    def __init__(self,file):
        self.filename = file
        self.workbook = xlrd.open_workbook(file)
        self.sheetNames = self.workbook.sheet_names()

    def printSheets(self):
        print("当前xls文件的所有表名以及行数和列数:\n")
        for sheet in self.sheetNames:
            sheet = self.workbook.sheet_by_name(sheet)
            print("表名     rows     cols\n"+sheet.name+"     "+str(sheet.nrows)+"     "+str(sheet.ncols))

    def chooseSheet(self,sheetName):
        self.sheet = self.workbook.sheet_by_name(sheetName)

    def printTheRowValue(self,row):
        data = self.sheet.row_values(row)
        print("第" + str(row) + "行数据：")
        for i in range(len(data)):
            print(str(i)+". "+data[i])

    def ColValueWrite(self,col,oname):
        data = self.sheet.col_values(col)
        name = oname
        for i in data:
            with open(oname,'a',encoding='utf-8') as f:
                f.write(str(i)+"\n")
        print("第"+str(col)+"列数据写入成功!"+"filename: "+name)