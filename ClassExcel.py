import xlrd
import sys
from datetime import date

class BaseClassExcel:
    ''' excel基础类 '''

        # 初始化，判断excel文件是否存在，并获取工作薄对象
    def __init__(self,file_name):
        
        self.file_name = file_name # excel文件名
        self.objWorkBook = None # excel文件的workbook对象
        self.nrows = 0 # 表总行数
        self.ncols = 0 # 表总列数

    # 获取workbook对象
    def get_obj_workbook(self):
        try:
            return xlrd.open_workbook(self.file_name)
        except OSError as reason:
            print("===>" + str(reason))
            sys.exit(1)

    # 获取所有sheet名字列表
    def get_sheet_names(self,objWorkbook):
        return objWorkbook.sheet_names()

    # 获取sheet_names[]成员数量
    def get_sheetnames_number(self,objWorkbook):
        return objWorkbook.nsheets        

    # 获取工作表sheet对象
    def get_obj_sheet(self,objWorkbook,sheet_name):
        return objWorkbook.sheet_by_name(sheet_name)

    # 获取指定单元格数据
    def get_cell_value(self,objSheet,x,y):
        return objSheet.cell(x,y).value

    # 获取指定单元格数据类型
    # ctype: 0 empty;1 string;2 number;3 date;4 boolean;5 error
    def get_cell_ctype(self,objSheet,x,y):
        return objSheet.cell(x,y).ctype

    # 获取表行数和列数
    def get_nrows_ncols(self,objSheet):
        self.nrows = objSheet.nrows
        self.ncols = objSheet.ncols       

    # 日期格式转换为字符串
    def date_to_str(self,date_value,objWorkbook):
        tupDate = xlrd.xldate_as_tuple(date_value,objWorkbook.datemode)
        return date(*tupDate[:3]).strftime("%Y/%m/%d")