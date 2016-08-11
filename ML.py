import xlrd
import datetime
from fileXL import *
class ML(fileXL):
    def __init__(self,name_file):
        fileXL.__init__(self,name_file)
        fil = xlrd.open_workbook(self.name_file)
        sheet = fil.sheet_by_index(0)
        self.device = sheet.row_values(6)[41]
        self.firm = sheet.row_values(8)[41]
        self.device_name = sheet.row_values(7)[41]
        self.ingerer = sheet.row_values(16)[48]
        self.quantity = int(sheet.row_values(19)[31])
        var = sheet.row_values(16)[43]
        if var != '':
            year, month, day = xlrd.xldate_as_tuple(var,0)[:3]
            self.data_vk_end = datetime.date(year, month, day)
        else:
            self.data_vk_end = ''
