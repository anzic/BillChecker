import xlwt
import xlrd
from datetime import datetime

from Bitem import Bitem
from Bill import Bill

class BillSheet(Bill):
    def __init__(self, name=None):
        Bill.__init__(self)
        self.name = name

    def load(self, sheet):
        self.name = sheet.name
        for i in range(1, sheet.nrows):
            row = sheet.row_values(i)
            time = datetime.strptime(row[0]+row[1], '%Y-%m-%d%H:%M:%S')
            self.conts.append(Bitem(btype=row[2], amount=row[3], time=time, des=row[5], src=row[6]))

    def save(self, wb):
        ws = wb.add_sheet(self.name)
        self.write_sheet_head(ws)

        for i in range(0, len(self.conts)):
            self.conts[i].write_sheet(ws, i+1)

class BillMonth(BillSheet):
    def __init__(self, year, month):
        name = str(year)+'%02d'%(month)
        BillSheet.__init__(self, name)
    
    def add_bill(self, bill):
        for item in bill.conts:
            time_str = item.time.strftime('%Y%m')
            if time_str == self.name:
                self.conts.append(item)


class BillExcel():
    def __init__(self, name):
        self.name = name
        self.sheets = []

    def load(self):
        book = xlrd.open_workbook(self.name)
        for i in range(0, book.nsheets):
            bsheet = BillSheet()
            bsheet.load(book.sheet_by_index(i))
            self.sheets.append(bsheet)

    def save(self, fname='bill.xls'):
        wb = xlwt.Workbook()
        for sheet in self.sheets:
            sheet.save(wb)
        wb.save(fname)
        
if __name__ == '__main__':
    excel = BillExcel('bill.xls')
    excel.load()
    excel.save()

