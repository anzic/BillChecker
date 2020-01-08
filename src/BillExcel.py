import xlwt
import xlrd
from datetime import datetime
import os

from Bitem import Bitem
from Bill import Bill
from Bill import ParseBillFolder

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
        name = '%04d-%02d' %(year, month)
        BillSheet.__init__(self, name)

    def add_item(self, item):
        time_str = item.time.strftime('%Y-%m')
        if time_str == self.name:
            Bill.add_item(self, item)

class BillExcel():
    def __init__(self, name):
        self.name = name
        self.sheets = []

    def load(self):
        if not os.path.exists(self.name):
            return
        book = xlrd.open_workbook(self.name)
        for i in range(0, book.nsheets):
            sheet = book.sheet_by_index(i)
            try:
                year = int(sheet.name[0:4])
                month = int(sheet.name[5:])
            except:
                continue
            bsheet = BillMonth(year, month)
            bsheet.load(book.sheet_by_index(i))
            self.sheets.append(bsheet)

    def save2excel(self, fname=None):
        if fname == None:
            fname = self.name
        wb = xlwt.Workbook()
        for sheet in self.sheets:
            sheet.save(wb)
        wb.save(fname)
    
    def add_bill(self, bill):
        for item in bill.conts:
            self.add_item(item)

    def add_item(self, item):
        sheet_name = item.time.strftime('%Y-%m')
        for bsheet in self.sheets:
            if sheet_name == bsheet.name:
                bsheet.add_item(item)
                return

        bsheet = BillMonth(item.time.year, item.time.month)
        bsheet.add_item(item)
        self.sheets.append(bsheet)
        pass

if __name__ == '__main__':
    bill = ParseBillFolder('./bill/')
    bill.sort_time()

    excel = BillExcel('bill.xls')
    excel.load()
    excel.add_bill(bill)

    excel.save2excel()

