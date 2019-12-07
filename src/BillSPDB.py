# encoding: utf-8
from datetime import datetime
import xlrd

from Bitem import Bitem
from Bill import Bill

class BillSPDB(Bill):

    def parse(self, fname):
        'parse bill content from SPDB xls'
        excel = xlrd.open_workbook(fname)
        sheet = excel.sheet_by_index(0)
        
        nrows = sheet.nrows
        
        for i in range(1, nrows):
            row = sheet.row_values(i)
            if len(row)==7:
                self.parse_content(row)
    
    def parse_content(self, row):
        des = row[2]
        if '支付宝' in des:
            return
        time = datetime.strptime(row[0], '%Y%m%d %H:%M:%S') 
        amount = -float(row[6])

        bitem = Bitem(amount=amount, time=time, des=des, src='SPDB')
        self.conts.append(bitem)

if __name__ == '__main__':
    bill = BillSPDB()
    bill.parse('./Bill/SPDB-201911.xls')
    bill.save()

