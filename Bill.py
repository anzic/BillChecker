import xlwt 
from xlwt import Workbook 

class Bitem():
    def __init__(self, btype=None, amount=None, time=None, des=None, src=None):
        self.type = btype
        self.amount = amount
        self.time = time
        self.des = des
        self.src = src

class Bill():
    def __init__(self):
        self.conts = []
    
    def parse(self, fname):
        'parse bill content from excel(fname)'
        pass
    
    def extend(self, bill_t):
        'extend bill with bill_t'
        self.conts.extend(bill_t.conts)
    
    def save(self, fname='bill.xls'):
        'save bill to excel'
        wb = Workbook()
        sheet1 = wb.add_sheet('Sheet 1')
        sheet1.write(0, 0, 'Date')
        sheet1.write(0, 1, 'Time')
        sheet1.write(0, 2, 'Income/Expense')
        sheet1.write(0, 3, 'Amount')
        sheet1.write(0, 4, 'Type')
        sheet1.write(0, 5, 'Description')
        sheet1.write(0, 6, 'Bill Source')
        
        idx = 1
        for item in self.conts:
            sheet1.write(idx, 0, item.time.strftime('%Y-%m-%d'))
            sheet1.write(idx, 1, item.time.strftime('%H:%M:%S'))
            if item.amount > 0:
                sheet1.write(idx, 2, 'Income')
            else:
                sheet1.write(idx, 2, 'Expense')

            sheet1.write(idx, 3, item.amount)
            sheet1.write(idx, 4, item.type)
            sheet1.write(idx, 5, item.des)
            sheet1.write(idx, 6, item.src)
            idx += 1
        
        wb.save(fname) 

    def deduplicate(self):
        'remove duplicated items'
        pass

    def autoclassify(self, key_fname='my.key'):
        pass




if __name__ == '__main__':
    pass
    
