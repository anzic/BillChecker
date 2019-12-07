from xlwt import Workbook 
from ClassifyRule import JsonParser

class Bitem():
    def __init__(self, btype='', amount=0, time=None, des='', src=''):
        self.type = btype
        self.amount = amount
        self.time = time
        self.des = des
        self.src = src

    def write_sheet(self, sheet, idx):
        'write item to sheet'
        sheet.write(idx, 0, self.time.strftime('%Y-%m-%d'))
        sheet.write(idx, 1, self.time.strftime('%H:%M:%S'))
        if self.amount > 0:
            sheet.write(idx, 2, 'Income')
        else:
            sheet.write(idx, 2, 'Expense')

        sheet.write(idx, 3, self.amount)
        sheet.write(idx, 4, self.type)
        sheet.write(idx, 5, self.des)
        sheet.write(idx, 6, self.src)

    def classify(self, rule):
        'classify item with rule'
        attr_dict = {'type':self.type, 'amount':self.amount, 'time':self.time, 'des':self.des, 'src':self.src}
        attr = attr_dict[rule.attr]

        if rule.relation.lower() == 'contain':
            if rule.value in attr:
                self.type = rule.type
        else:
            print('not support current rule.relation:'+rule.relation)

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
        self.write_sheet_head(sheet1)

        for i in range(0, len(self.conts)):
            self.conts[i].write_sheet(sheet1, i+1)
        
        wb.save(fname) 

    def write_sheet_head(self, sheet):
        Headers = ['Date', 'Time', 'Income/Expense', 'Amount', 'Type', 'Description', 'Bill Source']
        for i in range(0, len(Headers)):
            sheet.write(0, i, Headers[i])

    def deduplicate(self):
        'remove duplicated items'
        pass

    def classify(self, rule_fname):
        rules = JsonParser(rule_fname)
        for i in range(0, len(self.conts)):
            for rule in rules:
                self.conts[i].classify(rule)
        pass




if __name__ == '__main__':
    pass
    
