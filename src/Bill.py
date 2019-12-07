from xlwt import Workbook 
import os
from datetime import datetime

from Bitem import Bitem
from ClassifyRule import JsonParser

class Bill():
    def __init__(self):
        self.conts = []
    
    def parse(self, fname):
        'parse bill content from excel(fname)'
        pass
    
    def merge(self, bill_t):
        'merge bill with bill_t'
        for bitem in bill_t.conts:
            self.merge_item(bitem)

    def merge_item(self, bitem):
        'insert bitem to bill'
        for i in range(0, len(self.conts)):
            if self.conts[i].merge_item(bitem):
                return
        self.conts.append(bitem)
    
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

    def classify(self, rule_fname):
        rules = JsonParser(rule_fname)
        for i in range(0, len(self.conts)):
            for rule in rules:
                self.conts[i].classify(rule)
        pass

def ParseBillFolder(folder):
    from BillAliPay import BillAliPay
    from BillCMB import BillCMB
    from BillCOMM import BillCOMM

    bills = []

    bill_dir = r'./bill/'
    files = os.listdir(bill_dir)
    for f in files:
        if f.startswith('AliPay-'):
            bill = BillAliPay()
        elif f.startswith('CMB-'):
            bill = BillCMB()
        elif f.startswith('COMM-'):
            bill = BillCOMM()
        else:
            continue  
        bill.parse(bill_dir+f)
        bills.append(bill)

    # Merge Bills
    bill = Bill()
    for bill_t in bills:
        bill.merge(bill_t)

    bill.classify('./rule/classify_rule.json')

    return bill 

if __name__ == '__main__':
    bill = ParseBillFolder('./bill/')
    bill.save()

    
