from xlwt import Workbook 
import os
from datetime import datetime
import re

from Bitem import Bitem
from ClassifyRule import JsonParser

class Bill():
    def __init__(self):
        self.conts = []
    
    def parse(self, fname):
        'parse bill content from excel(fname)'
        pass
    
    def add_bill(self, bill):
        for bitem in bill.conts:
            self.add_item(bitem)

    def add_item(self, bitem):
        'insert bitem to bill'
        for i in range(0, len(self.conts)):
            if self.conts[i].merge_item(bitem):
                return
        self.conts.append(bitem)
    
    def save2excel(self, fname='bill.xls'):
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

    def classify(self, rule_fnames):
        rules = []
        for rule_fname in rule_fnames:
            rules.extend(JsonParser(rule_fname))
        for i in range(0, len(self.conts)):
            for rule in rules:
                self.conts[i].classify(rule)
        pass
    
    def get_eml_charset(self, eml):
        pattern = re.compile(r'charset="(.*?)"')
        res = pattern.findall(eml)
        return res[0]

    def sort_time(self):
        self.conts.sort(key=Bitem.get_time)
        pass

def ParseBillFolder(bill_dir):
    from BillAliPay import BillAliPay
    from BillCMB import BillCMB
    from BillCOMM import BillCOMM
    from BillSPDB import BillSPDB

    bills = []

    files = os.listdir(bill_dir)
    for f in files:
        if 'AliPay' in f:
            bill = BillAliPay()
        elif 'CMB' in f:
            bill = BillCMB()
        elif 'COMM' in f:
            bill = BillCOMM()
        elif 'SPDB' in f:
            bill = BillSPDB()
        else:
            continue  
        bill.parse(bill_dir+f)
        bills.append(bill)

    # Merge Bills
    bill = Bill()
    for bill_t in bills:
        bill.add_bill(bill_t)

    rules = ['./rule/rule_high.json', './rule/rule_mid.json', './rule/rule_low.json']
    bill.classify(rules)

    return bill 

if __name__ == '__main__':
    bill = ParseBillFolder('./bill/')
    bill.sort_time()
    bill.save2excel()

    
