# encoding: utf-8
from Bitem import Bitem
from Bill import Bill

import csv
from datetime import datetime

class BillAliPay(Bill):

    def parse(self, fname):
        'parse bill content from excel(fname)'
        with open(fname) as f:
            f_csv = csv.reader(f)
            
            # Paeser Header
            while(True):
                header = next(f_csv)
                if len(header) > 10:
                    break
            self.parse_header(header)

            # Parser Content
            while(True):
                row = next(f_csv)
                if len(row) != len(header):
                    break
                bitem = self.parse_row(row)
                if bitem:
                    self.conts.append(bitem)

    def parse_header(self, header):
        idx = 0
        for t in header:
            if u'交易创建时间' in t:
                self.col_idx_time = idx
            elif u'金额' in t:
                self.col_idx_amount0 = idx
            elif u'收/支' in t:
                self.col_idx_amount1 = idx
            elif u'交易对方' in t:
                self.col_idx_des0 = idx
            elif u'商品名称' in t:
                self.col_idx_des1 = idx        
            idx += 1

    def parse_row(self, row):
        
        amount = float(row[self.col_idx_amount0])
        if u'支出' in row[self.col_idx_amount1]:
            amount = -amount
        elif u'收入' in row[self.col_idx_amount1]:
            amount = amount       
        else:
            return None

        time = datetime.strptime(row[self.col_idx_time], '%Y-%m-%d %H:%M:%S ')
        des = row[self.col_idx_des0].rstrip() + ' ' + row[self.col_idx_des1].rstrip()

        bitem = Bitem(amount=amount, time=time, des=des, src='AliPay')

        return bitem


if __name__ == '__main__':
    bill = BillAliPay()
    bill.parse('./bill/AliPay-201911.csv')
    bill.classify('./rule/classify_rule.json')
    bill.save()
    
