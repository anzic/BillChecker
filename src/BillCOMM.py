# encoding: utf-8
from Bitem import Bitem
from Bill import Bill

from datetime import datetime

import base64
import re
from warnings import warn

class BillCOMM(Bill):

    def parse(self, fname):
        'parse bill content from eml(fname)'

        f = open(fname, 'rt')
        eml = f.read()
        f.close()

        pattern = re.compile(r'\n\n(.*)', re.S)
        eml = pattern.findall(eml)[0]
        eml = base64.b64decode(eml)
        eml = eml.decode('gbk')

        pattern = re.compile(r'还款、退货、费用返还明细(.*?)消费、取现、其他费用明细(.*)', re.S)
        (eml_income, eml_expense) = pattern.findall(eml)[0]

        self.parse_content(eml_income, 1)
        self.parse_content(eml_expense, -1)
    
    def parse_content(self, eml_str, sign):
        pattern = re.compile(r'''<td style="font-family: 'Hiragino Sans GB'"><span>(.*?)</span></td>''')
        res = pattern.findall(eml_str)

        if len(res) % 6 != 0:
            warn('tabel numbers mismatch!')
        
        for i in range(0, len(res)//6):
            target = res[i*6:(i+1)*6]
            
            des = target[3]
            if '支付宝' in des:
                continue       

            time = datetime.strptime(target[0], '%Y/%m/%d')
            amount = amount = target[4] # RMB 1.50
            amount = amount.replace('RMB', '')
            amount = amount.replace(' ', '')
            amount = float(amount)*sign
            
            bitem = Bitem(amount=amount, time=time, des=des, src='COMM')
            self.conts.append(bitem)


if __name__ == '__main__':
    bill = BillCOMM()
    bill.parse('./Bill/COMM-201909.eml')
    bill.save()

