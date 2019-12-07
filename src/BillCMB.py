# encoding: utf-8
from Bitem import Bitem
from Bill import Bill

from datetime import datetime

import quopri
import re
from warnings import warn

class BillCMB(Bill):

    def parse(self, fname):
        'parse bill content from eml(fname)'

        f = open(fname, 'rb')
        eml_bytes = f.read()
        f.close()

        # quoted-printable decode
        quo_dec = quopri.decodestring(eml_bytes)
        quo_dec = quo_dec.decode('utf-8')

        self.parse_year(quo_dec)
        self.parse_contents(quo_dec)
    
    def parse_year(self, quo_dec):
        pattern = re.compile(r'zdzq.jpg.*?(\d{4})/\d{2}/\d{2}-\d{4}/\d{2}/\d{2}', re.S)
        res = pattern.findall(quo_dec)[0]
        self.year = res

    def parse_contents(self, quo_dec):
        pattern = re.compile(r'rmbbt.jpg(.*?)★', re.S)
        res = pattern.findall(quo_dec)[0]
        pattern = re.compile(r'<font face="宋体" style="font-size:12px;line-height:120%;">(.*?)</font></div>')
        res = pattern.findall(res)

        if len(res) % 7 != 0:
            warn('tabel numbers mismatch!')
        
        for i in range(0, len(res)//7):
            target = res[i*7:(i+1)*7]
            
            des = target[2]
            if '支付宝' in des:
                continue

            time = datetime.strptime(self.year+target[0], '%Y%m%d')
            
            amount = amount = target[3] # ￥&nbsp;54.99
            amount = amount.replace('￥', '')
            amount = amount.replace('&nbsp', '')
            amount = amount.replace(';', '')
            amount = amount.replace(' ', '')
            amount = float(amount)
            
            bitem = Bitem(amount=amount, time=time, des=des, src='CMB')
            self.conts.append(bitem)

            pass

if __name__ == '__main__':
    bill = BillCMB()
    bill.parse('./Bill/CMB-201910.eml')
    bill.save()

