from datetime import datetime

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

    def merge_item(self, bitem):
        'Reture true if merge successful'
        if self.amount != bitem.amount:
            return False
        if self.match_time(bitem.time) == False:
            return False
        if self.des != bitem.des:
            return False       
        return True

    def match_time(self, time_t):
        date0 = datetime.strftime (self.time, '%Y-%m-%d')
        date1 = datetime.strftime (time_t, '%Y-%m-%d')
        if date0 != date1:
            return False
        time0 = datetime.strftime (self.time, '%H:%M:%S')
        time1 = datetime.strftime (time_t, '%H:%M:%S')
        if time0 == time1 or time0=='00:00:00' or time1=='00:00:00':
            return True
        return False

if __name__ == '__main__':
    pass

    
