import xlwt

class Style():
    user = xlwt.XFStyle()
    user.font.colour_index = 0
    def __init__(self):
        self.user.alignment.horz = 0x02
        self.user.alignment.vert = 0x02
        self.user.alignment.wrap = 1


 #1，支出者   2，日期   3，分类  4,细类 5，金额 6，备注
class Acount(Style):
    Acount_list = []
    def __int__(self,color):
        super(Acount,self).__init__()



