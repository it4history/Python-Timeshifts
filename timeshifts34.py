import xlrd, xlwt
#from difflib import Differ, get_close_matches
from difflib import get_close_matches
from threading import Thread

#from xlutils.filter import GlobReader,BaseFilter,DirectoryWriter,process
import os
import re                
recomp = re.compile(r"""
\d+
|(бря|бре|июня|июля|июне|июле)
|(-|–|—)
""",re.VERBOSE)

myfile = ["00_orig_00.xls", "00_orig_01.xls", "00_orig_02.xls", "00_orig_03.xls"]
new_myfile = ['00_out_00.xls','00_out_01.xls','00_out_02.xls','00_out_03.xls']
#mydir='C:/Users/drondin/Documents/'
mydir='C:/Users/Im Nox/Documents/'

shifts = [83,167,251,334,417,501,18,23,36,41,59,64,83,100,72,76,77,108]
exceptions = ["а","А",
            "был","Был","была","Была","были","Были",
            "в","В","во","велик","велика",
            "г.","где","год","год.","Год","году","году.","Году",
            "до","До","если","Если",
            "и","И","из","Из","или","Или",
            "к","К","как","Как","которого","который","которая",
            "л.","ли","либо","Либо","м.","мы",
            "на","На","нас","Нас","не","нет","них","н.э.","о","О","он","Он",
            "она","Она","они", "Они","оно","Оно","от","От","об","Об","перед","Перед","по","По","р.",
            "с","С","стал","стало","Стало",
            "также","Также","то","То","тоже","Тоже","х.","что","Что","Чтобы","чтобы","это","Это","этом","Этом"]

def Filter(beg, end, f1, f2, types):
    MTCH = 1
    rows_count = 0

    font0 = xlwt.Font()
    font0.bold = False
    style0 = xlwt.XFStyle()
    style0.font = font0

    wb = xlwt.Workbook()
    ws = wb.add_sheet('A Test Sheet')

    rb = xlrd.open_workbook(os.path.join(mydir, myfile[f1]), formatting_info=True)
    sheet = rb.sheet_by_index(0)
    #vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
    nrows = sheet.nrows
    years = [sheet.row_values(rownum)[0] for rownum in range(nrows)]
    years.reverse()
    l_years = len(years)
    
    for i in range(beg, end):
        if rows_count > 65536:
            print ("rows count > 65536 in thread {0}".format(f1+1))
            break
        print(i)
        if types == "B":
            row_v_i = sheet.row_values(nrows-i)[1]
            row_v_i_nn = recomp.sub('', row_v_i)
            print (row_v_i_nn)
            for j in range(i, l_years):
                abs_years = abs(years[i] - years[j])
                if abs_years in shifts:
                    row_v_j = sheet.row_values(nrows-j)[1]
                    new = [abs_years, years[i], years[j], row_v_i, row_v_j]
                    m = []
                    for k in row_v_i_nn.split():
                        if not k in exceptions:
                            matches = get_close_matches(k, row_v_j.split())
                            m.extend(matches)
                    l_mtch = len(m)
                    if l_mtch >= MTCH:
                        new.append(l_mtch)
                        for l in range(0, len(new)):
                            ws.write(rows_count, l, new[l], style0)
                        rows_count += 1
        elif types == "C":
            row_v_i = sheet.row_values(nrows-i)[1]
            row_v_i2 = sheet.row_values(nrows-i)[2]
            for j in range(i, l_years):
                abs_years = abs(years[i] - years[j])
                if abs_years in shifts:
                    row_v_j = sheet.row_values(nrows-j)[1]
                    row_v_j2 = sheet.row_values(nrows-j)[2]
                    new = [abs_years, years[i], years[j], row_v_i, row_v_j, row_v_i2, row_v_j2]
                    if row_v_i2 == row_v_j2:
                        for l in range(0, len(new)):
                            ws.write(rows_count, l, new[l], style0)
                        rows_count += 1
        wb.save(os.path.join(mydir, new_myfile[f2]))

t1 = Thread(target=Filter, args=(4840,5001,0,0,"C"))
t2 = Thread(target=Filter, args=(5403,5601,1,1,"C"))
t3 = Thread(target=Filter, args=(5936,6201,2,2,"C"))
t4 = Thread(target=Filter, args=(6548,6801,3,3,"C"))

t1.start()
t2.start()
t3.start()
t4.start()
t1.join()
t2.join()
t3.join()
t4.join()

'''
class ShiftsFilter(BaseFilter):

    goodlist = []
    
    def __init__(self,elist): 
        self.goodlist = goodlist
        self.wtw = 0
        self.wtc = 0
         

    def workbook(self, rdbook, wtbook_name): 
        self.next.workbook(rdbook, 'filtered_'+wtbook_name) 

    def row(self, rdrowx, wtrowx):
        pass

    def cell(self, rdrowx, rdcolx, wtrowx, wtcolx):
        value = self.rdsheet.cell(rdrowx,rdcolx).value
        if value in self.goodlist:
            self.wtc=self.wtc+1 
            self.next.row(rdrowx,wtrowx)
        else:
            return
        self.next.cell(rdrowx,rdcolx,self.wtc,wtcolx)
        
        
data = """somedata1
somedata2
somedata3
somedata4
somedata5
"""

goodlist = data.split("\n")

process(GlobReader(os.path.join(mydir,myfile)), ShiftsFilter(goodlist), DirectoryWriter(mydir))'''

