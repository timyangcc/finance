
import csv
from fpdf import FPDF

def pdate(sdate):
    slist=sdate.split("/")
    return ("中華民國"+slist[0]+"年"+slist[1]+"月"+slist[2]+"日")


class PDF(FPDF):
    def __init__(self):
        FPDF.__init__(self)
        self.right_margin=10
        self.left_margin=10
        self.top_margin=10
        self.bottom_margin=10
        self.a4_pxsize=210
        self.a4_pysize=297-self.top_margin-self.bottom_margin
        self.set_margins(self.left_margin,self.top_margin,self.right_margin)
        self.subpage=0

    def box(self,x1,y1,x2,y2):
        self.line(x1,y1+self.pageoffset,x1,y2+self.pageoffset)
        self.line(x1,y2+self.pageoffset,x2,y2+self.pageoffset)
        self.line(x2,y2+self.pageoffset,x2,y1+self.pageoffset)
        self.line(x1,y1+self.pageoffset,x2,y1+self.pageoffset)

    def pageline(self,x1,y1,x2,y2):
        self.line(x1,y1+self.pageoffset,x2,y2+self.pageoffset)

    def pagexy(self,x,y):
        self.set_xy(x,y+self.pageoffset)
    

    def new_page(self):
        if self.subpage==0:
            self.add_page()
            self.pageoffset= 0
            self.line(0,self.a4_pysize/2+self.top_margin,self.a4_pxsize,self.a4_pysize/2+self.top_margin)
            self.subpage=1
        else:
            self.pageoffset=self.a4_pysize/2
            self.subpage=0
    
def voucher_frontpage(pdf):
    pdf.new_page()
    pdf.set_font('fireflysung','',12)
#    for i in range(27):
#        pdf.set_xy(0,i*10)
#        pdf.cell(10,10,str(i*10))
#        pdf.line(10,10*i,100,10*i)
    pdf.set_font('fireflysung','',16)
    pdf.pagexy(60,20)
    pdf.cell(200,10,'財團法人台北市林坤地仁濟文教基金會')
    pdf.pagexy(80,40)
    pdf.set_font('fireflysung','',24)
    pdf.cell(70,24,'會    計    憑    證')
    pdf.line(75,60,145,60)
    pdf.line(75,61,145,61)
    pdf.pagexy(50,70)
    pdf.set_font('fireflysung','',16)
    pdf.cell(100,10,' 110 年 1 月 1 日至 110 年 12 月 31 日止')
    pdf.pagexy(50,90)
    pdf.cell(100,10,'統一編號:  81588114      稅籍編號:  140010025')

def voucher_page(pdf):
    pdf.new_page()
    pdf.set_font('fireflysung','',16)
    pdf.box(150,15,165,20)
    pdf.box(165,15,190,20)
    pdf.box(150,20,165,25)
    pdf.box(165,20,190,25)
    pdf.pageline(10,0,10,pdf.a4_pysize/2)
    pdf.pagexy(30,15)
    pdf.cell(120,8,'財團法人台北市林坤地仁濟文教基金會')
    pdf.pagexy(150,15)
    pdf.set_font('fireflysung','',12)
    pdf.cell(10,5,'總號')
    pdf.pagexy(150,20)
    pdf.cell(10,5,'分號')
    pdf.pagexy(80,23)
    pdf.set_font('fireflysung','',14)
    pdf.cell(40,5,'轉  帳  傳  票')
    pdf.pagexy(10,35)
    pdf.set_font('fireflysung','',12)
    pdf.cell(10,5,'借貸',1)
    pdf.pagexy(20,35)
    pdf.cell(45,5,'會計科目',1)
    pdf.pagexy(65,35)
    pdf.cell(55,5,'摘  要',1)
    pdf.pagexy(120,35)
    pdf.cell(35,5,'借方金額',1)
    pdf.pagexy(155,35)
    pdf.cell(35,5,'貸方金額',1)
    for i in range(16):
        pdf.box(10,35+(i+1)*5,20,35+(i+2)*5)
        pdf.box(20,35+(i+1)*5,65,35+(i+2)*5)
        pdf.box(65,35+(i+1)*5,120,35+(i+2)*5)
        pdf.box(120,35+(i+1)*5,155,35+(i+2)*5)
        pdf.box(155,35+(i+1)*5,190,35+(i+2)*5)
    pdf.box(10,120,120,125)
    pdf.pagexy(10,120)
    pdf.cell(100,5,'合計')
    pdf.box(120,120,155,125)
    pdf.box(155,120,190,125)
    pdf.pagexy(193,35)
    pdf.box(190,35,200,125)
    pdf.pagexy(193,40)
    pdf.cell(5,5,'附')
    pdf.pagexy(193,45)
    pdf.cell(5,5,'單')
    pdf.pagexy(193,50)
    pdf.cell(5,5,'據')
    pdf.pagexy(193,85)
    pdf.cell(5,5,'張')
    pdf.pagexy(10,130)
    pdf.cell(10,5,'核准')
    pdf.pagexy(40,130)
    pdf.cell(10,5,'會計')
    pdf.pagexy(70,130)
    pdf.cell(10,5,'稽核')
    pdf.pagexy(100,130)
    pdf.cell(10,5,'登帳')
    pdf.pagexy(130,130)
    pdf.cell(10,5,'出納')
    pdf.pagexy(160,130)
    pdf.cell(10,5,'製單')
    


class Voucher:
    def __init__(self):
        self.date='0/0/0'
        self.dorc=''
        self.index=''
        self.code=''
        self.account=''
        self.remark=''
        self.debit=0
        self.credit=0
        
    def month(self):
        s=self.date.split("/")
        return int(s[1])

def loadfbook(filename):
    yearvrec=[]
    with open(filename) as csvfile:
        freader = csv.reader(csvfile)
        count=0
        for row in freader:
            if count==0:
                titlelist=row
            else:
                v = Voucher()
                v.date=row[0]
                v.dorc=row[1]
                v.index=row[2]
                v.code=row[3]
                v.account=row[4]
                v.remark=row[5]
                v.debit=int('0'+row[6])
                v.credit=int('0'+row[7])
                yearvrec.append(v)
            count=count+1
        print("Total records = ",count-1)

    vbook=[]
    lastindex=''
    curv = []
    lastv = []
    for r in yearvrec:
        if r.index != lastindex:
            if curv!=[]:
                vbook.append(curv)
                curv=[]
            curv.append(r)
            lastindex=r.index
        else:
            curv.append(r)
    vbook.append(curv)
    return vbook

def voucherscvsrep(vbook):
    print("Creating Vouchers in CVS")
    turnpage=0
    with open("voucher.csv",'w') as f:
        f.write("\n\n\n\n")
        f.write(',財團法人林坤地仁濟文教基金會\n\n')
        f.write(',會    計   憑   證\n\n')
        f.write(',中華民國    年   月   日 至   年   月   日\n\n')
        f.write('\n')
        f.write('\n 統一編號:  81588114      稅籍編號:  140010025\n\n')
        for i in range(33):
           f.write('\n')
        
        for v in vbook:
            r=v[0]
            vdeb=0
            vcrt=0
            vsize=len(v)
            vpages=0
            vp = 0
            while vsize>0:
                vpages=vpages+1
                if vsize>10:
                    vs=10
                else:
                    vs=vsize
                f.write(',財團法人林坤地仁濟文教基金會\n')
                f.write('                        ,,,  總號,'+r.index+'\n')
                f.write('                        ,,,  分號,'+'-'+str(vpages)+'\n')
                f.write(',轉帳傳票\n')
                f.write(','+pdate(r.date)+'\n')
                f.write('===================================================================\n')
                f.write('借貸,會計科目,摘要,借方金額,貸方金額\n')
                f.write('===================================================================\n')
                for i in v[vp:vp+vs]:
                     f.write(i.dorc+','+i.account+','+i.remark+',\"'+'{:,}'.format(i.credit)+'\",\"'+ '{:,}'.format(i.debit)+'\"\n' )
                     vdeb=vdeb+i.debit
                     vcrt=vcrt+i.credit
                vp=vp+10
                if vsize>10:
                    f.write('===================================================================\n')
                    f.write("合計,,,\n")
                else:
                    for j in range(vsize,10):
                        f.write(",,,\n")
                    f.write('===================================================================\n')
                    f.write("合計,,,\""+'{:,}'.format(vcrt)+'\",\"'+'{:,}'.format(vdeb)+'\"\n')
                f.write('===================================================================\n')
                f.write("核准             會計             覆核             登帳            出納              製單\n")
                f.write("\n")
                turnpage=turnpage+1
                if turnpage%2==0:
                    f.write("\f")
                vsize=vsize-10
    f.close()

def voucherspdfrep(vbook):
    print("Creating Vouchers in PDF")
    turnpage=0

    pdf=PDF()
    pdf.add_font('fireflysung','','fireflysung.ttf',uni=True)

    voucher_frontpage(pdf)
    
    pdf.set_font('fireflysung','',14)

    for v in vbook:
        r=v[0]
        vdeb=0
        vcrt=0
        vsize=len(v)
        vpages=0
        vp = 0
        while vsize>0:
            vpages=vpages+1
            voucher_page(pdf)
            pdf.pagexy(70,30)
            pdf.cell(50,5,pdate(r.date))
            pdf.pagexy(165,15)
            pdf.cell(10,5,r.index)
            pdf.pagexy(165,20)
            pdf.cell(10,5,'-'+str(vpages))
            if vsize>16:
                vs=16
            else:
                vs=vsize
            for i in range(vs):
                iv=v[vp+i]
                pdf.pagexy(10,35+(i+1)*5)
                pdf.cell(5,5,iv.dorc)
                pdf.pagexy(20,35+(i+1)*5)
                pdf.cell(40,5,iv.account)
                pdf.pagexy(65,35+(i+1)*5)
                pdf.cell(50,5,iv.remark)
                pdf.pagexy(120,35+(i+1)*5)
                pdf.cell(35,5,'{:,}'.format(iv.credit),0,0,'R')
                pdf.pagexy(155,35+(i+1)*5)
                pdf.cell(35,5,'{:,}'.format(iv.debit),0,0,'R')
                vdeb=vdeb+iv.debit
                vcrt=vcrt+iv.credit

            pdf.pagexy(120,120)
            pdf.cell(35,5,'{:,}'.format(vcrt),0,0,'R')
            pdf.pagexy(155,120)
            pdf.cell(35,5,'{:,}'.format(vdeb),0,0,'R')
            vsize=vsize-16
            vp=vp+16
    pdf.output('voucher110.pdf','F')    
    

def trypage(f,linecount,pagecount):
    if linecount==0:
        f.write(',財團法人林坤地仁濟文教基金會\n')
        f.write(',,,日    記    簿\n')
        f.write(',,,,,頁次  '+str(pagecount+1)+'\n')
        f.write('=======================================================================\n')
        f.write('日期,編號,科目名稱,摘要,借方金額,貸方金額\n')
        f.write('=======================================================================\n')
        return pagecount+1
    else:
        return pagecount

def incline(linecount):
    if linecount>38:
        return 0
    else:
        return linecount+1

        
def journalrep(vbook):

    
    print("Creating Journal Report")
    linecount=0
    with open("journal.csv","w") as f:
        f.write("\n\n\n\n")
        f.write(',財團法人林坤地仁濟文教基金會\n\n')
        f.write(',   日   記   簿\n\n')
        f.write(',,110 年度\n\n')
        f.write('\n')
        f.write("\n\n負責人:  楊杜清香\n")
        f.write('\n 統一編號:  81588114      稅籍編號:  140010025\n\n')
        f.write('\n')
        f.write('台北市忠孝東路三段136號11樓\n')
        for i in range(27):
            f.write('\n')
        pageno=0
        lastmon=1
        debsum=0
        crtsum=0
        pageno=trypage(f,linecount,pageno)
        for v in vbook:
            curmon=v[0].month()
            if curmon!=lastmon:
                f.write('================================================================\n')
                linecount=incline(linecount)
                pageno=trypage(f,linecount,pageno)
                
                f.write(',,\"'+str(lastmon)+'月小計\"'+',,\"'+'{:,}'.format(crtsum)+'\",\"'+'{:,}'.format(debsum)+'\"\n')
                linecount=incline(linecount)
                pageno=trypage(f,linecount,pageno)
                
                f.write('================================================================\n')
                linecount=incline(linecount)
                trypage(f,linecount,pageno)
                crtsum=0
                debsum=0
                lastmon=curmon
            for i in v:
                f.write(i.date+','+i.index+','+i.account+','+i.remark+',\"'+'{:,}'.format(i.debit)+'\",\"'+ '{:,}'.format(i.credit)+'\"\n')
                debsum=debsum+i.debit
                crtsum=crtsum+i.credit
                linecount=incline(linecount)
                pageno=trypage(f,linecount,pageno)         
            f.write('___________________________________________________________________________\n')
            linecount=incline(linecount)
            pageno=trypage(f,linecount,pageno)
                    
        f.write('===============================================================\n')
        f.write(',,\"'+str(lastmon)+'月小計\"'+',,\"'+'{:,}'.format(crtsum)+'\",\"'+'{:,}'.format(debsum)+'\"\n')
        f.write('===============================================================\n')
                              
    
vbook=loadfbook('book110.csv')
voucherscvsrep(vbook)
voucherspdfrep(vbook)
journalrep(vbook)
    
