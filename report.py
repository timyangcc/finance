
import gi
gi.require_version("Gtk","3.0")
from gi.repository import Gtk
import json
import account
import voucher
import csv
import __main__

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

def open_pdf(title):
    c=canvas.Canvas(title+".pdf")
    return c

def pdf_prttitle(c,tit):
    c.setFont('新細明體',16)
    c.drawString(85,800,tit)
    c.grid((10,60,120,180,270,410,470,530,590),(780,760))
    c.setFont('新細明體',12)
    c.drawString(12,765,"編號")
    c.drawString(62,765,"日期")
    c.drawString(122,765,"科目")
    c.drawString(182,765,"科目名稱")
    c.drawString(272,765,"摘要")
    c.drawString(412,765,"借方金額")
    c.drawString(472,765,"貸方金額")
    c.drawString(532,765,"發票日期")
                     
        

def pdf_prtln(c,line,lst):
     lpos=780-(line+1)*20
     c.grid((10,60,120,180,270,410,470,530,590),(lpos,lpos-20))
     c.drawString(12,lpos-15,lst[0])
     c.drawString(62,lpos-15,lst[1])
     c.drawString(122,lpos-15,lst[2])
     c.drawString(182,lpos-15,lst[3])
     c.drawString(272,lpos-15,lst[4])
     c.drawString(412,lpos-15,"{:,d}".format(lst[5]))
     c.drawString(472,lpos-15,"{:,d}".format(lst[6]))
     c.drawString(532,lpos-15,lst[7])
               

def account_sums(year):
    asum={}
    dbt =0
    crt=0
    for a in account.accounts.accountbook:
        asum[a]=[account.accounts.accountbook[a],0,0,[]]
    for v in voucher.voucherbook.yearbook[year]:
        for r in  voucher.voucherbook.yearbook[year][v].records:
            x = asum[r.code]
            x[1]=x[1]+r.credit
            x[2]=x[2]+r.debit
            x[3].append(r)
            asum[r.code]=x
            crt=crt+r.credit
            dbt=dbt+r.debit
    return asum

class ReportWin(Gtk.Window):
    def __init__(self):
        super().__init__(title="報告")
        self.set_size_request(500,800)
        box=Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.add(box)
        box1=Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)
        box.pack_start(box1,False,False,0)
        
        self.yent=Gtk.Entry()
        ybtn=Gtk.Button(label="載入年度")
        ybtn.connect("clicked",self.loadyear)
        box1.add(self.yent)
        box1.add(ybtn)

        box2=Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)
        box.pack_start(box2,False,False,0)
        inoutbtn=Gtk.Button(label="收支餘絀表")
        box2.pack_start(inoutbtn,False,False,0)
        inoutbtn.connect("clicked",self.inoutreport)
        inoutdetailbtn=Gtk.Button(label="收支細目表")
        box2.pack_start(inoutdetailbtn,False,False,0)
        inoutdetailbtn.connect("clicked",self.inoutdetreport)
        
        self.store=Gtk.ListStore(str,str,str,str)
        scw=Gtk.ScrolledWindow()
        box.pack_start(scw,True,True,0)
        treeview=Gtk.TreeView(model=self.store)
        renderer=Gtk.CellRendererText()
        col1=Gtk.TreeViewColumn("科目",renderer,text=0)
        treeview.append_column(col1)
        col2=Gtk.TreeViewColumn("科目名稱",renderer,text=1)
        treeview.append_column(col2)
        col3=Gtk.TreeViewColumn("借方",renderer,text=2)
        treeview.append_column(col3)
        col4=Gtk.TreeViewColumn("貸方",renderer,text=3)
        treeview.append_column(col4)
        treeview.connect("row-activated",self.openaccount)
        scw.add(treeview)

    def openaccount(self,tree,path,col):
        treeiter=self.store.get_iter(path)
        self.selectedacc=self.store.get_value(treeiter,0)
        r=self.asum[self.selectedacc]
        x=r[3]
        accwin=AccountWin()
        if r[1]>=r[2]:
            sumx=r[1]-r[2]
            dc = "借"
        else:
            sumx=r[2]-r[1]
            dc="貸"
        accwin.set_sum(r[1],r[2],dc,sumx)
        title=x[0].code+x[0].account
        accwin.set_title(title)
        
        accwin.set_content(x)
        accwin.show_all()

    def inoutdetreport(self,widget):
        yt=self.yent.get_text()
        year=int(yt)
        inrep = []
        outrep =[]
        for v in voucher.voucherbook.yearbook[year]:
             for r in  voucher.voucherbook.yearbook[year][v].records:
                  if  account.accounts.accountbook[r.code][1] == "收入":
                      inrep.append(r)
                  elif  account.accounts.accountbook[r.code][1]=="支出":
                      outrep.append(r)
        detwin=DetailedWin(yt+"年")
        detwin.loaddata(inrep,outrep)
        detwin.show_all()
       
                      
    def inoutreport(self,widget):
        catsum={}
        detsum={}
        for ai in self.asum:
            a=self.asum[ai]
            cat = a[0][1]
            det = a[0][2]
            crt = a[1]
            dbt = a[2]
            if cat in catsum:
                catsum[cat][0]=catsum[cat][0]+crt
                catsum[cat][1]=catsum[cat][1]+dbt
                catsum[cat][2]=catsum[cat][0]-catsum[cat][1]
            else:
                catsum[cat]=[crt,dbt,crt-dbt,[]]

            if det in detsum:
                detsum[det][0]=detsum[det][0]+crt
                detsum[det][1]=detsum[det][1]+dbt
                detsum[det][2]=detsum[det][0]-detsum[det][1]
            else:
                detsum[det]=[crt,dbt,crt-dbt,cat,0.0]
        for c in catsum:
            for d in detsum:
                if detsum[d][3]==c:
                    detsum[d][4]='{:.2%}'.format(detsum[d][2]/catsum[c][2])
                    catsum[c][3].append([d,detsum[d]])
                    
        print(catsum)
        
        self.tsto=Gtk.ListStore(str,str,str,str)
        self.tsto.append(['','科目','金額','%'])
        self.tsto.append(['收入','','',''])
        self.tsto.append(['','捐贈收入','{:,d}'.format(-detsum['捐贈收入'][2]),detsum['捐贈收入'][4]])
        self.tsto.append(['','利息收入','{:,d}'.format(-detsum['利息收入'][2]),detsum['利息收入'][4]])
        self.tsto.append(['','政府補助收入','{:,d}'.format(-detsum['政府補助收入'][2]),detsum['政府補助收入'][4]])
        self.tsto.append(['','其他收入','{:,d}'.format(-detsum['其他收入'][2]),detsum['其他收入'][4]])
        self.tsto.append(['','收入合計','{:,d}'.format(-catsum['收入'][2]),''])
        income=-catsum['收入'][2]
        spending=catsum['支出'][2]
        self.tsto.append(['支出','','',''])
        self.tsto.append(['','業務支出','{:,d}'.format(detsum['業務支出'][2]),detsum['業務支出'][4]])
        self.tsto.append(['','行政管理支出','{:,d}'.format(detsum['行政管理支出'][2]),detsum['行政管理支出'][4]])
        self.tsto.append(['','捐助支出','{:,d}'.format(detsum['捐助支出'][2]),detsum['捐助支出'][4]])
        self.tsto.append(['','支出合計','{:,d}'.format(catsum['支出'][2]),''])
        self.tsto.append(['本期餘絀','','{:,d}'.format(income-spending),''])
            
        self.tsto.append(['支出占比','','{:.2%}'.format(spending/income),''])
        self.tsto.append(['足額支出','','{:,.2f}'.format(income*0.6),''])
        if income*0.6>spending:
            self.tsto.append(['尚須支出','','{:,.2f}'.format(income*0.6-spending),''])

        vrender=Gtk.CellRendererText()
        vb = Gtk.TreeView(model=self.tsto)
        col1=Gtk.TreeViewColumn('',cell_renderer=vrender,text=0)
        col2=Gtk.TreeViewColumn('',cell_renderer=vrender,text=1)
        col3=Gtk.TreeViewColumn('',cell_renderer=vrender,text=2)
        col4=Gtk.TreeViewColumn('',cell_renderer=vrender,text=3)
        vb.append_column(col1)        
        vb.append_column(col2)        
        vb.append_column(col3)        
        vb.append_column(col4)        
        rep1Win=ReportWindow()  
        rep1Win.set_title("收支餘絀表")
        rep1Win.set_text(vb)
        rep1Win.set_pdffun(self.printfun)
        rep1Win.show_all()

    def printfun(self):
        fname=self.yent.get_text()+"年"+"收支餘絀表"
        c=canvas.Canvas(fname+".pdf")
        c.setFont('新細明體',16)
        c.drawString(85,800,fname)
        ypos = 750
        c.setFont('新細明體',14)

        n=self.tsto.get_iter_first()
        while n != None:
            for i in range(4):
                 t=self.tsto.get_value(n,i)
                 c.drawString(20+100*i,ypos,t)
            n=self.tsto.iter_next(n)
            ypos=ypos-18
            
        c.showPage()
        c.save()
        donemsg(fname+" 印出")

            
    def loadyear(self,widget):
        self.store.clear()
        try:
             year=int(self.yent.get_text())
        except:
             year=0
        if year in voucher.voucherbook.years:
            crsum=0
            dbsum=0
            self.asum=account_sums(year)
            for r in sorted(self.asum):
                crsum=crsum+self.asum[r][1]
                dbsum=dbsum+self.asum[r][2]
                self.store.append([r,self.asum[r][0][0],str(self.asum[r][1]),str(self.asum[r][2])])
            print(crsum,dbsum)
        self.show_all()

class ReportWindow(Gtk.Window):
    def __init__(self):
        super().__init__(title='報表')
        self.set_size_request(500,600)
        box=Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        hbox=Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.add(box)
        box.pack_start(hbox,False,False,0)
        repbut=Gtk.Button("印表")
        hbox.pack_start(repbut,False,False,0)
        repbut.connect("clicked",self.printreport)
        self.scw=Gtk.ScrolledWindow()
        box.pack_start(self.scw,True,True,0)

    def set_text(self,t):
        self.tw=t
        self.scw.add(self.tw)

    def set_pdffun(self,fun):
        self.printfun=fun

    def printreport(self,widget):
        self.printfun()

class AccountWin(Gtk.Window):
    def __init__(self):
        super().__init__(title="科目")
        self.set_size_request(800,500)
        box=Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.add(box)
        grd=Gtk.Grid()
        grd.set_column_spacing(5)
        grd.set_row_spacing(5)
        scw=Gtk.ScrolledWindow()
        box.pack_start(grd,False,False,0)
        ct=Gtk.Label(label="借方總額")
        dt=Gtk.Label(label="貸方總額")
        dct=Gtk.Label(label="借貸")
        st=Gtk.Label(label="總額")
        self.cs=Gtk.Label("0")
        self.ds=Gtk.Label("0")
        self.dcs=Gtk.Label("")
        self.ss=Gtk.Label("0")
        grd.attach(ct,0,0,1,1)
        grd.attach(dt,1,0,1,1)
        grd.attach(dct,2,0,1,1)
        grd.attach(st,3,0,1,1)
        grd.attach(self.cs,0,1,1,1)
        grd.attach(self.ds,1,1,1,1)
        grd.attach(self.dcs,2,1,1,1)
        grd.attach(self.ss,3,1,1,1)
        reportbut=Gtk.Button("報表")
        grd.attach(reportbut,4,0,1,1)
        reportbut.connect("clicked",self.printrep)
        csvbut=Gtk.Button("CSV 報表")
        grd.attach(csvbut,5,0,1,1)
        csvbut.connect("clicked",self.csvrep)
        box.pack_start(scw,True,True,0)
        self.store=Gtk.ListStore(str,str,str,str,int,int,str,int)
        vb=Gtk.TreeView(model=self.store)
        scw.add(vb)
        vrender=Gtk.CellRendererText()
        dtcol=Gtk.TreeViewColumn("日期",cell_renderer=vrender,text=0)
        vb.append_column(dtcol)
        dorccol=Gtk.TreeViewColumn("借貸",cell_renderer=vrender,text=1)
        vb.append_column(dorccol)
        idcol=Gtk.TreeViewColumn("編號",cell_renderer=vrender,text=2)
        vb.append_column(idcol)
        rmcol=Gtk.TreeViewColumn("摘要",cell_renderer=vrender,text=3)
        vb.append_column(rmcol)
        crcol=Gtk.TreeViewColumn("借方金額",cell_renderer=vrender,text=4)
        vb.append_column(crcol)
        dbcol=Gtk.TreeViewColumn("貸方金額",cell_renderer=vrender,text=5)
        vb.append_column(dbcol)
        vdcol=Gtk.TreeViewColumn("發票日期",cell_renderer=vrender,text=6)
        vb.append_column(vdcol)
        sumcol=Gtk.TreeViewColumn("餘額",cell_renderer=vrender,text=7)
        vb.append_column(sumcol)

    def set_content(self,rlist):
        sum=0
        for r in rlist:
            sum=sum+r.credit-r.debit
            self.store.append([r.date,r.dorc,r.index,r.remark,r.credit,r.debit,r.vdate,sum])

    def set_sum(self,cr,db,dc,sumx):
        self.cs.set_text(str(cr))
        self.ds.set_text(str(db))
        self.dcs.set_text(dc)
        self.ss.set_text(str(sumx))

    def printrep(self,button):
        tit=self.get_title()
        c=open_pdf(tit)
        line=0
        crsum=0
        debsum=0
        sumprt=False
        n=self.store.get_iter_first()
        while n!= None:
            if line==0:
                pdf_prttitle(c,tit)
            pdf_prtln(c,line,self.store.get(n,2,0)+(tit,"")+self.store.get(n,3,4,5,6))
            crsum = crsum+self.store.get_value(n,4)
            debsum = debsum + self.store.get_value(n,5)
            n = self.store.iter_next(n)
            line=line+1
            if line>34:
                if n==None:
                    sumprt=True
                    c.setFont('新細明體',16)
                    c.drawString(100,30,"借方金額:{:,d} 貸方金額:{:,d}  總數:{:,d}".format(crsum,debsum,crsum-debsum))
                c.showPage()
                line=0
        if not sumprt:
            c.setFont('新細明體',16)
            c.drawString(100,30,"借方金額:{:,d} 貸方金額:{:,d}  總數:{:,d}".format(crsum,debsum,crsum-debsum))
        c.showPage()
        c.save()
        donemsg("科目報表印出")

    def csvrep(self,widget):
        tit=self.get_title()
        with open(tit+".csv",'w',newline='') as csvfile:
            writer=csv.writer(csvfile,delimiter=',')
            writer.writerow(["日期","借貸","編號","說明","借方","貸方","發票日期","餘額"])
            n = self.store.get_iter_first()
            while n!=None:
                writer.writerow( self.store.get(n,0,1,2,3,4,5,6,7))
                n = self.store.iter_next(n)
            csvfile.close()    
        donemsg("CSV 輸出")

        
class DetailedWin(Gtk.Window):
    def __init__(self,tit):
        self.tit=tit
        super().__init__(title=tit+"收支細目表")
        box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)
        box1=Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        box2=Gtk.Box(orientation=Gtk.Orientation.VERTICAL)

        title1=Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)
        title2=Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)

        self.lab1=Gtk.Label()
        self.lab2=Gtk.Label()

        repb1=Gtk.Button(label="報表")
        repb2=Gtk.Button(label="報表")
        repb1.connect("clicked",self.income_report)
        repb2.connect("clicked",self.spending_report)

        scwin1 = Gtk.ScrolledWindow()
        scwin2 = Gtk.ScrolledWindow()
        
        scwin1.set_size_request(900,600)
        scwin2.set_size_request(900,600)
        self.insto = Gtk.ListStore(str,str,str,str,str,str,int,int,str)
        self.outsto = Gtk.ListStore(str,str,str,str,str,str,int,int,str)
        self.inbox = Gtk.TreeView(model=self.insto)
        vrender=Gtk.CellRendererText()
        in_indexcol=Gtk.TreeViewColumn("編號",cell_renderer=vrender,text=0)
        self.inbox.append_column(in_indexcol)
        in_acol=Gtk.TreeViewColumn("日期",cell_renderer=vrender,text=1)
        self.inbox.append_column(in_acol)
        in_dorccol=Gtk.TreeViewColumn("借貸",cell_renderer=vrender,text=2)
        self.inbox.append_column(in_dorccol)
        in_acol=Gtk.TreeViewColumn("科目",cell_renderer=vrender,text=3)
        self.inbox.append_column(in_acol)
        in_ancol=Gtk.TreeViewColumn("科目名稱",cell_renderer=vrender,text=4)
        self.inbox.append_column(in_ancol)
        in_rmcol=Gtk.TreeViewColumn("摘要",cell_renderer=vrender,text=5)
        self.inbox.append_column(in_rmcol)
        in_crcol=Gtk.TreeViewColumn("借方金額",cell_renderer=vrender,text=6)
        self.inbox.append_column(in_crcol)
        in_dbcol=Gtk.TreeViewColumn("貸方金額",cell_renderer=vrender,text=7)
        self.inbox.append_column(in_dbcol)
        in_vdcol=Gtk.TreeViewColumn("發票日期",cell_renderer=vrender,text=8)
        self.inbox.append_column(in_vdcol)
        
        self.outbox = Gtk.TreeView(model=self.outsto)
        out_indexcol=Gtk.TreeViewColumn("編號",cell_renderer=vrender,text=0)
        self.outbox.append_column(out_indexcol)
        out_acol=Gtk.TreeViewColumn("日期",cell_renderer=vrender,text=1)
        self.outbox.append_column(out_acol)
        out_dorccol=Gtk.TreeViewColumn("借貸",cell_renderer=vrender,text=2)
        self.outbox.append_column(out_dorccol)
        out_acol=Gtk.TreeViewColumn("科目",cell_renderer=vrender,text=3)
        self.outbox.append_column(out_acol)
        out_ancol=Gtk.TreeViewColumn("科目名稱",cell_renderer=vrender,text=4)
        self.outbox.append_column(out_ancol)
        out_rmcol=Gtk.TreeViewColumn("摘要",cell_renderer=vrender,text=5)
        self.outbox.append_column(out_rmcol)
        out_crcol=Gtk.TreeViewColumn("借方金額",cell_renderer=vrender,text=6)
        self.outbox.append_column(out_crcol)
        out_dbcol=Gtk.TreeViewColumn("貸方金額",cell_renderer=vrender,text=7)
        self.outbox.append_column(out_dbcol)
        out_vdcol=Gtk.TreeViewColumn("發票日期",cell_renderer=vrender,text=8)
        self.outbox.append_column(out_vdcol)

        self.add(box)
        box.pack_start(box1,False,False,0)
        box.pack_start(box2,False,False,0)
        box1.pack_start(title1,False,False,0)
        box2.pack_start(title2,False,False,0)
        title1.pack_start(self.lab1,False,False,0)
        title2.pack_start(self.lab2,False,False,0)
        title1.pack_start(repb1,False,False,0)
        title2.pack_start(repb2,False,False,0)
        box1.pack_start(scwin1,False,False,0)
        box2.pack_start(scwin2,False,False,0)
        scwin1.add(self.inbox)
        scwin2.add(self.outbox)


    def income_report(self,widget):
        c=open_pdf(self.tit+"收入")
        line=0
        crsum=0
        debsum=0
        sumprt=False
        n = self.insto.get_iter_first()
        while n !=  None:
            if line==0:
                pdf_prttitle(c,self.tit+"收入")
            pdf_prtln(c,line,self.insto.get(n,0,1,3,4,5,6,7,8))
            crsum = crsum+self.insto.get_value(n,6)
            debsum = debsum+self.insto.get_value(n,7)
            n=self.insto.iter_next(n)
            line=line+1
            if line>34:
                if n==None:
                    sumprt=True
                    c.setFont('新細明體',16)
                    c.drawString(100,30,"借方金額:{:,d} 貸方金額:{:,d}  總數:{:,d}".format(crsum,debsum,crsum-debsum))
                c.showPage()
                line=0
        if not sumprt:
            c.setFont('新細明體',16)
            c.drawString(100,30,"借方金額:{:,d} 貸方金額:{:,d}  總數:{:,d}".format(crsum,debsum,crsum-debsum))
        c.showPage()
        c.save()
        donemsg("收入報表印出")

        
    def spending_report(self,widget):
        c=open_pdf(self.tit+"支出")
        line=0
        crsum=0
        debsum=0
        sumprt=False
        n = self.outsto.get_iter_first()
        while n !=  None:
            if line==0:
                pdf_prttitle(c,self.tit+"支出")
            pdf_prtln(c,line,self.outsto.get(n,0,1,3,4,5,6,7,8))
            crsum = crsum+self.outsto.get_value(n,6)
            debsum = debsum+self.outsto.get_value(n,7)
            n=self.outsto.iter_next(n)
            line=line+1
            if line>34:
                if n==None:
                    sumprt=True
                    c.setFont('新細明體',16)
                    c.drawString(100,30,"借方金額:{:,d} 貸方金額:{:,d}  總數:{:,d}".format(crsum,debsum,crsum-debsum))
                c.showPage()
                line=0
        if not sumprt:
            c.setFont('新細明體',16)
            c.drawString(100,30,"借方金額:{:,d} 貸方金額:{:,d}  總數:{:,d}".format(crsum,debsum,debsum-crsum))
        c.showPage()
        c.save()
        donemsg("支出報表印出")
        

    def loaddata(self,inrep,outrep):
        icredit=0
        idebit=0
        ocredit=0
        odebit=0
        for i in inrep:
            self.insto.append([i.index,i.date,i.dorc,i.code,i.account,i.remark,i.debit,i.credit,i.vdate])
            icredit=icredit+i.credit
            idebit=idebit+i.debit
            self.lab1.set_text("收入 credit={:,d} debit={:,d} 收入總數={:,d}".format(icredit,idebit,idebit-icredit))
        for i in outrep:
            self.outsto.append([i.index,i.date,i.dorc,i.code,i.account,i.remark,i.debit,i.credit,i.vdate])
            ocredit=ocredit+i.credit
            odebit=odebit+i.debit
            self.lab2.set_text("支出  credit={:,d} debit={:,d} 支出總數={:,d}".format(ocredit,odebit,ocredit-odebit))

def donemsg(message):
    __main__.win.log(message)
