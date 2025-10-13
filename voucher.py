import gi
gi.require_version("Gtk","3.0")
from gi.repository import Gtk
import json
import csv
from account import accounts
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
import __main__


class VoucherRecord:
    def __init__(self):
        self.date='0/0/0'
        self.dorc=''
        self.index=''
        self.code=''
        self.account=''
        self.remark=''
        self.debit=0
        self.credit=0
        self.vdate=''
        
    def month(self):
        s=self.date.split("/")
        return int(s[1])

    def year(self):
        s.self.date.split("/")
        return int(s[0])

    def datenum(self):
        s=self.date.split("/")
        return [int(s[0]),int(s[1]),int(s[2])]

class Voucher:
    def __init__(self):
        self.index=''
        self.date=''
        self.year=0
        self.month=0
        self.day=0
        self.records=[]
        
    
    
class VoucherBook:
    def __init__(self):
        self.years=set()
        self.yearbook={}

    def importcsv(self,filename):
        yearvrec=[]
        with open(filename,'r',encoding='utf-8') as csvfile:
            freader = csv.reader(csvfile)
            count=0
            for row in freader:
                if count==0:
                    titlelist=row
                else:
                    v = VoucherRecord()
                    v.date=row[0]
                    v.dorc=row[1]
                    v.index=row[2]
                    v.code=row[3]
                    v.account=row[4]
                    v.remark=row[5]
                    v.credit=int('0'+row[6])
                    v.debit=int('0'+row[7])
                    v.vdate=row[8]
                    yearvrec.append(v)
                count=count+1
            print("Total records = ",count)

        vbook=[]
        lastr=None
        curv = Voucher()
        for r in yearvrec:
            if lastr!=None and r.index!=lastr.index:
                curv.index=lastr.index
                curv.date=lastr.date
                [curv.year,curv.month,curv.day]=lastr.datenum()
                print(curv.index,len(curv.records))
                vbook.append(curv)
                curv=Voucher()
            lastr=r
            curv.records.append(r)
        self.insertbook(vbook)
        __main__.win.log(self.years)

    def insertbook(self,vbook):
        for r in vbook:
            if not (r.year in self.years):
                self.years.add(r.year)
                self.yearbook[r.year]=dict()
            self.yearbook[r.year][r.index]=r

    def dumpbook(self):
        a=[]
        for i in self.years:
            a.append(i)
        yb=[]
        for y in self.yearbook:
            vl = self.yearbook[y]
            r=[]
            for v in sorted(vl):
                vrs=[]
                for vr in vl[v].records:
                    vrs.append([vr.date,vr.dorc,vr.index,vr.code,vr.account,vr.remark,vr.debit,vr.credit,vr.vdate])
                r.append( [vl[v].index,vl[v].date,vrs])
            yb.append([y,r])
        x = [a, yb]
        f = open('voucher.db','w')
        json.dump(x,f)
        __main__.win.log("Database dumped!\n")
        f.close()

    def saveyearbook(self,year):
        a=[]
        vl=self.yearbook[year]
        r=[]
        for v in sorted(vl):
            vrs = []
            for vr in vl[v].records:
                vrs.append([vr.date,vr.dorc,vr.index,vr.code,vr.account,vr.remark,vr.debit,vr.credit,vr.vdate])
            r.append( [vl[v].index,vl[v].date,vrs])
        f = open('voucher'+str(year)+'.db','w')
        json.dump(r,f)
        __main__.win.log("Database of year"+str(year)+" dumped\n")
        with open('voucher'+str(year)+".csv",'w',newline='',encoding='utf-8') as csvfile:
            writer=csv.writer(csvfile,delimiter=',')
            writer.writerow(["日期","借貸","編號","科目編號","科目","說明","借方","貸方","發票日期"])
            for rx in r :
                for rr in rx[2]:
                    writer.writerow( [rr[0],rr[1],rr[2],rr[3],rr[4],rr[5],str(rr[7]), str(rr[6]),rr[8]])
            csvfile.close()    
        __main__.win.log(" CSV 輸出\n")
        f.close()

    def loadyearbook(self,year):
        f=open('voucher'+str(year)+'.db','r')
        x=json.load(f)
        print(year)
        vbb={}
        for y in x:
            vc=Voucher()
            vc.index=y[0]
            vc.date=y[1]
            vc.records=[]
            for i in y[2]:
                vcn = VoucherRecord()
                vcn.date=i[0]
                vcn.dorc=i[1]
                vcn.index=i[2]
                vcn.code=i[3]
                vcn.account=i[4]
                vcn.remark=i[5]
                vcn.debit=i[6]
                vcn.credit=i[7]
                vcn.vdate=i[8]
                vc.records.append(vcn)
            vbb[vc.index]=vc
            print(vc)
        self.yearbook[year]=vbb
        self.years.add(year)

    def loadbook(self):
        f = open('voucher.db','r')
        x=json.load(f)
        print("Database loaded")
        self.years=set()
        self.yearbook={}
        for i in x[0]:
             self.years.add(i)
        for yb in x[1]:
            year=yb[0]
            yearb=yb[1]
            vbb = {}
            for y in yearb:
                vc = Voucher()
                vc.index=y[0]
                vc.date=y[1]
                vc.records=[]
                for i in y[2]:
                    vcn = VoucherRecord()
                    vcn.date=i[0]
                    vcn.dorc=i[1]
                    vcn.index=i[2]
                    vcn.code=i[3]
                    vcn.account=i[4]
                    vcn.remark=i[5]
                    vcn.debit=i[6]
                    vcn.credit=i[7]
                    vcn.vdate=i[8]
                    vc.records.append(vcn)
                vbb[vc.index]=vc
            self.yearbook[year]=vbb    
        f.close()

class VoucherWin(Gtk.Window):
    def __init__(self):
        Gtk.Window.__init__(self,title="傳票管理")

        vbox=Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.add(vbox)
        hbox1=Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)
        vbox.pack_start(hbox1,False,False,0)
        addb=Gtk.Button(label="輸入傳票")
        addb.connect("clicked",self.add_voucher)
        hbox1.pack_start(addb,False,False,0)
        edtb=Gtk.Button(label="編輯傳票")
        edtb.connect("clicked",self.edit_voucher)
        hbox1.pack_start(edtb,False,False,0)
        delb=Gtk.Button(label="刪除傳票")
        hbox1.pack_start(delb,False,False,0)
        delb.connect("clicked",self.delete_voucher)
        savb=Gtk.Button(label="儲存傳票")
        hbox1.pack_start(savb,False,False,0)
        impb=Gtk.Button(label="導入傳票")
        hbox1.pack_start(impb,False,False,0)
        impb.connect("clicked",self.importcsv)
        savb.connect("clicked",self.savedb)
        savexb=Gtk.Button(label="儲存年度")
        hbox1.pack_start(savexb,False,False,0)
        savexb.connect("clicked",self.save_yeardb)
        self.loadent=Gtk.Entry()
        hbox1.pack_start(self.loadent,False,False,0)
        loadxb=Gtk.Button(label="載入年度")
        loadxb.connect("clicked",self.load_yeardb)
        hbox1.pack_start(loadxb,False,False,0)

        self.hbox2=Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)
        vbox.pack_start(self.hbox2,False,False,0)
   
        self.hbox3=Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)
        vbox.pack_start(self.hbox3,True,True,0)

        self.yearsto = Gtk.ListStore(int)
        for i in voucherbook.years:
            self.yearsto.append([i])
        self.yearbox=Gtk.TreeView(model=self.yearsto)
        self.yearbox.connect("row-activated",self.year_clicked)
        trender = Gtk.CellRendererText()
        yearcol=Gtk.TreeViewColumn("年度",cell_renderer=trender,text=0)
        self.yearbox.append_column(yearcol)
        yearwin=Gtk.ScrolledWindow()
        yearwin.add(self.yearbox)
        yearframe=Gtk.Frame()
        yearframe.add(yearwin)
        self.hbox3.pack_start(yearframe,False,False,0)

        self.idsto = Gtk.ListStore(str,str)

        self.idbox=Gtk.TreeView(model=self.idsto)
        irender = Gtk.CellRendererText()
        dcol=Gtk.TreeViewColumn("日期",cell_renderer=irender,text=0)
        self.idbox.append_column(dcol)
        idcol = Gtk.TreeViewColumn("傳票編號",cell_renderer=irender,text=1)
        self.idbox.append_column(idcol)
        self.idbox.connect("row-activated",self.id_clicked)
        idwin=Gtk.ScrolledWindow()
        idwin.set_size_request(180,300)
        idwin.add(self.idbox)
        idframe=Gtk.Frame()
        idframe.add(idwin)
        self.hbox3.pack_start(idframe,False,False,0)
        
        self.vsto = Gtk.ListStore(str,str,str,str,int,int,str)
        vrender = Gtk.CellRendererText()
        self.vobox = Gtk.TreeView(model=self.vsto)
        dorccol=Gtk.TreeViewColumn("借貸",cell_renderer=vrender,text=0)
        self.vobox.append_column(dorccol)
        acol=Gtk.TreeViewColumn("科目",cell_renderer=vrender,text=1)
        self.vobox.append_column(acol)
        ancol=Gtk.TreeViewColumn("科目名稱",cell_renderer=vrender,text=2)
        self.vobox.append_column(ancol)
        rmcol=Gtk.TreeViewColumn("摘要",cell_renderer=vrender,text=3)
        self.vobox.append_column(rmcol)
        crcol=Gtk.TreeViewColumn("借方金額",cell_renderer=vrender,text=4)
        self.vobox.append_column(crcol)
        dbcol=Gtk.TreeViewColumn("貸方金額",cell_renderer=vrender,text=5)
        self.vobox.append_column(dbcol)
        vdcol=Gtk.TreeViewColumn("發票日期",cell_renderer=vrender,text=6)
        self.vobox.append_column(vdcol)
        vowin=Gtk.ScrolledWindow()
        vowin.set_size_request(650,500)
        vowin.add(self.vobox)
        voframe=Gtk.Frame()
        voframe.add(vowin)
        self.hbox3.pack_start(voframe,False,False,0)

        sumbox=Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        sumbox.set_border_width(2)
        sumframe=Gtk.Frame()
        sumframe.add(sumbox)
        self.hbox3.pack_start(sumframe,False,False,0)
        crlab=Gtk.Label("借方總額")
        self.crnum=Gtk.Label()
        dblab=Gtk.Label("貸方總額")
        self.dbnum=Gtk.Label()
        sumbox.pack_start(crlab,False,False,0)
        sumbox.pack_start(self.crnum,False,False,0)
        sumbox.pack_start(dblab,False,False,0)
        sumbox.pack_start(self.dbnum,False,False,0)
    
    def save_yeardb(self,widget):
        curpath=self.yearbox.get_cursor()
        print(curpath[0])
        treeiter=self.yearsto.get_iter(curpath[0])
        y = self.yearsto.get_value(treeiter,0)
        voucherbook.saveyearbook(y)
        print('voucher'+str(y)+'.db saved')
        

    def load_yeardb(self,widget):
        year=int(self.loadent.get_text())
        voucherbook.loadyearbook(year)
        print('voucher'+str(year)+'.db loaded')
    
        
    def savedb(self,widget):
        voucherbook.dumpbook()
        self.but = Gtk.Button("完成")
        self.but.connect("clicked",self.savedbdone)
        self.hbox2.pack_start(self.but,False,False,0)

    def savedbdone(self,widget):
        self.remove(self.but)

    def show_voucher_list(self):
        for y in voucherbook.years:
            self.yearsto.append([y])
        self.show_all()

    def year_clicked(self,tree,path,col):
        treeiter=self.yearsto.get_iter(path)
        self.selectedyear=self.yearsto.get_value(treeiter,0)
        self.idsto.clear()
        idbk = voucherbook.yearbook[self.selectedyear]
        for i in idbk:
            self.idsto.append([idbk[i].date,idbk[i].index])
        self.show_all()

    def id_clicked(self,tree,path,col):
        treeiter=self.idsto.get_iter(path)
        id=self.idsto.get_value(treeiter,1)
        print (id)
        recs=voucherbook.yearbook[self.selectedyear][id].records
        self.vsto.clear()
        cr = 0
        db = 0
        for i in recs:
            self.vsto.append([i.dorc,i.code,i.account,i.remark,i.credit,i.debit,i.vdate])
            cr = cr+i.credit
            db = db+i.debit
        self.crnum.set_text(str(cr))
        self.dbnum.set_text(str(db))
        self.show_all()
        
    def importcsv(self,widget):
        print("ImportCSV")
        self.but=Gtk.Button("導入")
        self.hbox2.pack_start(self.but,False,False,0)
        self.lab=Gtk.Label("檔案名稱:")
        self.hbox2.pack_start(self.lab,False,False,0)
        self.importent=Gtk.Entry()
        self.hbox2.pack_start(self.importent,False,False,0)
        self.but.connect("clicked",self.importing)
        self.show_all()

    def importing(self,widget):
        fname=self.importent.get_text()
        voucherbook.importcsv(fname)
        self.show_voucher_list()
        self.hbox2.remove(self.lab)
        self.hbox2.remove(self.but)
        self.hbox2.remove(self.importent)

    def add_voucher(self,widget):
        diag=AddVoucher(self)
        resp=diag.run()
        year=diag.yent.get_text()
        month=diag.ment.get_text()
        day=diag.dent.get_text()
        v=Voucher()
        v.year=int(diag.yent.get_text())
        v.month=int(diag.ment.get_text())
        v.day=int(diag.dent.get_text())
        v.index=diag.indexlab.get_text()
        v.date="{:d}/{:02d}/{:02d}".format(v.year,v.month,v.day)
        v.records=[]
        diag.destroy()
        if (resp==Gtk.ResponseType.OK) and (v.index[0]=='T'):
            diag=EditVoucher(self,v)
            resp=diag.run()
            if (resp==Gtk.ResponseType.OK):
                v.records=self.make_voucher(v.date,v.index,diag.vrsto)
                if not(v.year in voucherbook.years):
                    voucherbook.years.add(v.year)
                    self.yearsto.append([v.year])
                    voucherbook.yearbook[v.year]={}
                voucherbook.yearbook[v.year][v.index]=v
            diag.destroy()

    def delete_voucher(self,widget):
        path=self.idbox.get_cursor()[0]
        if path!=None:
            iter=self.idsto.get_iter(path)
            index=self.idsto.get_value(iter,1)
            self.idsto.remove(iter)
            path=self.yearbox.get_cursor()[0]
            itery=self.yearsto.get_iter(path)
            y=self.yearsto.get_value(itery,0)
            del(voucherbook.yearbook[y][index])
            if voucherbook.yearbook[y]=={}:
                self.yearsto.remove(itery)
                voucherbook.years.remove(y)

    def make_voucher(self,d,index,sto):
        recs=[]
        iter=sto.get_iter_first()
        while iter!=None:
            a=VoucherRecord()
            a.date=d
            a.index=index
            a.dorc=sto.get_value(iter,0)
            a.code=sto.get_value(iter,1)
            a.account=sto.get_value(iter,2)
            a.remark=sto.get_value(iter,3)
            a.credit=sto.get_value(iter,4)
            a.debit=sto.get_value(iter,5)
            a.vdate=sto.get_value(iter,6)
            recs.append(a)
            iter=sto.iter_next(iter)
        return recs
        
    def edit_voucher(self,widget):
        path=self.idbox.get_cursor()[0]
        if path!=None:
            iter=self.idsto.get_iter(path)
            id=self.idsto.get_value(iter,1)
            path=self.yearbox.get_cursor()[0]
            iter=self.yearsto.get_iter(path)
            year=self.yearsto.get_value(iter,0)
            v=voucherbook.yearbook[year][id]
            diag=EditVoucher(self,v)
            resp=diag.run()
            if (resp==Gtk.ResponseType.OK):
                voucherbook.yearbook[year][id].records=self.make_voucher(v.date,v.index,diag.vrsto)
            diag.destroy()

monthdays={1:31,2:29,3:31,4:30,5:31,6:30,7:31,8:31,9:30,10:31,11:30,12:31}

class AddVoucher(Gtk.Dialog):
    def __init__(self,parent):
        super().__init__(title="輸入傳票",transient_for=parent,flags=0)
        self.add_buttons(Gtk.STOCK_CANCEL,Gtk.ResponseType.CANCEL,Gtk.STOCK_OK,Gtk.ResponseType.OK)
        self.set_default_size(200,150)
        box=self.get_content_area()
        fgrid=Gtk.Grid()
        box.add(fgrid)
        tlab=Gtk.Label("傳票日期")
        fgrid.attach(tlab,0,0,2,1)
        ylab=Gtk.Label("年")
        fgrid.attach_next_to(ylab,tlab,Gtk.PositionType.BOTTOM,1,1)
        self.yent=Gtk.Entry()
        fgrid.attach_next_to(self.yent,ylab,Gtk.PositionType.RIGHT,1,1)
        mlab=Gtk.Label("月")
        fgrid.attach_next_to(mlab,ylab,Gtk.PositionType.BOTTOM,1,1)
        self.ment=Gtk.Entry()
        fgrid.attach_next_to(self.ment,mlab,Gtk.PositionType.RIGHT,1,1)
        dlab=Gtk.Label("日")
        fgrid.attach_next_to(dlab,mlab,Gtk.PositionType.BOTTOM,1,1)
        self.dent=Gtk.Entry()
        fgrid.attach_next_to(self.dent,dlab,Gtk.PositionType.RIGHT,1,1)
        getindexb=Gtk.Button("產生傳票編號")
        fgrid.attach_next_to(getindexb,dlab,Gtk.PositionType.BOTTOM,2,1)
        self.indexlab=Gtk.Label()
        fgrid.attach_next_to(self.indexlab,getindexb,Gtk.PositionType.BOTTOM,2,1)
        getindexb.connect("clicked",self.genindex)
        self.show_all()
        
    def genindex(self,widget):
        try:
            y = int(self.yent.get_text())
            m = int(self.ment.get_text())
            d = int(self.dent.get_text())
            if y<100:
                raise ValueError
            if m<1 or m>12:
                raise ValueError
            if d<1 or d>monthdays[m]:
                raise ValueError
        except (ValueError,UnboundLocalError):
            self.indexlab.set_text("日期錯誤")
        else:
            self.vdate="{0:03d}/{1:02d}/{2:02d}".format(y,m,d)
            if  y in voucherbook.years:
                ind = 1
                trindx="T{0:02d}{1:02d}001".format(m,d)
                while trindx in voucherbook.yearbook[y]:
                    ind=ind+1
                    trindx = "T{0:02d}{1:02d}{2:03d}".format(m,d,ind)
                else:
                    self.indexlab.set_text(trindx)
            else:
                self.indexlab.set_text("T{0:02d}{1:02d}001".format(m,d))

def pdate(sdate):
    slist=sdate.split("/")
    return ("中華民國"+slist[0]+"年"+slist[1]+"月"+slist[2]+"日")


class EditVoucher(Gtk.Dialog):
    def __init__(self,parent,voucher):
        super().__init__(title="編輯傳票",transient_for=parent,flags=0)
        self.year=voucher.year
        self.index=voucher.index
        self.date=voucher.date
        self.add_buttons(Gtk.STOCK_CANCEL,Gtk.ResponseType.CANCEL,Gtk.STOCK_OK,Gtk.ResponseType.OK)
        self.set_default_size(800,150)
        box=self.get_content_area()
        hbox=Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        box.add(hbox)
        fbox=Gtk.Box()
        fgrid2=Gtk.Grid()
        
        self.totalcr=0
        self.totaldb=0
        self.vrsto=Gtk.ListStore(str,str,str,str,int,int,str)
        for r in voucher.records:
            self.vrsto.append([r.dorc,r.code,r.account,r.remark,r.credit,r.debit,r.vdate])
            self.totalcr=self.totalcr+r.credit
            self.totaldb=self.totaldb+r.debit
        self.sumlabel=Gtk.Label()
        self.update_sum()

        self.vrbox=Gtk.TreeView(model=self.vrsto)
        vowin=Gtk.ScrolledWindow()
        vowin.add(self.vrbox)
        vframe=Gtk.Frame()
        vframe.add(vowin)
        vframe.set_size_request(300,300)
       
        sep = Gtk.Separator.new(Gtk.Orientation.HORIZONTAL)
        hbox.pack_start(fbox,False,False,0)
        hbox.pack_start(fgrid2,False,False,0)
        hbox.pack_start(sep,False,False,0)
        hbox.pack_start(vframe,True,True,0)

        title=Gtk.Label("編號: "+self.index+" 日期: "+voucher.date)
        
        tbt1=Gtk.Button("加入")
        tbt1.connect("clicked",self.add_record)
        tbt2=Gtk.Button("修改")
        tbt2.connect("clicked",self.edit_record)
        tbt3=Gtk.Button("刪除")
        tbt3.connect("clicked",self.delete_record)
        tbt4=Gtk.Button("交換借貸")
        tbt4.connect("clicked",self.swap_numbers)
        tbt5=Gtk.Button("列印")
        tbt5.connect("clicked",self.print_voucher)
        fbox.pack_start(title,False,False,0)
        fbox.pack_start(self.sumlabel,False,False,0)
        fbox.pack_end(tbt5,False,False,0)
        fbox.pack_end(tbt4,False,False,0)
        fbox.pack_end(tbt3,False,False,0)
        fbox.pack_end(tbt2,False,False,0)
        fbox.pack_end(tbt1,False,False,0)

        self.dorcbox=Gtk.ComboBoxText()
        self.dorcbox.set_entry_text_column(0)
        self.dorcbox.append_text("借")
        self.dorcbox.append_text("貸")
        self.dorcbox.set_active(0)
        self.dorcbox.set_id_column(0)

        self.asto = Gtk.ListStore(str,str)
        for a in accounts.accountbook:
            self.asto.append([a,accounts.accountbook[a][0]])
        self.abox=Gtk.ComboBox.new_with_model_and_entry(self.asto)
        self.abox.set_entry_text_column(0)
        self.abox.get_child().connect("activate",self.abox_changed)
        self.abox.set_id_column(0)
        render1=Gtk.CellRendererText()
        self.abox.pack_start(render1,True)
        render2=Gtk.CellRendererText()
        self.abox.pack_start(render2,True)
#        self.abox.add_attribute(render1,"text",0)
        self.abox.add_attribute(render2,"text",1)
        self.abox.set_active(0)
        
        tlab1=Gtk.Label("借貸")
        self.tlab2=Gtk.Label("科目:")
        tlab4=Gtk.Label("摘要")
        tlab5=Gtk.Label("借方金額")
        tlab6=Gtk.Label("貸方金額")
        tlab7=Gtk.Label("發票日期")
        fgrid2.attach(tlab1,0,0,1,1)
        fgrid2.attach(self.tlab2,1,0,1,1)
        fgrid2.attach(tlab4,2,0,1,1)
        fgrid2.attach(tlab5,3,0,1,1)
        fgrid2.attach(tlab6,4,0,1,1)
        fgrid2.attach(tlab7,5,0,1,1)
        self.remark=Gtk.Entry()
        self.credit=Gtk.Entry()
        self.debit=Gtk.Entry()
        self.vdate=Gtk.Entry()
        fgrid2.attach(self.dorcbox,0,1,1,1)
        fgrid2.attach(self.abox,1,1,1,1)
        fgrid2.attach(self.remark,2,1,1,1)
        fgrid2.attach(self.credit,3,1,1,1)
        fgrid2.attach(self.debit,4,1,1,1)
        fgrid2.attach(self.vdate,5,1,1,1)
        
        vrender=Gtk.CellRendererText()
        dorccol=Gtk.TreeViewColumn("借貸",cell_renderer=vrender,text=0)
        self.vrbox.append_column(dorccol)
        acol=Gtk.TreeViewColumn("科目",cell_renderer=vrender,text=1)
        self.vrbox.append_column(acol)
        ancol=Gtk.TreeViewColumn("科目名稱",cell_renderer=vrender,text=2)
        self.vrbox.append_column(ancol)
        rmcol=Gtk.TreeViewColumn("摘要",cell_renderer=vrender,text=3)
        self.vrbox.append_column(rmcol)
        crcol=Gtk.TreeViewColumn("借方金額",cell_renderer=vrender,text=4)
        self.vrbox.append_column(crcol)
        dbcol=Gtk.TreeViewColumn("貸方金額",cell_renderer=vrender,text=5)
        self.vrbox.append_column(dbcol)
        vdcol=Gtk.TreeViewColumn("發票日期",cell_renderer=vrender,text=6)
        self.vrbox.append_column(vdcol)
        self.vrbox.connect("row-activated",self.row_clicked)
        self.show_all()


    def print_voucher_form(self,c):
        c.setFont('新細明體',16)
#        for s in\
#          ((85,782.5,'財團法人台北市林坤地仁濟文教基金會')):
#            c.drawString(s[0],s[1],s[2])
        c.drawString(85,782.5,'財團法人台北市林坤地仁濟文教基金會')
        c.setFont('新細明體',14)
        c.drawString(226.8,759.8,'轉  帳  傳  票')
        c.setFont('新細明體',12)
        for s in (\
          (430.2,775.5,'總號'),(430.2,761.3,'分號'),
                  (28.3,718.8,'借貸'),(194.3,718.8,'摘  要'),
          (60,718.8,'會計科目'),(350,718.8,'借方金額'),
          (450,718.8,'貸方金額'),(50,477.8,'合  計'),
          (547.1,711.6,'附'),(547.1,697.6,'單'),
          (547.1,683.3,'據'),(547.1,584.1,'張'),
          (28.3,456.5,'核准'),(113.4,456.5,'會計'),
          (198.4,456.5,'稽核'),(283.5,456.5,'登帳'),
          (368.5,456.5,'出納'),(453.5,456.5,'製單')):
            c.drawString(s[0],s[1],s[2])
        c.grid((425.2,467.7,538.6),(787.5,773.3,759.1))
        c.grid((28.3,56.7,189.2,340.2,439.4,538.6),
               (730.8,716.6,702.4,688.3,674.1,659.9,645.7,631.6,617.4,603.2,589.0,574.9,560.7,546.5,532.4,518.2,504,489.8))
        c.grid((28.3,340.2,439.4,538.6),(489.8,475.7))
        c.grid((538.6,566.9),(730.8,475.7))
                                           
        
    def print_voucher(self,widget):
        c=canvas.Canvas(self.index+'.pdf')
        vsize=len(self.vrsto)
        vpages=0
        vdeb=0
        vcrt=0
        iter=self.vrsto.get_iter_first()
        while vsize>0:
            vpages=vpages+1
            self.print_voucher_form(c)
            c.drawString(198.4,733,pdate(self.date))
            c.drawString(469.7,775.5,self.index)
            c.drawString(469.7,761.3,'-'+str(vpages))
            if vsize>16:
                vs=16
            else:
                vs=vsize
            for i in range(vs):
                ypos=718.8-14.1*(i+1)
                c.drawString(34,ypos,self.vrsto.get_value(iter,0))
                c.drawString(62.4,ypos,self.vrsto.get_value(iter,2))
                remark=self.vrsto.get_value(iter,3)
                vdate=self.vrsto.get_value(iter,6)
                if vdate!='':
                    remark=remark+'('+vdate+')'
                c.drawString(195,ypos,remark)
                c.drawString(345.8,ypos,"{:14,}".format(self.vrsto.get_value(iter,4)))
                c.drawString(445,ypos,"{:14,}".format(self.vrsto.get_value(iter,5)))
                vcrt=vcrt+self.vrsto.get_value(iter,4)
                vdeb=vdeb+self.vrsto.get_value(iter,5)
                iter=self.vrsto.iter_next(iter)
            vsize=vsize-16
            if vsize>0:
                c.showPage()
        c.drawString(345.8,477.8,"{:14,}".format(vcrt))
        c.drawString(445,477.8,"{:14,}".format(vdeb))
        print(self.index+' printed')
        c.showPage()
        c.save()
        __main__.win.log(self.index+'印出')

    def abox_changed(self,widget):
        text=self.abox.get_child().get_text()
        iter=self.asto.get_iter_first()
        astr=self.asto.get_value(iter,0)
        while astr<text:
            nextiter=self.asto.iter_next(iter)
            if nextiter!=None:
                iter=nextiter
                astr=self.asto.get_value(iter,0)
        self.abox.set_active_iter(iter)
        name=self.asto.get_value(iter,1)
        self.tlab2.set_text('科目:'+name)


    def row_clicked(self,tree,path,col):
        iter=self.vrsto.get_iter(path)
        r=self.vrsto.get(iter,0,1,2,3,4,5,6)
        self.dorcbox.set_active_id(r[0])
        self.abox.set_active_id(r[1])
        self.remark.set_text(r[3])
        self.credit.set_text(str(r[4]))
        self.debit.set_text(str(r[5]))
        self.vdate.set_text(r[6])

    def add_record(self,widget):
        rc=self.make_record()
        if rc != None:
            self.vrsto.append(rc)    
            self.totalcr=self.totalcr+rc[4]
            self.totaldb=self.totaldb+rc[5]
            self.update_sum()
            

    def edit_record(self,widget):
        path=self.vrbox.get_cursor()[0]
        rc=self.make_record()
        if rc!=None:
            iter=self.vrsto.get_iter(path)
            cr=self.vrsto.get_value(iter,4)
            db=self.vrsto.get_value(iter,5)
            self.totalcr=self.totalcr-cr+rc[4]
            self.totaldb=self.totaldb-db+rc[5]
            self.update_sum()
            self.vrsto.set_row(iter,rc)
            
    def make_record(self):
        dorc=self.dorcbox.get_active_text()
        iter=self.abox.get_active_iter()
        code=self.asto.get_value(iter,0)
        acc=self.asto.get_value(iter,1)
        try:
             cr = self.credit.get_text()
             if cr=='':
                 crd = 0
             else:
                 crd = int(cr)
             db = self.debit.get_text()
             if db =='':
                 dbd = 0
             else:
                 dbd=int(db)
        except ValueError:
            errormsg("數值錯誤")
            return None
        else:
            return [dorc,code,acc,self.remark.get_text(),crd,dbd,self.vdate.get_text()]
    def swap_numbers(self,widget):
        x = self.credit.get_text()
        y = self.debit.get_text()
        self.credit.set_text(y)
        self.debit.set_text(x)

    def delete_record(self,widget):
        path=self.vrbox.get_cursor()[0]
        iter=self.vrsto.get_iter(path)
        self.vrsto.remove(iter)
        
    def errormsg(self,message):
        dialog = Gtk.MessageDialog(
                 transient_for=self,
                 flags=0,
                 message_type=Gtk.MessageType.ERROR,
                 buttons=Gtk.ButtonsType.CANCEL,
                 text= message)
        dialog.run()
        dialog.destroy()

    def update_sum(self):
        self.sumlabel.set_text("借方總額: {0:2d} 貸方總額: {1:2d}".format(self.totalcr,self.totaldb))
    
voucherbook=VoucherBook()
try:
    pdfmetrics.registerFont(TTFont('新細明體','mingliu.ttc'))
except:
    pdfmetrics.registerFont(TTFont('新細明體','uming.ttc'))
