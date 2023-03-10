
import gi
gi.require_version("Gtk","3.0")
from gi.repository import Gtk
import csv
import json
import __main__

class Accounts():
    def __init__(self):
        self.accountbook={}

    def importcsv(self,csvfile):
        with open(csvfile,'r',encoding='utf-8') as c:
            freader = csv.reader(c)
            for row in freader:
                self.accountbook[row[0]]=[row[1],row[2],row[3]]
        __main__.win.log("Accounts imported from "+csvfile)

    def dump_accounts(self):
        a=[]
        for i in sorted(self.accountbook):
            a.append([i, self.accountbook[i]])
            print (i, self.accountbook[i])
        f = open('account.db','w')
        json.dump(a,f)
        __main__.win.log("Accounts dumped")
        f.close()

    def load_accounts(self):
        self.accountbook={}
        f=open('account.db','r')
        x=json.load(f)
        __main__.win.log("Accounts loaded")
        for i in x:
            self.accountbook[i[0]]=i[1]

account_types = ("資產","負債","收入","支出")
cat_types = ("資產","負債","捐贈收入","政府補助收入","其他收入","利息收入","行政管理支出","業務支出")
              
        
class AccountBox(Gtk.Window):
    def __init__(self):
        super().__init__(title="科目")
        box=Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.add(box)
        box1=Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)
        addb=Gtk.Button("增加")
        addb.connect("clicked",self.add_account)
        edtb=Gtk.Button("編輯")
        edtb.connect("clicked",self.edit_account)
        delb=Gtk.Button("刪除")
        savb=Gtk.Button("儲存")
        savb.connect("clicked",self.saveaccount)
        impb=Gtk.Button("導入")
        impb.connect("clicked",self.importaccount)
        box1.pack_start(addb,False,False,0)
        box1.pack_start(edtb,False,False,0)
        box1.pack_start(delb,False,False,0)
        box1.pack_start(savb,False,False,0)
        box1.pack_start(impb,False,False,0)
        self.box2=Gtk.Box()
        box3=Gtk.Box()
        box.pack_start(box1,False,False,0)
        box.pack_start(self.box2,False,False,0)
        box.pack_start(box3,True,True,0)
        self.store=Gtk.ListStore(str,str,str,str)
        for a in sorted(accounts.accountbook):
            self.store.append([a,accounts.accountbook[a][0],accounts.accountbook[a][1],accounts.accountbook[a][2]])
        scw = Gtk.ScrolledWindow()
        box3.pack_start(scw,True,True,0)
        treeview = Gtk.TreeView(model=self.store)
        renderer = Gtk.CellRendererText()
        column1=Gtk.TreeViewColumn("科目",renderer, text=0)
        treeview.append_column(column1)
        column2=Gtk.TreeViewColumn("科目名稱",renderer, text=1)
        treeview.append_column(column2)
        column3=Gtk.TreeViewColumn("類型",renderer, text=2)
        treeview.append_column(column3)
        column4=Gtk.TreeViewColumn("細目",renderer,text=3)
        treeview.append_column(column4)
        scw.add(treeview)
        self.adding=False
        
    def add_account(self,widget):
        if not self.adding:
            self.adding=True
            self.fgrid=Gtk.Grid()
            self.box2.add(self.fgrid)
            codelab=Gtk.Label("科目編號")
            self.fgrid.attach(codelab,0,0,2,1)
            alab=Gtk.Label("科目")
            self.fgrid.attach_next_to(alab,codelab,Gtk.PositionType.BOTTOM,2,1)
            typlab=Gtk.Label("類別")
            self.fgrid.attach_next_to(typlab,alab,Gtk.PositionType.BOTTOM,2,1)
            catlab=Gtk.Label("細目")
            self.fgrid.attach_next_to(catlab,typlab,Gtk.PositionType.BOTTOM,2,1)
            
            self.codeent=Gtk.Entry()
            self.fgrid.attach_next_to(self.codeent,codelab,Gtk.PositionType.RIGHT,1,1)
            self.acntent=Gtk.Entry()
            self.fgrid.attach_next_to(self.acntent,alab,Gtk.PositionType.RIGHT,1,1)
            self.typ = Gtk.ComboBoxText()
            self.fgrid.attach_next_to(self.typ,typlab,Gtk.PositionType.RIGHT,1,1)
            self.typ.set_entry_text_column(0)
            for a in account_types:
                 self.typ.append_text(a)
            self.typ.set_active(0)
            self.typ.set_id_column(0)
            self.cat = Gtk.ComboBoxText()
            self.fgrid.attach_next_to(self.cat,catlab,Gtk.PositionType.RIGHT,1,1)
            for a in cat_types:
                self.cat.append_text(a)
            self.cat.set_entry_text_column(0)
            self.cat.set_active(0)
            self.cat.set_id_column(0)
            addb=Gtk.Button("增加")
            self.fgrid.attach_next_to(addb,catlab,Gtk.PositionType.BOTTOM,1,1)
            canb=Gtk.Button("取消")
            self.fgrid.attach_next_to(canb,addb,Gtk.PositionType.RIGHT,1,1)
            addb.connect("clicked",self.add_action)
            canb.connect("clicked",self.cancel_action)
            self.show_all()

    def edit_account(self,widget):
        pass
            
    def add_action(self,widget):
        self.box2.remove(self.fgrid)
        self.adding=False
        

    def cancel_action(self,widget):
        self.box2.remove(self.fgrid)
        self.adding=False

    def saveaccount(self,widget):
        accounts.dump_accounts()
        __main__.win.log("Accounts saved")

    def importaccount(self,widget):
        print("Import Accounts")
        self.but=Gtk.Button("導入")
        self.box2.pack_start(self.but,False,False,0)
        self.lab=Gtk.Label("檔案名稱:")
        self.box2.pack_start(self.lab,False,False,0)
        self.importent=Gtk.Entry()
        self.box2.pack_start(self.importent,False,False,0)
        self.but.connect("clicked",self.importing)
        self.show_all()
        
    def importing(self,widget):
        fname=self.importent.get_text()
        accounts.importcsv(fname)
        self.store.clear()
        for a in sorted(accounts.accountbook):
            self.store.append([a,accounts.accountbook[a][0],accounts.accountbook[a][1],accounts.accountbook[a][2]])
        self.show_all()
        self.box2.remove(self.lab)
        self.box2.remove(self.but)
        self.box2.remove(self.importent)

        
        
        
accounts=Accounts()
