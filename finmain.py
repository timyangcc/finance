
import gi
gi.require_version("Gtk","3.0")
from gi.repository import Gtk

import account
import voucher
import report

class FinMain(Gtk.Window):
    def __init__(self):
        Gtk.Window.__init__(self,title="財務管理")
        self.set_size_request(500,800)

        vbox=Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.add(vbox)
        fbox1=Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)
        self.fbox2=Gtk.Box()
        vbox.pack_start(fbox1,False,False,0)
        vbox.pack_start(self.fbox2,True,True,0)
        scw=Gtk.ScrolledWindow()
        self.fbox2.pack_start(scw,True,True,0)
        logwin=Gtk.TextView()
        self.logbuf=logwin.get_buffer()
        scw.add(logwin)
        self.log("\n\nSystem started\n")

        catbutton=Gtk.Button(label="科目管理")
        fbox1.pack_start(catbutton,False,False,0)
        catbutton.connect("clicked",self.on_cat)

        vocmbutton=Gtk.Button(label="傳票管理")
        fbox1.pack_start(vocmbutton,False,False,0)
        vocmbutton.connect("clicked",self.on_voucher)

        repbutton=Gtk.Button(label="報表")
#       repvbutton=Gtk.Button(label="傳票")
#        repjbutton=Gtk.Button(label="日記帳")
#        repgbutton=Gtk.Button(label="總帳")
        fbox1.pack_start(repbutton,False,False,0)
        repbutton.connect("clicked",self.on_report)


    def on_cat(self,widget):
        print("Open Account Window")
        catbox=account.AccountBox()
        catbox.show_all()

    def on_voucher(self,widget):
        print("Open Vouche Window")
        vbox=voucher.VoucherWin()
        vbox.show_all()

    def on_report(self,widget):
        print("Open Report Window")
        rbox=report.ReportWin()
        rbox.show_all()
        
    def log(self,text):
        ip=self.logbuf.get_end_iter()
        self.logbuf.insert(ip,text)

win=FinMain()
win.connect("destroy",Gtk.main_quit)

try:
    account.accounts.load_accounts()
    win.log("Account DB loaded.\n")
except FileNotFoundError:
    win.log("No Account DB!")
try:
    voucher.voucherbook.loadbook()
    win.log("Voucher DB loaded.\n")
except FileNotFoundError:
    win.log("No Voucher DB!")


win.show_all()
Gtk.main()
