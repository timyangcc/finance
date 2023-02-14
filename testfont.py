from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

def hello(c):
    c.setFont('新細明體',16)
    c.drawString(10,100,'開門見山')
    


pdfmetrics.registerFont(TTFont('新細明體','mingliu.ttc'))
c=canvas.Canvas('hello.pdf')
hello(c)
c.showPage()
c.save()
