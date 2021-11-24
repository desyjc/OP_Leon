import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape

def CrearPdf(Folio_str, dir_origen):
    print(Folio_str)
    folio_file = dir_origen + "/folio/Folio.pdf"
    c = canvas.Canvas(folio_file, pagesize=landscape(letter))
    c.setFont('Helvetica', 12)
    #folio = 'Folio: ' + 'OP-MEME-SAT-24/08/2021'
    c.drawString(550,550, "Folio: " + str(Folio_str))
    c.save()
