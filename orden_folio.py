import PyPDF2


def Orden_final(dir_origen, Folio_str, PdfFileName):
    new_file = Folio_str + '.pdf'
    new_path = dir_origen + "/final/" + new_file

    if os.path.isfile(new_path) == True:
        Folio_full = Folio_str + 'x'
 
    else:
        Folio_full = str(Folio_str)


    op_pdf = dir_origen + "/" + PdfFileName
    ordenOrigin = open(op_pdf, 'rb')
    pdfReader = PyPDF2.PdfFileReader(ordenOrigin)
    op_origin1p = pdfReader.getPage(0)
    file_folio = dir_origen + "/folio/" + "Folio.pdf"
    pdfWatermarkReader = PyPDF2.PdfFileReader(open(file_folio, 'rb'))
    op_origin1p.mergePage(pdfWatermarkReader.getPage(0))
    pdfWriter = PyPDF2.PdfFileWriter()
    pdfWriter.addPage(op_origin1p)
    for pageNum in range(1, pdfReader.numPages):
        pageObj = pdfReader.getPage(pageNum)
        pdfWriter.addPage(pageObj)
    newFile = str(Folio_full) + '.pdf'
    folio_path = dir_origen + "/final/" + newFile 
    resultPdfFile = open(folio_path, 'wb')
    pdfWriter.write(resultPdfFile)
    ordenOrigin.close()
    resultPdfFile.close()