from glob import escape
from turtle import delay
from PyPDF2 import PdfFileWriter, PdfFileReader
import os
import time
import sys
from datetime import datetime
import csv
from validate_email import validate_email
from borrador import funtiones



Aller=0
FECHA= datetime.today().strftime('%d-%h-%Y')
Hora=time.strftime("%H:%M:%S")
fichero = open('H:/Mi unidad/Escaner_AR/Remitos/2023/Remitos Reporte/Reporte.txt','a',encoding='utf-8')

fichero.write("\n\n "+FECHA+"     "+Hora)
"""
def extract_page  (doc_name, page_num):
    pdf_reader = PdfFileReader(open(doc_name,'rb'))
    pdf_writer = PdfFileWriter()

    pdf_writer.addPage(pdf_reader.getPage(page_num))

    with open (f'document-page{page_num}.pdf','wb') as doc_file:
        pdf_writer.write(doc_file)"""
r=0
def splt (doc_name, page_num,r):
    pdf_leido=PdfFileReader(open(doc_name,'rb'))
    j=pdf_leido.getNumPages()
    cantidad_pdf=j/2
    a=0
    pdf_leidos=PdfFileWriter()
    while(a<cantidad_pdf):
        pdf_leidos=PdfFileWriter()
        pdf_leidos.addPage(pdf_leido.getPage(page_num))
        pdf_leidos.addPage(pdf_leido.getPage(page_num+1))
        
        
        with open("docT"+str(a)+"_"+str(r)+".pdf",'wb') as file1:
            pdf_leidos.write(file1)
        
        a=a+1
        page_num=page_num+2
        Aller=a
    
def spl (doc_name, page_num,r):
    pdf_leido=PdfFileReader(open(doc_name,'rb'))
    j=pdf_leido.getNumPages()
    cantidad_pdf=j
    a=0
    pdf_leidos=PdfFileWriter()
    while(a<cantidad_pdf):
        pdf_leidos=PdfFileWriter()
        pdf_leidos.addPage(pdf_leido.getPage(page_num))
        
        
        
        with open("doc"+str(a)+"_"+str(r)+".pdf",'wb') as file1:
            pdf_leidos.write(file1)
        
        a=a+1
        page_num=page_num+1
        Aller=a
        
    
        
os.chdir('H:/Mi unidad/Escaner_AR/Remitos/2023/Remitos Transporte/')
for file in os.listdir():
    scr=file
    os.rename(scr,"basuraT.pdf")

for file in os.listdir():
    B=0
    print(file)
    scr=file
    
    r=r+1
    
    splt(scr,0,r)
    time.sleep(5)

    fichero.write("\n Se realizaron correctamente las divisiones de los remitos")
     
    print("Fin")

os.chdir('H:/Mi unidad/Escaner_AR/Remitos/2023/Remitos Enertik pdf/')
for file in os.listdir():
    scr=file
    os.rename(scr,"basura.pdf")

for file in os.listdir():
    print(file)
    scr=file
    
    r=r+1
    
    spl(scr,0,r)
    time.sleep(5)

    fichero.write("\n Se realizaron correctamente las divisiones de los remitos de Transporte")
             
    print("Fin")
#funtiones()






"""       
i=0
def split_pdf (doc_name, page_num,i):
    
    pdf_reader =PdfFileReader(open(doc_name,'rb'))
    pdf_writer1 = PdfFileWriter()
    pdf_writer2 = PdfFileWriter()

    for page in range(page_num):
        pdf_writer1.addPage(pdf_reader.getPage(page))

    for page in range (page_num, pdf_reader.getNumPages()):
        pdf_writer2.addPage(pdf_reader.getPage(page))

    with open("doc"+str(i)+".pdf",'wb') as file1:
        pdf_writer1.write(file1)

    with open("doc2.pdf",'wb') as file2:
        pdf_writer2.write(file2)
    os.remove("Nuevo"+str(i)+".pdf")
    os.rename("doc2.pdf","Nuevo.pdf")
   

#split_pdf(doc_name,2)
doc_name="doc.pdf"
while(1):
    #try:
        
        i=i+1
        doc_name="Nuevo"+str(i)+".pdf"
        split_pdf(doc_name,2,i)
    #except:
        print("Ya esta")
       # break
"""