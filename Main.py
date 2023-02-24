from email import header, message
import smtplib
from email.mime.text import MIMEText
from turtle import st, title
import getpass
import cv2
import re
import pytesseract
import os
import shutil
from email.mime.base import MIMEBase
from email.encoders import encode_base64
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from pdf2image import convert_from_path
import time
from PyPDF2 import PdfFileWriter, PdfFileReader
import sys
import csv
from validate_email import validate_email

from dns.resolver import query


#Funcion de Validacion 
def ValidacionDeEmail(Email): 
    validate=Email
    is_valid=validate_email(validate)
    try:
        if(is_valid): 
            domai=validate.rsplit('@',1)[-1]
            final=bool(query(domai,'MX'))
            if final:
                print("Email correcto y verificado")
                return True
        else:
            print("Email invalido")
            return False
    except:
        print("Error de dominio")
        return False
#import divisor
#time.sleep(60)
#**********--------------OBTENER FECHA---------------**************
Hora=time.strftime("%H:%M:%S")
FECHA= datetime.today().strftime('%Y-%m-%d')
#Guardamos variables con destinos para eliminar basura y desechos
try:
    eliminar='H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Enertik pdf/basura.pdf'
    eliminarT='H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Transporte/basuraT.pdf'
    desechos = 'H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos desechos/basura.pdf'
    desechosT = 'H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos desechos/basuraT.pdf'
    shutil.move(eliminar,desechos)

    shutil.move(eliminarT,desechosT)
except:
    print("No hay basuraaaaaa")

#***********--------------CAMBIO DE NOMBRE DE ARCHIVO y MIGRACION A CARPETA EXPECIFICA-------------**************


#Ahora revisa la carpeta de escaneos en pdf
Reporte='H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Reporte/Reportes.csv'
#fichero = open('H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Reporte/Reporte.txt','a',encoding='utf-8')
os.chdir('H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Enertik pdf/')
r=0
#Recorremos la lsita de archivos en la carpeta especifica y cambiamos los nombres
k=1

for file in os.listdir():
    Nom,extencion=os.path.splitext(file)
    scr=file
    dst= str(k)+"_Remito "+str(FECHA)+str(extencion)
    k=k+1
    os.rename(scr,dst)
#Recorremos nuevamente la carpeta porque necesitamos que haya impactado el cambio de nombres
k=1
for file in os.listdir():
    #Arreglo donde guardaremos los mails
    Email=[]
    #Funcion para separar el nombre y la extencion del archivo
    Nom,extencion=os.path.splitext(file)
    scr=file
    dst= str(k)+"_Remito "+str(FECHA)+str(extencion)
    #si la extencion es pdf vamos a usar una funcion para guardar la pagina del pdf
    if(extencion=='.pdf'):
        
        time.sleep(2)
        pdf_file = scr
        
        try:
            pages = convert_from_path(pdf_file)
        except:
            print("Entro una vez")
            try:
                pages = convert_from_path(pdf_file)
            except:
                print("Entro 2 vez")
                try:
                    pages = convert_from_path(pdf_file)
                except:
                    print("Puta vida")
                    print("NO ANDA")

       
        img_file = pdf_file.replace(".pdf","")

        count=0
        #recorremos las paginas para convertir la primera en imagen, ya que paython hace tratamiento de imagenes y no de pdf
        for page in pages:
            
            jpeg_file=img_file+".jpg"
            page.save(jpeg_file,'JPEG') 
            Nombre2='H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Enertik pdf/'+str(jpeg_file)
            Nombre='H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Enertik pdf/'+str(jpeg_file)
            NombreA='H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Enertik pdf/'+str(dst)
    try:
        #nos desacemos de un archivo ini que se encuentra en la nube
        if(extencion=='.ini'):
            des="H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos desechos/"+str(dst)
            shutil.move(scr,des)
        else:
            pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
            #cargo imagen
            image= cv2.imread(Nombre)
            #cargo el texto leido en la imagen
            text = pytesseract.image_to_string(image)
            #print(text)
            #print("TExto",text)
            
            cv2.waitKey(0)
            cv2.destroyAllWindows()
            #patron capaz de reconocer los Email
            patron = r"[a-zA-ZÀ-ÿ0-9\u00f1\u00d1._%+-]+@[a-zÀ-ÿ0-9\u00f1\u00d1.-]+.[a-zA-Z]{2,}"
            #compara el patron con el texto extraido de la imagen y asi po der encontrar las 
            coincidencias =re.findall(patron,text)
            numeroRemito="000([0-9]+)"
            coincidenciaRemito=re.findall(numeroRemito,text)
            nombreRemito='H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Enertik pdf/000'+str(coincidenciaRemito[1])+str(extencion)
            os.rename(NombreA,nombreRemito)
            for coincidencia in coincidencias:
                Email.append(coincidencia)
            i=len(Email)
            
            if not Email:
                print("No hay Email")
            else:
                if (i-1)>0:
                    print("Hay",(i-1),"Email")
                    if i>2:
                        Email[1]=Email[2]

                    try:
                        for j in range(i-1):
                            
                            print("Email:",Email[j+1])
                    except:
                        print("No hay email")
                    




            #*******---------ENVIO DE EMAIL---------------******


            print ("*** ENviar email con Gmail ***")
            user = "gestion@enertik.ar"
            password ="brEuFrQ5RuoxK#mhE2@6"
            print(Email[1])
            print(Email[0])
            #Para las cabeceras del email
            remitente = "Enertik <enertikcliente@gmail.com"
            destinatario = Email[1]
            asunto= "Adjunto de remito"
            mensaje = "Estimado cliente se adjunta remito \n \n Saludos. "
            #ruta de la imagen a adjuntar
            #if extencion==".JPG":
            #   archivo= Nombre
            if extencion ==".pdf":
                archivo=nombreRemito


            #Host y puerto SMTP de Gmail
            gmail = smtplib.SMTP('smtp.gmail.com', 587)

            #protocolo de cifrado de datos utilizado por gmail
            gmail.starttls()

            #Credenciales
            gmail.login(user,password)

            #muestra la depuracion de la operacion de envio 
            gmail.set_debuglevel(1)

            header=MIMEMultipart()

            header['subject']= asunto
            header['From'] = remitente
            header['To'] = destinatario

            mensaje = MIMEText(mensaje,'plain')
            header.attach(mensaje)
            #if para ver si hay un archivo adjunto
            if(os.path.isfile(archivo)):
                adjunto= MIMEBase('aplication','octet-stream')
                adjunto.set_payload(open(archivo,"rb").read())
                encode_base64(adjunto)
                adjunto.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(archivo))
                header.attach(adjunto)
            #enviar email
            if(i>1):
                if(Email[1]=="info@electronicsa.com.ar"):
                    Email[1]="Luciano.perugini@enertik.ar"
                elif(Email[1]=="hurbietal@gmail.com"):
                    Email[1]="hurbieta1@gmail.com"
                elif(Email[1]=="odoi.net.ar@gmail.com"):
                    Email[1]="odoi.net.ar@gmail.com"
                if ValidacionDeEmail(Email[1]):
                    gmail.sendmail(remitente,Email[1],header.as_string())
                    #Cerrar la conexion SMTP
                    gmail.quit()
                    Email.clear()
                    enviados = 'H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Enviados/000'+str(coincidenciaRemito[1])+str(extencion)
                    shutil.move(nombreRemito,enviados)
                    shutil.move(Nombre,desechos)
                    Repor=[FECHA,Hora,str(coincidenciaRemito[1]),destinatario,'Enertik','Enviado']
                    with open(Reporte,'a',newline='')as file:
                        write=csv.writer(file,delimiter=';')
                        write.writerow(Repor)
                    #fichero.write("\n Se envio correctamente el remito pdf "+str(dst)+" al mail: "+destinatario+"  "+str(FECHA))
                    k+=1 
                    
                else:
                    Repor=[FECHA,Hora,str(coincidenciaRemito[1]),'Error 03','Enertik','Sintaxis incorrecta o dominio inexistente']
                    with open(Reporte,'a',newline='')as file:
                        write=csv.writer(file,delimiter=';')
                        write.writerow(Repor)
                    des="H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos desechos/000"+str(coincidenciaRemito[1])+str(extencion)
                    shutil.move(Nombre,des)
                    err="H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Errores/000"+str(coincidenciaRemito[1])+str(extencion)
                    shutil.move(nombreRemito,err)
                    #*******---------ENVIO DE EMAIL---------------******


                    print ("*** ENviar email con Gmail ***")
                    user = "gestion@enertik.ar"
                    password ="brEuFrQ5RuoxK#mhE2@6"
                    
                    
                    #Para las cabeceras del email
                    print("Entre aca papaapapapapaa")
                    remitente = "Enertik <enertikcliente@gmail.com"
                    destinatario = "deposito@enertik.ar"
                    asunto= "Error de remito"
                    mensaje = "El remito adjunto no se envio correctamene porque la sintaxis del mail es incorrecta o el domino no existe, por favor verificarlo y enviarlo manualmente.\nSi no estas seguro consulta con el Luchi \n \n Saludos. "
                    #ruta de la imagen a adjuntar
                    #if extencion==".JPG":
                    #   archivo= Nombre
                    if extencion ==".pdf":
                        archivo=nombreRemito


                    #Host y puerto SMTP de Gmail
                    gmail = smtplib.SMTP('smtp.gmail.com', 587)

                    #protocolo de cifrado de datos utilizado por gmail
                    gmail.starttls()

                    #Credenciales
                    gmail.login(user,password)

                    #muestra la depuracion de la operacion de envio 
                    gmail.set_debuglevel(1)

                    header=MIMEMultipart()

                    header['subject']= asunto
                    header['From'] = remitente
                    header['To'] = destinatario

                    mensaje = MIMEText(mensaje,'plain')
                    header.attach(mensaje)
                    #if para ver si hay un archivo adjunto
                    if(os.path.isfile(archivo)):
                        adjunto= MIMEBase('aplication','octet-stream')
                        adjunto.set_payload(open(archivo,"rb").read())
                        encode_base64(adjunto)
                        adjunto.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(archivo))
                        header.attach(adjunto)
                    #enviar email
                    gmail.sendmail(remitente,destinatario,header.as_string())

                    #Cerrar la conexion SMTP
                    gmail.quit()
                    Email.clear()
                    k+=1 
                    
                
                    

            
               
    except:
        des="H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos desechos/000"+str(coincidenciaRemito[1])+str(extencion)
        shutil.move(Nombre,des)
        err="H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Errores/000"+str(coincidenciaRemito[1])+str(extencion)
        e=sys.exc_info()[1]
        print (e)
        print(e.args[0])
        print(e.args[0]," en ", str(dst), "el mial es: ")
        if(i<2):
            #*******---------ENVIO DE EMAIL---------------******


            print ("*** ENviar email con Gmail ***")
            user = "gestion@enertik.ar"
            password ="brEuFrQ5RuoxK#mhE2@6"
            
            
            #Para las cabeceras del email
            print("Entre aca papaapapapapaa")
            remitente = "Enertik <enertikcliente@gmail.com"
            destinatario = "deposito@enertik.ar"
            asunto= "Error de remito"
            mensaje = "El remito adjunto no se envio correctamene porque no se pudo detectar mail, por favor verificarlo y enviarlo manualmente.\nSi no estas seguro consulta con el Luchi \n \n Saludos. "
            #ruta de la imagen a adjuntar
            #if extencion==".JPG":
            #   archivo= Nombre
            if extencion ==".pdf":
                archivo=nombreRemito


            #Host y puerto SMTP de Gmail
            gmail = smtplib.SMTP('smtp.gmail.com', 587)

            #protocolo de cifrado de datos utilizado por gmail
            gmail.starttls()

            #Credenciales
            gmail.login(user,password)

            #muestra la depuracion de la operacion de envio 
            gmail.set_debuglevel(1)

            header=MIMEMultipart()

            header['subject']= asunto
            header['From'] = remitente
            header['To'] = destinatario

            mensaje = MIMEText(mensaje,'plain')
            header.attach(mensaje)
            #if para ver si hay un archivo adjunto
            if(os.path.isfile(archivo)):
                adjunto= MIMEBase('aplication','octet-stream')
                adjunto.set_payload(open(archivo,"rb").read())
                encode_base64(adjunto)
                adjunto.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(archivo))
                header.attach(adjunto)
            #enviar email
            gmail.sendmail(remitente,destinatario,header.as_string())

            #Cerrar la conexion SMTP
            gmail.quit()
            Email.clear()
            Repor=[FECHA,Hora,str(coincidenciaRemito[1]),'Error 01','Enertik','No Enviado']
            with open(Reporte,'a',newline='')as file:
                write=csv.writer(file,delimiter=';')
                write.writerow(Repor)
            #fichero.write("\n Se detecto 0 mail, no esta claro la imagen del remito "+str(dst)+"  "+str(FECHA))
            print("Solamente detecto 1 mail, no esta claro la imagen")
            shutil.move(nombreRemito,err)
            
        else:
            Repor=[FECHA,Hora,e,'Error 02','Enertik','No Enviado']
            with open(Reporte,'a',newline='')as file:
                write=csv.writer(file,delimiter=';')
                write.writerow(Repor)
            #fichero.write("\n Se detecto un error en el envio de pdf "+str(dst)+" el error fue: "+str(e)+"  "+str(FECHA))
            print("Se detecto un error en el envio de pdf")
            shutil.move(nombreRemito,err)
        time.sleep(5)
        k+=1
#AHora revisamos la carpeta transporte
os.chdir('H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Transporte/')
r=0

k=1
for file in os.listdir():
    Nom,extencion=os.path.splitext(file)
    scr=file
    dst= str(k)+"_RemitoT "+str(FECHA)+str(extencion)
    k=k+1
    os.rename(scr,dst)
    
k=1
for file in os.listdir():
    Email=[]
    Nom,extencion=os.path.splitext(file)
    scr=file
    dst= str(k)+"_RemitoT "+str(FECHA)+str(extencion)
    if(extencion=='.pdf'):
        
        time.sleep(2)
        pdf_file = scr
        
        try:
            pages = convert_from_path(pdf_file)
        except:
            print("Entro una vez")
            try:
                pages = convert_from_path(pdf_file)
            except:
                print("Entro 2 vez")
                try:
                    pages = convert_from_path(pdf_file)
                except:
                    print("Puta vida")
                    print("NO ANDA")

       
        img_file = pdf_file.replace(".pdf","")

        count=0
     
        #for page in pages:
        l=0
        for page in pages:
            if l==0:  
                jpeg_file=img_file+"-"+ str(count)+".jpg"
                page.save(jpeg_file,'JPEG') 
                Nombre2='H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Transporte/'+str(jpeg_file)
                Nombre='H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Transporte/'+str(jpeg_file)
                NombreA='H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Transporte/'+str(dst)
                l=1
            else:
                print("listo el pollo")
        
    try:
        if(extencion=='.ini'):
            des="H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos desechos/"+str(dst)
            shutil.move(scr,des)
        else:
            pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
            #cargo imagen
            image= cv2.imread(Nombre)
            #cargo el texto leido en la imagen
            text = pytesseract.image_to_string(image)
            #print(text)
            #print("TExto",text)
            
            cv2.waitKey(0)
            cv2.destroyAllWindows()
            numeroRemito="000([0-9]+)"
            patron = r"[a-zA-ZÀ-ÿ0-9\u00f1\u00d1._%+-]+@[a-zÀ-ÿ0-9\u00f1\u00d1.-]+.[a-zA-Z]{2,}"
            coincidencias =re.findall(patron,text)
            coincidenciaRemito=re.findall(numeroRemito,text)
            nombreRemito='H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Transporte/000'+str(coincidenciaRemito[1])+str(extencion)
            os.rename(NombreA,nombreRemito)
            for coincidencia in coincidencias:
                Email.append(coincidencia)
            i=len(Email)
            
            if not Email:
                print("No hay Email")
            else:
                if i>2:
                        Email[1]=Email[2]
                if (i-1)>0:
                    print("Hay",(i-1),"Email")
                    try:
                        for j in range(i-1):
                            
                            print("Email:",Email[j+1])
                    except:
                        print("No hay email")




            #*******---------ENVIO DE EMAIL---------------******


            print ("*** ENviar email con Gmail ***")
            user = "gestion@enertik.ar"
            password ="brEuFrQ5RuoxK#mhE2@6"
            try:
                print(Email[1])
            except:
                print("No hay mail")
            print(Email[0])
            #Para las cabeceras del email
            remitente = "Enertik <enertikcliente@gmail.com"
            destinatario = Email[1]
            asunto= "Adjunto de remito"
            mensaje = "Estimado cliente se adjunta remito \n \n Saludos. "
            #ruta de la imagen a adjuntar
            #if extencion==".JPG":
            #   archivo= Nombre
            if extencion ==".pdf":
                archivo=nombreRemito


            #Host y puerto SMTP de Gmail
            gmail = smtplib.SMTP('smtp.gmail.com', 587)

            #protocolo de cifrado de datos utilizado por gmail
            gmail.starttls()

            #Credenciales
            gmail.login(user,password)

            #muestra la depuracion de la operacion de envio 
            gmail.set_debuglevel(1)

            header=MIMEMultipart()

            header['subject']= asunto
            header['From'] = remitente
            header['To'] = destinatario

            mensaje = MIMEText(mensaje,'plain')
            header.attach(mensaje)
            #if para ver si hay un archivo adjunto
            if(os.path.isfile(archivo)):
                adjunto= MIMEBase('aplication','octet-stream')
                adjunto.set_payload(open(archivo,"rb").read())
                encode_base64(adjunto)
                adjunto.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(archivo))
                header.attach(adjunto)
            #enviar email
            if(i>1):
                if(Email[1]=="info@electronicsa.com.ar"):
                    Email[1]="Luciano.perugini@enertik.ar"
                elif(Email[1]=="hurbietal@gmail.com"):
                    Email[1]="hurbieta1@gmail.com"
                elif(Email[1]=="odoi.net.ar@gmail.com"):
                    Email[1]="odoi.net.ar@gmail.com"
                if ValidacionDeEmail(Email[1]):
                    gmail.sendmail(remitente,Email[1],header.as_string())
                    #Cerrar la conexion SMTP
                    gmail.quit()
                    Email.clear()
                    enviados = 'H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Enviados/000'+str(coincidenciaRemito[1])+str(extencion)
                    shutil.move(nombreRemito,enviados)
                    shutil.move(Nombre,desechos)
                    Repor=[FECHA,Hora,str(coincidenciaRemito[1]),destinatario,'Transporte','Enviado']
                    with open(Reporte,'a',newline='')as file:
                        write=csv.writer(file,delimiter=';')
                        write.writerow(Repor)
                    #fichero.write("\n Se envio correctamente el remito pdf "+str(dst)+" al mail: "+destinatario+"  "+str(FECHA))
                    k+=1 
                    
                else:
                    Repor=[FECHA,Hora,str(coincidenciaRemito[1]),'Error 03','Transporte','Sintaxis incorrecta o dominio inexistente']
                    with open(Reporte,'a',newline='')as file:
                        write=csv.writer(file,delimiter=';')
                        write.writerow(Repor)
                    des="H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos desechos/000"+str(coincidenciaRemito[1])+str(extencion)
                    shutil.move(Nombre,des)
                    err="H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Errores/000"+str(coincidenciaRemito[1])+str(extencion)
                    shutil.move(nombreRemito,err)
                    #*******---------ENVIO DE EMAIL---------------******


                    print ("*** ENviar email con Gmail ***")
                    user = "gestion@enertik.ar"
                    password ="brEuFrQ5RuoxK#mhE2@6"
                    
                    
                    #Para las cabeceras del email
                    print("Entre aca papaapapapapaa")
                    remitente = "Enertik <enertikcliente@gmail.com"
                    destinatario = "deposito@enertik.ar"
                    asunto= "Error de remito"
                    mensaje = "El remito adjunto no se envio correctamene porque la sintaxis del mail es incorrecta o el domino no existe, por favor verificarlo y enviarlo manualmente.\nSi no estas seguro consulta con el Luchi \n \n Saludos. "
                    #ruta de la imagen a adjuntar
                    #if extencion==".JPG":
                    #   archivo= Nombre
                    if extencion ==".pdf":
                        archivo=nombreRemito


                    #Host y puerto SMTP de Gmail
                    gmail = smtplib.SMTP('smtp.gmail.com', 587)

                    #protocolo de cifrado de datos utilizado por gmail
                    gmail.starttls()

                    #Credenciales
                    gmail.login(user,password)

                    #muestra la depuracion de la operacion de envio 
                    gmail.set_debuglevel(1)

                    header=MIMEMultipart()

                    header['subject']= asunto
                    header['From'] = remitente
                    header['To'] = destinatario

                    mensaje = MIMEText(mensaje,'plain')
                    header.attach(mensaje)
                    #if para ver si hay un archivo adjunto
                    if(os.path.isfile(archivo)):
                        adjunto= MIMEBase('aplication','octet-stream')
                        adjunto.set_payload(open(archivo,"rb").read())
                        encode_base64(adjunto)
                        adjunto.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(archivo))
                        header.attach(adjunto)
                    #enviar email
                    gmail.sendmail(remitente,destinatario,header.as_string())

                    #Cerrar la conexion SMTP
                    gmail.quit()
                    Email.clear()
                    k+=1
           
    except:
        #des="H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos desechos/"+str(dst)
        k+=1
        err="H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Errores/000"+str(coincidenciaRemito[1])+str(extencion)
        #shutil.move(Nombre,des)
        e=sys.exc_info()[1]
        print (e)
        print(e.args[0]," en ", str(dst), "el mial es: ")
        if(i<2):
            #*******---------ENVIO DE EMAIL---------------******


            print ("*** ENviar email con Gmail ***")
            user = "gestion@enertik.ar"
            password ="brEuFrQ5RuoxK#mhE2@6"
            
            
            #Para las cabeceras del email
            remitente = "Enertik <enertikcliente@gmail.com"
            destinatario = "deposito@nertik.ar"
            asunto= "Error de remito"
            mensaje = "El remito adjunto no se envio correctamene porque no se pudo detectar mail, por favor verificarlo y enviarlo manualmente.\nSi no estas seguro consulta con el Luchi \n \n Saludos. "
            #ruta de la imagen a adjuntar
            #if extencion==".JPG":
            #   archivo= Nombre
            if extencion ==".pdf":
                archivo=NombreA


            #Host y puerto SMTP de Gmail
            gmail = smtplib.SMTP('smtp.gmail.com', 587)

            #protocolo de cifrado de datos utilizado por gmail
            gmail.starttls()

            #Credenciales
            gmail.login(user,password)

            #muestra la depuracion de la operacion de envio 
            gmail.set_debuglevel(1)

            header=MIMEMultipart()

            header['subject']= asunto
            header['From'] = remitente
            header['To'] = destinatario

            mensaje = MIMEText(mensaje,'plain')
            header.attach(mensaje)
            #if para ver si hay un archivo adjunto
            if(os.path.isfile(archivo)):
                adjunto= MIMEBase('aplication','octet-stream')
                adjunto.set_payload(open(archivo,"rb").read())
                encode_base64(adjunto)
                adjunto.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(archivo))
                header.attach(adjunto)
            #enviar email
            if(i>1):
                gmail.sendmail(remitente,Email[1],header.as_string())

            #Cerrar la conexion SMTP
            gmail.quit()
            Email.clear()
            Repor=[FECHA,Hora,str(coincidenciaRemito[1]),'Error 01','Transporte','No Enviado']
            with open(Reporte,'a',newline='')as file:
                write=csv.writer(file,delimiter=';')
                write.writerow(Repor)
            #fichero.write("\n Solamente detecto 1 mail, no esta claro la imagen del remito "+str(dst)+"  "+str(FECHA))
            print("Solamente detecto 1 mail, no esta claro la imagen")
            shutil.move(nombreRemito,err)
        else:
            Repor=[FECHA,Hora,e,'Error 02','Transporte','No Enviado']
            with open(Reporte,'a',newline='')as file:
                write=csv.writer(file,delimiter=';')
                write.writerow(Repor)
            #fichero.write("\n Se detecto un error en el envio de pdf "+str(dst)+"  "+str(FECHA))
            print("Se detecto un error en el envio de pdf")
        time.sleep(5)
                

 
    
