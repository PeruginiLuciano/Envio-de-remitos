import os
import shutil

def funtiones():
    os.chdir('H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos Enertik pdf/')
    for file in os.listdir():
        if (file=="basura.pdf"):
            scrr=file
            desechos='H:/Mi unidad/Escaner_AR/Remitos/2022/Remitos desechos/basura.pdf'
            shutil.move(scrr,desechos)
    print ("termine")