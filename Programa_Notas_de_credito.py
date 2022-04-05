#Libraries to open filedialog
from tkinter import *
from tkinter import filedialog
import tkinter as tk
from PIL import Image
from PIL import ImageTk
import cv2
import imutils
import numpy as np
#To get user name
import getpass
#To read files on directory
import os
from os.path import isfile, join
import errno
#Excel Libraries
import pandas as pd
#To open file
import win32com.client
#to create an excel file
from openpyxl import Workbook
import webbrowser as wb
import shutil


#Globals
titleDialog = "Selecciona una carpeta"
mainDirectory = "C:/Users/"
creditNotesDirectory = "/ABB/Customer Support - Documentos/Notas de crédito/Pago oportuno"
user=getpass.getuser()
logoRute = "logo.png" #C:/Users/"+user+"/Documents/Carlos/Mis propios trabajos/automati/para el share point/
#programRute = "C:/Users/"+user+"/Documents/Carlos/Mis propios trabajos/automati/para el share point/"
excelFileName = "documento.xlsx"

def open_directory():
    os.startfile(ruta)

def open_excel_file():
    book = Workbook()
    sheet = book.active
    book.save(excelFileName)
    wb.open_new(excelFileName) #wb.open_new(programRut+excelFileName)

def open_txt_file():
    wb.open_new('tempD.txt') #wb.open_new(programRute+'tempD.txt')

def copyFiles():

    lines = lstNCNotOk.get(0, END)
    print(lines)
    
    #with open('tempT.txt') as f_obj:
    #    lines = f_obj.readlines()

        
    try:
        os.mkdir(ruta+"/3 Uploaded")
    except OSError as e:
        if e.errno != errno.EEXIST:
            raise
    
    for copyLines in lines:

        shutil.copy(ruta+"/"+copyLines, ruta+"/3 Uploaded/"+copyLines)

    
    
    

def elegir_ruta():
    os.system("cls")
     
    width_value=root.winfo_screenwidth()
    height_value=root.winfo_screenheight()
    root.geometry("%dx%d+0+0" % (width_value, height_value))

    global ruta
    global i
    global e
    global c
    global v
    global lstNCNotOk

    ruta = filedialog.askdirectory(title=titleDialog, initialdir= mainDirectory+user+creditNotesDirectory)
    

    
    pasoDosBA = tk.Button(root, wraplength=371,text="Lista de archivos de NC listas para procesar", command=open_txt_file)
    pasoDosBA.place(x=950,y=415)


    pasoDos = Label(root, text="______________________________________________\n PASO 2) \n\n Copiar archivos de NC listas para procesar procesarse \n desde este botón",bg="white")#.place(x=400,y=28)
    pasoDos.place(x=925,y=450)
    pasoDosB = tk.Button(root, wraplength=371,text="Copiar Archivos", command=copyFiles)
    pasoDosB.place(x=1020,y=525)



    pasoTres = Label(root, text="PASO 3) \n Ejecutar boton de Master Data en excel",bg="white")#.place(x=400,y=28)
    pasoTres.place(x=960,y=550)
    pasoDosB = tk.Button(root, wraplength=371,text="Abrir excel", command=open_excel_file)
    pasoDosB.place(x=1030,y=585)

    
 
    archivosInfoAS = Label(root, text="Estas notas de crédito NECESITAN REVISION: \n(DOCUMENT NUMBER)",bg="white")#.place(x=400,y=28)
    archivosInfoAS.place(x=910,y=10)#place(x=400,y=420)
    archivosInfoASD = Label(root, text="Estas notas de crédito  NECESITAN \n REVISION: (Ammount in document currency)",bg="white")#.place(x=400,y=28)
    archivosInfoASD.place(x=910,y=155)
    archivosInfoAD = Label(root, text="Estas notas de crédito requieren revisión, no coinciden\n el Ammount in document currency y el Document Number :",bg="white").place(x=400,y=200)
    #archivosInfoAD.grid(column=600, row=4,padx=5,pady=5)
    archivosInfoAQ = Label(root, text="Estas Notas de crédito estan listas para procesarse: ",bg="white").place(x=400,y=10)
    #archivosInfoAQ.grid(column=600, row=22,padx=50,pady=50)      
    infoSevenListBox = Label(root, text="Estas Notas de crédito no contienen datos",bg="white").place(x=910, y=290) 

    width_value=root.winfo_screenwidth()
    height_value=root.winfo_screenheight()
    root.geometry("%dx%d+0+0" % (width_value, height_value))

 

    rutaInfo = tk.Label(root, wraplength=200, text="Ruta Seleccionada: ",bg="white").place(x=135,y=180)
    #rutaInfo.grid(column=0, row=4,padx=5,pady=5) 

    rutaInfoB = tk.Button(root, wraplength=371,text=ruta, command=open_directory).place(x=10,y=199) #(x=100,y=150)
    #rutaInfoB.grid(column=0, row=5)#padx=5,pady=5) 


    contenido = os.listdir(ruta)
    archivos = [nombre for nombre in contenido if isfile(join(ruta,nombre))]


    archivosInfo = Label(root, text="Contiene estos archivos: ",bg="white").place(x=125,y=245)
    #archivosInfo.grid(column=0, row=6,padx=5,pady=5) 

    lbDoc=Listbox(root,width=60, selectmode='EXTENDED', height=23)
    lbDoc.place(x=10,y=270)

    lstNCNot=Listbox(root,width=50, height=5,selectmode='EXTENDED')
    lstNCNot.place(x=910,y=54)#place(x=400,y=444 ) #444
    lstNCNotD=Listbox(root,width=50, height=5,selectmode='EXTENDED')
    lstNCNotD.place(x=910,y=190)#place(x=400,y=444 ) #444


    lstNCNot2=Listbox(root,width=70, selectmode='EXTENDED')
    lstNCNot2.place(x=400,y=240)
    lstNCNotOk=Listbox(root,width=70, selectmode='EXTENDED')
    lstNCNotOk.place(x=400,y=30)

    docsOkOk = Label(root, text="Estos documentos ya estan procesados. Por favor, cambielos de carpeta",bg="white")#.place(x=400,y=28)
    docsOkOk.place(x=400,y=420)	#place(x=920,y=10)

    lbDocsOk=Listbox(root,width=70, selectmode='EXTENDED')
    lbDocsOk.place(x=400,y=444 )#place(x=910,y=54) x=910,y=90)

    lbSix=Listbox(root, width=50, height=5)
    lbSix.place(x=910,y=320)

    for i in archivos:

        lbDoc.insert(+1,i)
        lbCicno = lbDoc.size()

        nDocsMain= Label(root, text=lbCicno,bg="white").place(x=300,y=245)

        excelDirectory = ruta+"/"+i

        xl =pd.ExcelFile(excelDirectory)
        res = len(xl.sheet_names)


        if res > 3 :
            lbDocsOk.insert(+1,i)
        else:
            #wb = pd.read_excel(excelDirectory)
            indicator = pd.read_excel(excelDirectory, sheet_name="Bank accounts records", header=None).iloc[15:, 4] #Seleccionar columna iloc [; 3]
            valOne = pd.read_excel(excelDirectory, sheet_name="Bank accounts records", header=None).iloc[15:, 10]
            valTwo = pd.read_excel(excelDirectory, sheet_name="Cleared items report", header=None).iloc[9:, 2]
            valThree = pd.read_excel(excelDirectory, sheet_name="Cleared items report", header=None).iloc[9:, 10].style('background:#70AD47')
            
            print(valThree)
            

          


            valOneNeg = valOne * -1

            numA = valOneNeg.astype(float)
            numB = valThree.astype(float)
            numOne = indicator.astype(float)
            numTwo = valTwo.astype(float)

            priceA = round(numA)
            priceB = round(numB)
            docA = round(numOne)
            docB = round(numTwo)
            
            priceC =priceA.to_list()
            priceD =priceB.to_list()
            docOne =docA.to_list()
            docTwo =docB.to_list()

            
            '''print("priceC")
            print(priceC)
            print("priceD")
            print(priceD)
            print("docOne")
            print(docOne)
            print("docTwo")
            print(docTwo)'''
            

            q=""
            w=""
            e=""
            c = []
            d = []
            ñ = []


      


            if priceC == []:
                lbSix.insert(+1,i)                
            else:
                for number in docTwo:
                    if number in docOne:
                        if number not in c:
                            c.append(number)
                #print("C")
                #print(c)
                for numberTwo in priceD:
                    if numberTwo in priceC:
                        if numberTwo not in d:
                            d.append(numberTwo)
                #print("D")            
                #print(d)
                if c == docOne and d == priceC:
                    lstNCNotOk.insert(+1,i)
                elif c != docOne and d != priceC:
                    lstNCNot2.insert(+1,i)
                elif c != docOne and d == priceC:
                    lstNCNot.insert(+1,i)
                elif c == docOne and d != priceC:
                    lstNCNotD.insert(+1,i)    

                lbUno = lstNCNotOk.size()
                lbDos = lstNCNot2.size()
                lbTres = lstNCNot.size()
                lbCuatro = lstNCNotD.size()

                lbSeis = lbDocsOk.size()
                #lbSiete = lbSix.size()

                nArchivosA = Label(root, text=lbUno,bg="white").place(x=790,y=10)
                nArchivosB = Label(root, text=lbDos,bg="white").place(x=790,y=200)
                nArchivosC = Label(root, text=lbTres,bg="white").place(x=1200,y=10)
                nArchivosD = Label(root, text=lbCuatro,bg="white").place(x=1200,y=155)

                nArchivosE = Label(root, text=lbSeis,bg="white").place(x=790,y=420)
                #nArchivosF = Label(root, text=lbSiete).place(x=1200,y=300)
            

               
                with open('temp.txt', 'w') as f:
                    f.write(''.join(lstNCNotOk.get(0, END)))
                    #f.write('\n\n\n')
                    f.close()



                f = open('temp.txt')
                t = open('tempT.txt','w+')
                t.write(f.read().replace('.xlsx','.xlsx \n'))
                f.close()
                t.close()

                fDos = open('tempT.txt')
                tDos = open('tempD.txt','w+')
                tDos.write(fDos.read().replace('200',ruta+'/200'))
                fDos.close()
                tDos.close()




  
   #asdas 
root = Tk()
#scrollbar 
root.geometry("400x190+0+0")
root.config(bg="white")
root.title("Analisis de Notas de credito por pago oportuno V1.1")

#Boton para seleccionar la carpeta
btninfo = tk.Label(root, text="______________________________________\nSaludos "+user+"!\n PASO 1) Selecciona una carpeta a analizar",bg="white").place(x=80,y=100)
#btninfo.grid(column=0, row=2,padx=5,pady=5) 

btn = tk.Button(root, text="Selecciona una carpeta a analizar", width=25, command=elegir_ruta).place(x=100,y=150)
#btn.grid(column=0, row=3,padx=5,pady=5) 


#Label para logo ABB
imagenL=tk.PhotoImage(file=logoRute)
lblImagen=tk.Label(root, image=imagenL)
lblImagen.place(x=104,y=10)

root.mainloop()


