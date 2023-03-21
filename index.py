from tkinter import *
from tkinter import ttk
from subprocess import Popen
import chromedriver_autoinstaller
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from msedge.selenium_tools import Edge, EdgeOptions
from subprocess import CREATE_NO_WINDOW
from selenium import webdriver
import time
import importlib
import os
import zipfile
import urllib.request
import requests
from io import BytesIO
import pandas as pd
import numpy as np
from openpyxl import load_workbook



class App:

    def __init__(self):
        self.root = Tk()
        self.colorTeam = '#E11419'
        self.root.geometry('350x150')
        self.root.config(bg='#fff')
        self.link = 'https://poliedrodist.comcel.com.co/'
        self.linktemp = 'https://www.google.com/'
        self.link2 = 'https://poliedrodist.comcel.com.co/activaciones/http/REINGENIERIA/pagDispatcherEntradaModernizacion.asp?Site=1'
        self.interfas()

    def interfas(self):
        titulo = Label(self.root, text= "Activaciones Team")
        titulo.pack(anchor = CENTER)
        titulo.config(fg=self.colorTeam, bg='white', font=("Verdana",18))

        titulo = Label(self.root, text= "Comunicaciones")
        titulo.pack(anchor = CENTER)
        titulo.config(fg=self.colorTeam, bg='white', font=("Verdana",18))
        
        botonExcel = Button(self.root, text="ABRIR LISTA", command=self.abrirArchivo, bg=self.colorTeam, fg="white")
        botonExcel.place(relx=0.00, rely= 0.5, relwidth = 0.33, relheight= 0.5)

        botonActivar = Button(self.root, text="ABRIR PAGINA", command=self.abrirPagina, bg=self.colorTeam, fg="white")
        botonActivar.place(relx=0.333, rely= 0.5, relwidth = 0.33, relheight= 0.5)

        botonActivar = Button(self.root, text="ACTIVAR LINEAS", command=self.activarLineas, bg=self.colorTeam, fg="white")
        botonActivar.place(relx=0.666, rely= 0.5, relwidth = 0.33, relheight= 0.5)

    def abrirArchivo(self):
        p = Popen("openExcel.bat")
        stdout, stderr = p.communicate()
    
    def excel(self):
        file = "lineas.xlsx"
        fileExcel = pd.read_excel(file)
        # numbers = np.asarray(fileExcel)
        numbers = fileExcel
        return numbers
    
    def saveExcel(self,posicion):
        self.lineas['Min'][posicion]= self.min
        self.lineas['Mensaje'][posicion]= self.mensaje
        self.lineas['ICC_ID_Identificacion_Tarjeta_de_Circuito_Integrada'][posicion]= self.icc
        self.lineas['IMEI_Identificacion_Internacional_del_Equipo_Movil'][posicion]= self.imei
        self.lineas['Validacion_Tecnologia'][posicion]= self.vTecnologia
        self.lineas['Validacion_Kit_Prepago'][posicion]= self.vKit
        self.lineas['Validacion_Region_ICCID_Distribuidor'][posicion]= self.vRegion
        self.lineas['Validacion_Equipo'][posicion]= self.vEquipo
        self.lineas['Validacion_Lista'][posicion]= self.vLista
        self.lineas.to_excel('lineas.xlsx', index=False)
    
    def quitarFormatoCientifico(self, cantidad):
        for i in range (0,cantidad):
            self.lineas['Iccid'][i]= " "+str(self.lineas['Iccid'][i]).strip()
            print(self.lineas['Iccid'][i])

    def activarLineas(self):
        self.elegirOpcion()
        time.sleep(2)
        self.lineas = self.excel()
        cantidad = len(self.lineas['Iccid'])
        self.quitarFormatoCientifico(cantidad)
        for i in range (0,cantidad):
            dato1= str(self.lineas['Imei'][i])
            dato2= str(self.lineas['Iccid'][i]).strip()
            dato2= dato2[5:]
            dato3= str(self.lineas['Cedula vendedor'][i])
            print(dato1, dato2, dato3)
            self.llenarInfo(dato1, dato2, dato3)
            time.sleep(2)
            self.saveExcel(i)
            


   



    def abrirPagina(self):
        self.installEdge()
        self.openEdge()
        

    def continuar(self,by,str):
        validate = True
        while(validate):
            try:
                if by == "xpath": find = self.browser.find_element_by_xpath(str)
                elif by == "id": find = self.browser.find_element_by_id(str)
                elif by == "name": find = self.browser.find_element_by_name(str)
                validate = FALSE
            except:
                print('Cargando')
                time.sleep(1)

    def errorControl(self,func,by,str):
        validate = True
        while(validate):
            try:
                func(by,str)
                validate = FALSE
            except:
                print('Cargando')
                time.sleep(1)

       

    def elegirOpcion(self):

        
  
        self.browser.get(self.link2)
        time.sleep(0.5)
        self.click('name', 'shortcutProduct')

    def llenarInfo(self,equipo,sim,vendedor):
        self.continuar('id', 'DetailProduct_SellerId')
        self.insert('id','DetailProduct_SellerId', vendedor)
        self.insert('id','DetailProduct_Imei', equipo)
        self.insert('id','DetailProduct_Iccid', sim)
        self.continuar('id', 'btnNext')
        time.sleep(2)
        self.click('id', 'btnNext')
        time.sleep(5)
        self.identificarPaso()

    def identificarPaso(self):
        op = 0
        try:
            intento1 = self.leerData('/html/body/div/div[2]/section/div/div[2]/div[2]/main/form/div[2]/div[4]/ul/li')
            if intento1 == "none":
                valor = False
            elif "no pertenece" in intento1:
                valor = True
                op = 1
                self.mensaje = intento1
            else:
                valor = True
                op = 1
                self.mensaje = intento1
        except Exception as e:
            valor = False

        if (valor == False):
            try:
                intento2 = self.leerData('/html/body/div/div[2]/section/div/div[2]/div[2]/main/form/div/div[6]/div/span')
                if intento2 == "none":
                    valor = False
                elif "Correcta" in intento2:
                    valor = True
                    op = 2
                    self.mensaje = intento2
            except Exception as e:
                valor = False
        if (valor == False):
            try:
                op = 3
                self.mensaje = 'No deja preactivar por seriales en uso o principal'
            except Exception as e:
                print('ninguna coincide')
        
        if op==1:
            print('error1')
            time.sleep(1)
            self.error1()
        if op==2:
            print('valida')
            time.sleep(1)
            self.validado()
        if op==3:
            print('error2')
            time.sleep(1)
            try:
                self.error2()
            except:
                pass
        
        

    def validado(self):
        self.icc = ""
        self.imei = ""
        self.vTecnologia = ""
        self.vKit = ""
        self.vLista = ""
        self.vEquipo = ""
        self.vRegion = ""
        self.continuar('id','btnNext')
        self.click('id', 'btnNext')
        time.sleep(1)
        self.continuar('id','btnNext')
        self.click('id', 'btnNext')
        self.continuar('xpath', '/html/body/div/div[2]/section/div/div[2]/div[2]/main/form/div/div[1]/div/div[2]/div/div[1]/div[3]/div/span/span[1]/span/span[1]')
        self.click('xpath', '/html/body/div/div[2]/section/div/div[2]/div[2]/main/form/div/div[1]/div/div[2]/div/div[1]/div[3]/div/span/span[1]/span/span[1]')
        self.continuar('xpath','/html/body/span/span/span[2]/ul/li[2]')
        self.click('xpath','/html/body/span/span/span[2]/ul/li[2]')  
        time.sleep(2)
        self.continuar('id','btnNext')
        self.errorControl(self.click,'id','btnNext')
        time.sleep(1)
        self.continuar('id','btnNext')
        self.click('id', 'btnNext')
        self.continuar('xpath', '/html/body/div/div[2]/section/div/div[2]/div[2]/main/div/div/div/strong/strong/div/div/div/p/strong[2]')
        self.min = self.leerData('/html/body/div/div[2]/section/div/div[2]/div[2]/main/div/div/div/strong/strong/div/div/div/p/strong[2]')
        print(self.min)
        self.continuar('id','btnPrev')
        self.click('id','btnPrev')
        self.continuar('name', 'shortcutProduct')
        self.click('name', 'shortcutProduct')
        


        # ECM5430C
        # Marzo23*$%
        

    def error1(self):
        self.icc = ""
        self.imei = ""
        self.vTecnologia = ""
        self.vKit = ""
        self.vLista = ""
        self.vEquipo = ""
        self.vRegion = ""
        self.min = ""
        self.borrar('id','DetailProduct_Imei')
        time.sleep(0.5)
        self.borrar('id','DetailProduct_Iccid')
        time.sleep(0.5)
        self.borrar('id','DetailProduct_SellerId')
        time.sleep(0.5)
        self.borrarLetras('id','DetailProduct_SellerId',8)
        time.sleep(0.5)


    def error2(self):
        self.icc = self.leerData('/html/body/div/div[2]/section/div/div[2]/div[2]/main/form/div/div[3]/div[2]/div[1]/div/div/div')
        self.imei = self.leerData('/html/body/div/div[2]/section/div/div[2]/div[2]/main/form/div/div[3]/div[2]/div[2]/div/div/div')
        self.vTecnologia = self.leerData('/html/body/div/div[2]/section/div/div[2]/div[2]/main/form/div/div[3]/div[2]/div[3]/div/div/div')
        self.vKit = self.leerData('/html/body/div/div[2]/section/div/div[2]/div[2]/main/form/div/div[3]/div[2]/div[4]/div/div/div')
        self.vLista = self.leerData('/html/body/div/div[2]/section/div/div[2]/div[2]/main/form/div/div[3]/div[2]/div[5]/div/div/div')
        self.vEquipo = self.leerData('/html/body/div/div[2]/section/div/div[2]/div[2]/main/form/div/div[3]/div[2]/div[6]/div/div/div')
        self.vRegion = self.leerData('/html/body/div/div[2]/section/div/div[2]/div[2]/main/form/div/div[3]/div[2]/div[7]/div/div/div')
        self.min= ""
        time.sleep(0.5)
        self.click('id', 'btnPrev')
        time.sleep(0.5)
        self.borrar('id','DetailProduct_Imei')
        time.sleep(0.5)
        self.borrar('id','DetailProduct_Iccid')

    
    def leerData(self, path):
        data = self.browser.find_element_by_xpath(path)
        if data is not None:
            # print(data.text)
            return data.text
        else: return "none"
    


    def click(self, by, str):
        if by == "xpath": find = self.browser.find_element_by_xpath(str)
        elif by == "id": find = self.browser.find_element_by_id(str)
        elif by == "name": find = self.browser.find_element_by_name(str)
        else: find =None
        if find is not None:
            find.click()
        
    def insert(self, by, str, text):
        if by == "xpath": find = self.browser.find_element_by_xpath(str)
        elif by == "id": find = self.browser.find_element_by_id(str)
        elif by == "name": find = self.browser.find_element_by_name(str)
        else: find =None
        if find is not None:
            find.send_keys(text)
        
    def borrar(self, by, str):
        if by == "xpath": find = self.browser.find_element_by_xpath(str)
        elif by == "id": find = self.browser.find_element_by_id(str)
        elif by == "name": find = self.browser.find_element_by_name(str)
        else: find =None
        if find is not None:
            find.clear()
        
    def borrarLetras(self, by, str, cantidad):
        
        if by == "xpath": find = self.browser.find_element_by_xpath(str)
        elif by == "id": find = self.browser.find_element_by_id(str)
        elif by == "name": find = self.browser.find_element_by_name(str)
        else: find =None
        if find is not None:
            for i in range(0,cantidad):
                find.send_keys(Keys.BACKSPACE)
        
    def select(self, by, str, text):
        
        if by == "xpath": find = self.browser.find_element_by_xpath(str)
        elif by == "id": find = self.browser.find_element_by_id(str)
        elif by == "name": find = self.browser.find_element_by_name(str)
        else: find =None
        if find is not None:
            select = Select(find)
            select.select_by_visible_text(text)
    

    
    def openEdge(self):
        options = EdgeOptions()
        options.use_chromium = True
        options.add_argument("start-maximized")
        self.browser = Edge(executable_path='msedgedriver.exe', options=options)
        self.browser.get(self.link)


    def installEdge(self):
        # Obtener la última versión del controlador de Microsoft Edge WebDriver
        response = requests.get('https://msedgewebdriverstorage.blob.core.windows.net/edgewebdriver/LATEST_STABLE')
        latest_version = response.text.strip()
        print(latest_version)

        # URL de descarga del controlador
        url = f'https://msedgedriver.azureedge.net/{latest_version}/edgedriver_win64.zip'

        # Descargar y extraer el archivo zip del controlador
        response = urllib.request.urlopen(url)
        zipfile.ZipFile(BytesIO(response.read())).extractall(os.getcwd())

        # Agregar el controlador al PATH del sistema
        os.environ['PATH'] += os.pathsep + os.getcwd()


root = App ()
root.root.mainloop()

# root.abrirPagina()
# root.activarLineas()

# EC0717A      Clave Febrer*+17
