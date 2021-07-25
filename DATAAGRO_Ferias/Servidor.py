import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.firefox.options import Options

import numpy as np
import requests
import glob
from datetime import datetime

import xlrd
import glob
import xlsxwriter
import os
import os.path

enlace = "https://www.odepa.gob.cl/contenidos-rubro/boletines-del-rubro/boletin-semanal-de-precios-asoc-gremial-de-ferias-ganaderas"

def getDriver(enlace):
    
    options = Options()
    options.log.level = "trace"
    options.add_argument("--headless")
    options.set_preference("browser.download.manager.showWhenStarting", False)
    options.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/csv")
    driver = webdriver.Firefox(options=options)
    driver.set_page_load_timeout("60")
    driver.get(enlace)
    
    return driver

def normalize(s):
    replacements = (
        ("á", "a"),
        ("'", ""),
        ("é", "e"),
        ("í", "i"),
        ("ó", "o"),
        ("ú", "u"),
        ("ñ", "ni"),
        (";", "_"),
        (",", "_"),
        ("-", "_"),
        (" ", "_"),
        (".xlsx", "")
    )
    for a, b in replacements:
        s = s.replace(a, b).replace(a.upper(), b.upper())
    return s

def filesData():
    # file = "DATAAGRO_Ferias/files/*.xlsx"
    file = "files/*.xlsx"
    files = glob.glob(file)

    archivos = np.array(files)
    # print(archivos)
    
    for i in range(len(archivos)):
        
        f = str(archivos[i])
        # print(f)
        
        name = normalize(f[22:]).lower()
        # print(normalize(f).lower())
        # print(name)
        
        precios = pd.concat([tabla1(f), tabla2(f)])
        # precios.to_excel("DATAAGRO_Ferias/files/consolidados_precios/" + str(name) + "_precios.xlsx", index=False)
        precios.to_excel("files/consolidados_precios/" + str(name) + "_precios.xlsx", index=False)

        # tabla3(f).to_excel("DATAAGRO_Ferias/files/consolidados_cantidad/" + str(name) + "_cantidad.xlsx", index=False)
        tabla3(f).to_excel("files/consolidados_cantidad/" + str(name) + "_cantidad.xlsx", index=False)
    # print(len(archivos))
    
    consolidarPrecios()
    consolidarCantidad()

def tabla1(namefile):
    
    book = xlrd.open_workbook(namefile)
    sh = book.sheet_by_index(1) 
    val = sh.cell_value(1,0)
    
    df = pd.read_excel(namefile, sheet_name="Promedio (5 primeros precios)", skiprows=7)
    
    df["Detalle"] = val
    df.columns = ['Feria','Comuna','Fecha', 'Novillo Gordo', 'Novillo Engorda', 'Vaca Gorda', 'Vaca Engorda', 'Vaquilla Gorda', 'Vaquilla Engorda', 'Toros', 'Terneros', 'Terneras', 'Cerdos', 'Lanares', 'Caballos', 'Detalle']
    df = df[:-1]
    
    return df

def tabla2(namefile):
    book = xlrd.open_workbook(namefile)
    sh = book.sheet_by_index(2) 
    val = sh.cell_value(1,0)
    
    df2 = pd.read_excel(namefile, sheet_name="Precio promedio", skiprows=7)
    
    df2["Detalle"] = val
    df2.columns = ['Feria','Comuna','Fecha', 'Novillo Gordo', 'Novillo Engorda', 'Vaca Gorda', 'Vaca Engorda', 'Vaquilla Gorda', 'Vaquilla Engorda', 'Toros', 'Terneros', 'Terneras', 'Cerdos', 'Lanares', 'Caballos', 'Detalle']
    df2 = df2[:-1]
    
    return df2

def consolidarPrecios():
    # fileP = "DATAAGRO_Ferias/files/consolidados_precios/*.xlsx"
    fileP = "files/consolidados_precios/*.xlsx"
    filesP = glob.glob(fileP)

    archivosP = np.array(filesP)
    # finalP = "DATAAGRO_Ferias/consolidado/consolidado_feria_precios.xlsx"
    finalP = "consolidado/consolidado_feria_precios.xlsx"
    # print(len(archivos))
    
    if (os.path.isfile(finalP)):
        os.remove(finalP)
    else: 
        workbookP = xlsxwriter.Workbook(finalP)
        workbookP.close()
    
    '''try:
        os.remove(finalP)
    except:
        workbookP = xlsxwriter.Workbook(finalP)
        workbookP.close()'''
        
    for i in range(len(filesP)):
        
        df_inicialP = pd.read_excel(finalP)
        df_inicialP

        if(str(archivosP[i])!=finalP):
            dfP = pd.read_excel(archivosP[i])

            nP = df_inicialP.append([dfP])
            nP.to_excel(finalP, index=False)

def consolidarCantidad():
    # fileC = "DATAAGRO_Ferias/files/consolidados_cantidad/*.xlsx"
    fileC = "files/consolidados_cantidad/*.xlsx"
    filesC = glob.glob(fileC)

    archivosC = np.array(filesC)
    # finalC = "DATAAGRO_Ferias/consolidado/consolidado_feria_cantidad.xlsx"
    finalC = "consolidado/consolidado_feria_cantidad.xlsx"
    # print(len(archivos))
    
    '''if (os.path.isfile(finalC)):
        os.remove(finalC)
    else: 
        workbookC = xlsxwriter.Workbook(finalC)
        workbookC.close()'''
    
    try:
        os.remove(finalC)
    except:
        workbookC = xlsxwriter.Workbook(finalC)
        workbookC.close()
        
    for i in range(len(filesC)):
        
        df_inicialC = pd.read_excel(finalC)
        df_inicialC

        if(str(archivosC[i])!=finalC):
            dfC = pd.read_excel(archivosC[i])

            nC = df_inicialC.append([dfC])
            nC.to_excel(finalC, index=False)

def tabla3(namefile):
    book = xlrd.open_workbook(namefile)
    sh = book.sheet_by_index(3) 
    val = sh.cell_value(1,0)
    
    df3 = pd.read_excel(namefile, sheet_name="Número de cabezas", skiprows=5)
    df3["Detalle"] = val
    df3.columns = ['Feria','Comuna','Fecha', 'Novillo Gordo', 'Novillo Engorda', 'Vaca Gorda', 'Vaca Engorda', 'Vaquilla Gorda', 'Vaquilla Engorda', 'Toros', 'Terneros', 'Terneras', 'Cerdos', 'Lanares', 'Caballos', 'Detalle']
    df3 = df3[:-1]
    
    return df3

def cantidadArchivos():
    # file = "DATAAGRO_Ferias/files/*.xlsx"
    file = "files/*.xlsx"
    files = glob.glob(file)

    archivos = np.array(files)
    
    return len(archivos)

def descargarDatos():
    #while(True):
    archivos = cantidadArchivos() + 6 # Se suman 6 por las filas que no son tomadas en cuenta al momento de la descarga

    hoy = datetime.today().strftime('%A')
    driver = getDriver(enlace)
    time.sleep(15)

    _enlaces = driver.find_elements_by_xpath("/html/body/div[7]/div/div/div/div/div/article/div[2]/div/div/div[2]/table/tbody/tr")

    # IMPLEMENTACIÓN DEL CRON EN ANCTION

    if (archivos < len(_enlaces)):

        for i in range(len(_enlaces)):
            try:
                semana = driver.find_element_by_xpath("/html/body/div[7]/div/div/div/div/div/article/div[2]/div/div/div[2]/table/tbody/tr[" + str(i + 3) +"]/td[1]")
                descarga = driver.find_element_by_xpath("/html/body/div[7]/div/div/div/div/div/article/div[2]/div/div/div[2]/table/tbody/tr[" + str(i + 3) +"]/td[3]/a")

                time.sleep(1)

                # print(semana.text)
                # print(url)

                url = descarga.get_attribute("href")
                file = requests.get(url)
                # open("DATAAGRO_Ferias/files/" + str(semana.text) + ".xlsx", "wb").write(file.content)
                open("files/" + str(semana.text) + ".xlsx", "wb").write(file.content)


            except:
                pass

        driver.close()
        filesData()
        # print("Datos descargados correctamente.")
        #time.sleep(85800)

    else:
        driver.close()
            #time.sleep(86400)

if __name__ == '__main__':
    descargarDatos()