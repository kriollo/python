# -*- coding: utf-8 -*-

__author__ = 'Jorge Jara H'
__copyright__ = "Copyright 2022, RPA Project"
__license__ = "CC Attribution-NonCommercial-NoDerivs 4.0 International"
__version__ = 1.0
__maintainer__ = "Jorge Jara H"
__email__ = "jjara@wys.cl"
__status__ = "production"

from pickle import TRUE
import win32com.client as win32
import time
from datetime import datetime

PATH = 'C:/directorio_base/'

book_update = [
    "archivo1.xlsx",
    "carpeta/archivo2.xlsx",
    "archivo3.xlsx",
]

def update_book(book):
    print("Abriendo Libro: " + book)
    EXCEL = win32.DispatchEx('Excel.Application')
    INFORME = EXCEL.Workbooks.Open(PATH + book)
    EXCEL.Visible = False
    EXCEL.EnableEvents = False
    time.sleep(3)
    print("Inicio actualización: " + book)
    INFORME.RefreshAll()
    EXCEL.CalculateUntilAsyncQueriesDone()
    print("Fin actualización: " + book)
    INFORME.Save()
    INFORME.Close(True)
    EXCEL.Application.Quit()
    print("Libro Cerrado: " + book)
    print("")


print("Inciando update....")


for book in book_update:
    update_book(book)

print("Fin update.... Gracias por preferirnos...")
