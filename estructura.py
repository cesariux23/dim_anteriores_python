import os
import glob
from openpyxl import load_workbook, worksheet, utils
meses = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC']
quincenas = ['1RA', '2DA']
directorios = []
ruta = os.getcwd()
a = 2014
for mes in meses:

    for quincena in quincenas:
        carpeta = '{0}/{1}'.format(mes, quincena)
        #inicial = '{0}/{1}'.format(q,m)
        #os.makedirs('{0}/{1}'.format(a,carpeta))
        files = glob.glob(('{0}/{1}/{2}/*.xlsx'.format(ruta, a, carpeta)))
        leer_libro(files, None)
    
def leer_libro (archivos, libro):
    for file in files:
        print(file)

