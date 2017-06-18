import os

meses = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC']
quincenas = ['1RA', '2DA']
directorios = []

a=2014
for m in meses:
    for q in quincenas:
        carpeta = '{0}/{1}'.format(m,q)
        inicial = '{0} {1}'.format(q,m)
        os.makedirs('{0}/{1}'.format(a,carpeta))
