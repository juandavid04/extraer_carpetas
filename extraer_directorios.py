import os
import pandas as pd
import shutil
import datetime
from zipfile import ZipFile

dt = datetime.datetime.now()

file = input('Digite el nombre del archivo excel: Ejemplo: nombre_archivo.xlsx \n')

try:
    listDirExcel = pd.read_excel(file)                   #Se lee el archivo excel con el listado de carpetas a buscar
    dataframe = pd.DataFrame(listDirExcel)               #Se convierte el archivo en un DataFrame
    print(dataframe.head())
    
    columna = input('¿Cual columna contiene el nombre de los archivos? \n')
    carpetas_dt = list(dataframe[columna].copy())     #Se realiza una copia de la columna con el listado de carpetas
except:
    print('Ha ocurrido un error. Vuelve a intentarlo.')


directorios = []

try:
    with os.scandir() as ficheros:                       #Se lee el listado de carpetas de la ubicación actual donde se halla este archivo de python

        carpetas_dir = [fichero for fichero in ficheros if fichero.is_dir()]

        if not os.path.exists('../backup'):              #Se realiza un backup de la información
            os.mkdir('../backup')

        cs = "Copia de seguridad {}-{}-{}-{}-{}-{}".format(dt.second,dt.minute,dt.hour,dt.day,dt.month,dt.year)

        shutil.make_archive(cs,'zip','.')
        shutil.move('{}.zip'.format(cs),'../backup')
        print('Se creo copia de seguridad en ../backup/{}-{}-{}-{}-{}-{}.zip exitosamente... \n'.format(dt.second,dt.minute,dt.hour,dt.day,dt.month,dt.year))


        #Se comparan las carpetas con los nombres almacenados en el archivo excel y se copian en una nueva carpeta

        resultados = '../resultados_{}-{}-{}-{}-{}-{}'.format(dt.second,dt.minute,dt.hour,dt.day,dt.month,dt.year)

        os.mkdir(resultados)
        
        resultados_len = 0

        for directorio in carpetas_dir:

            dir_name = directorio.name.split('_')

            if(len(dir_name) == 4):

                for dt in carpetas_dt:   

                    if dir_name[2] == dt[:4] and dir_name[3] == dt[-6:]:

                        shutil.make_archive(directorio.name,'zip',directorio.path)
                        shutil.move('{}.zip'.format(directorio.name),resultados)

                        #with ZipFile('{}/{}.zip'.format(resultados,directorio.name), 'r') as zipfile:
                        #    print(zipfile)
                        #    zipfile.extractall()
                        
                        resultados_len = resultados_len + 1

                        print('se copio el directorio {} en resultados'.format(directorio.name))
        print()
        print('Resultados encontrados: {}'.format(resultados_len))
except:
    print('Ha ocurrido un error. Vuelve a intentarlo.')