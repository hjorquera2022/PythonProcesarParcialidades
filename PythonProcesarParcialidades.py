#PythonProcesarParcialidades.py

import pandas as pd
import os
import subprocess

#*****
#***** ESTRUCTURA COMPLETA 
#*****

#├───ESTRUCTURA DE CARPETAS DE CADA PARCIALIDAD
#│   ├───DOCUMENTOS APROBADOS
#│   │   ├───REV LETRA
#│   │   │   ├───01 PDF
#│   │   │   └───02 EDITABLE
#│   │   └───REV NUMERO
#│   │       ├───01 PDF
#│   │       └───02 EDITABLE
#│   └───DOCUMENTOS VIGENTES
#│       ├───01 PDF
#│       │   ├───CON OBSERVACIONES
#│       │   └───SIN OBSERVACIONES
#│       └───02 EDITABLE
#│           ├───CON OBSERVACIONES
#│           └───SIN OBSERVACIONES



# Ruta base donde se deben verificar los subdirectorios
ruta_base = 'R:\\01 PARCIALIDADES\\'  ### REAL

# Nombre del archivo de log
archivo_log = ruta_base + '0000-00 ADMINISTRACION\\LOG\\log_ProcesarParcialidades.txt'

# Nombre del archivo Bat General
archivo_bat = ruta_base + '0000-00 ADMINISTRACION\\BAT\\Bat_ProcesarParcialidades.bat'

# Planilla con la lista de parcialidades
archivo_excel = ruta_base + 'Listado de Parcialidades.xlsx'

# Carga el archivo Excel en un DataFrame Hoja de Parcialidades.
df = pd.read_excel(archivo_excel, sheet_name='PARCIALIDADES')

# Filtra el DataFrame para considerar solo parcialidades a 'PROCESAR' igual a 'S'
df_parcialidades = df[df['PROCESAR'] == 'S']

# Abre el archivo de log en modo de escritura
with open(archivo_log, 'w') as log_file:

# Abre el archivo de log en modo de escritura
 with open(archivo_bat, 'w') as bat_file:

    # Itera a través de cada parcialidad y la procesa
    for parcialidad in df_parcialidades['PARCIALIDAD']:
        log_file.write(f'Parcialidad: {parcialidad}\n')

        #******* Abrir Planilla CONTROL DOCUMENTOS ING DEF con las 8 hojas para traspasar a BAT
        #******* Cargar cada una de las 8 hojas del archivo Excel en un DataFrame (DFxxxx)
        #******* Generar cada uno de los BAT 

        #******* Armar nombre de archivo del BAT por cada uno de las 8 hojas
        # ACTUALIZA EDITABLE DOC VIG - RUTA
        # ACTUALIZA EDITABLE REV LETRA - RUTA
        # ACTUALIZA EDITABLE REV NUM - RUTA
        # ACTUALIZA PDF DOC VIG - RUTA
        # ACTUALIZA PDF REV LETRA - RUTA
        # ACTUALIZA PDF REV NUM - RUTA
        
        archivo_ACTUALIZA_EDITABLE_DOC_VIG = ruta_base + '0000-00 ADMINISTRACION\\BAT\\' + parcialidad + '_BAT_ACTUALIZA_EDITABLE_DOC_VIG.bat'
        archivo_ACTUALIZA_EDITABLE_REV_LETRA = ruta_base + '0000-00 ADMINISTRACION\\BAT\\' + parcialidad + '_BAT_ACTUALIZA_EDITABLE_REV_LETRA.bat'        
        archivo_ACTUALIZA_EDITABLE_REV_NUM = ruta_base + '0000-00 ADMINISTRACION\\BAT\\' + parcialidad + '_BAT_ACTUALIZA_EDITABLE_REV_NUM.bat'
        archivo_ACTUALIZA_PDF_DOC_VIG = ruta_base + '0000-00 ADMINISTRACION\\BAT\\' + parcialidad + '_BAT_ACTUALIZA_PDF_DOC_VIG.bat'
        archivo_ACTUALIZA_PDF_REV_LETRA = ruta_base + '0000-00 ADMINISTRACION\\BAT\\' + parcialidad + '_BAT_ACTUALIZA_PDF_REV_LETRA.bat'
        archivo_ACTUALIZA_PDF_REV_NUM = ruta_base + '0000-00 ADMINISTRACION\\BAT\\' + parcialidad + '_BAT_ACTUALIZA_PDF_REV_NUM.bat'        
        
      
         #******* 
        #******* RECORRER TODAS LAS PARCIALIDADES CONTANDO LOS ARCHIVOS DE LA SIGUIENTE ESTRUCTURA de las Carpetas en REVISORES  por cada parcialidad.
        #******* 

        parcialidad_0_7_10 = parcialidad[0:7]
        if parcialidad_0_7_10   == '0029-14':
            parcialidad_0_7_10 = parcialidad[0:10]
        elif parcialidad_0_7_10 == '032ESO-':
            parcialidad_0_7_10 = parcialidad[0:9]
        elif parcialidad_0_7_10 == '032ESP-':
            parcialidad_0_7_10 = parcialidad[0:9]

        archivo_parcialidad = ruta_base + parcialidad + '\\CONTROL DOCUMENTOS ING DEF P' + parcialidad_0_7_10 + '.xlsx'
        
        bat_file.write(f'CALL \"{archivo_ACTUALIZA_EDITABLE_DOC_VIG}\" > \"{ruta_base}0000-00 ADMINISTRACION\LOG\P{parcialidad}.log\" \n')
        bat_file.write(f'CALL \"{archivo_ACTUALIZA_EDITABLE_REV_LETRA}\" >> \"{ruta_base}0000-00 ADMINISTRACION\LOG\P{parcialidad}.log\" \n')
        bat_file.write(f'CALL \"{archivo_ACTUALIZA_EDITABLE_REV_NUM}\" >> \"{ruta_base}0000-00 ADMINISTRACION\LOG\P{parcialidad}.log\" \n')
        bat_file.write(f'CALL \"{archivo_ACTUALIZA_PDF_DOC_VIG}\" >> \"{ruta_base}0000-00 ADMINISTRACION\LOG\P{parcialidad}.log\" \n')
        bat_file.write(f'CALL \"{archivo_ACTUALIZA_PDF_REV_LETRA}\" >> \"{ruta_base}0000-00 ADMINISTRACION\LOG\P{parcialidad}.log\" \n')
        bat_file.write(f'CALL \"{archivo_ACTUALIZA_PDF_REV_NUM}\" >> \"{ruta_base}0000-00 ADMINISTRACION\LOG\P{parcialidad}.log\" \n')

        if not os.path.exists(archivo_parcialidad):
              log_file.write(f'Parcialidad: {parcialidad} SIN ARCHIVO DE INGENIERIA {archivo_parcialidad}\n')
        else:
                print(f'Procesando Parcialidad: {parcialidad} ARCHIVO:  {archivo_parcialidad}\n')

                # Nombre de hojas
                # ACTUALIZA EDITABLE DOC VIG
                # ACTUALIZA EDITABLE REV LETRA
                # ACTUALIZA EDITABLE REV NUM
                # ACTUALIZA PDF DOC VIG
                # ACTUALIZA PDF REV LETRA
                # ACTUALIZA PDF REV NUM


                #******* Cargar cada una de las 8 hojas del archivo Excel en un DataFrame (DF_xxxx)

                df_ACTUALIZA_EDITABLE_DOC_VIG = pd.read_excel(archivo_parcialidad, sheet_name='ACTUALIZA EDITABLE DOC VIG')
                df_ACTUALIZA_EDITABLE_REV_LETRA = pd.read_excel(archivo_parcialidad, sheet_name='ACTUALIZA EDITABLE REV LETRA')   
                df_ACTUALIZA_EDITABLE_REV_NUM = pd.read_excel(archivo_parcialidad, sheet_name='ACTUALIZA EDITABLE REV NUM')
                df_ACTUALIZA_PDF_DOC_VIG = pd.read_excel(archivo_parcialidad, sheet_name='ACTUALIZA PDF DOC VIG')
                df_ACTUALIZA_PDF_REV_LETRA = pd.read_excel(archivo_parcialidad, sheet_name='ACTUALIZA PDF REV LETRA')
                df_ACTUALIZA_PDF_REV_NUM = pd.read_excel(archivo_parcialidad, sheet_name='ACTUALIZA PDF REV NUM')
               

                #******* Generar cada uno de los BAT 
                
                # Abre el archivo BAT en modo de escritura
                log_file.write(f'Parcialidad: {parcialidad} BAT {archivo_ACTUALIZA_EDITABLE_DOC_VIG}\n')
                print(f'BAT {archivo_ACTUALIZA_EDITABLE_DOC_VIG}')
                with open(archivo_ACTUALIZA_EDITABLE_DOC_VIG, 'w') as bat_file_ACTUALIZA_EDITABLE_DOC_VIG:
                # Itera a través de la hoja por cada linea
                    for linea in df_ACTUALIZA_EDITABLE_DOC_VIG['RUTA']:
                        bat_file_ACTUALIZA_EDITABLE_DOC_VIG.write(f'{linea}\n')

                # Abre el archivo BAT en modo de escritura
                log_file.write(f'Parcialidad: {parcialidad} BAT {archivo_ACTUALIZA_EDITABLE_REV_LETRA}\n')
                print(f'BAT {archivo_ACTUALIZA_EDITABLE_REV_LETRA}')
                with open(archivo_ACTUALIZA_EDITABLE_REV_LETRA, 'w') as bat_file_ACTUALIZA_EDITABLE_REV_LETRA:
                # Itera a través de la hoja por cada linea
                    for linea in df_ACTUALIZA_EDITABLE_REV_LETRA['RUTA']:
                        bat_file_ACTUALIZA_EDITABLE_REV_LETRA.write(f'{linea}\n')

                # Abre el archivo BAT en modo de escritura
                log_file.write(f'Parcialidad: {parcialidad} BAT {archivo_ACTUALIZA_EDITABLE_REV_NUM}\n')
                print(f'BAT {archivo_ACTUALIZA_EDITABLE_REV_NUM}')
                with open(archivo_ACTUALIZA_EDITABLE_REV_NUM, 'w') as bat_file_ACTUALIZA_EDITABLE_REV_NUM:
                # Itera a través de la hoja por cada linea
                    for linea in df_ACTUALIZA_EDITABLE_REV_NUM['RUTA']:
                        bat_file_ACTUALIZA_EDITABLE_REV_NUM.write(f'{linea}\n')

                # Abre el archivo BAT en modo de escritura
                log_file.write(f'Parcialidad: {parcialidad} BAT {archivo_ACTUALIZA_PDF_DOC_VIG}\n')
                print(f'BAT {archivo_ACTUALIZA_PDF_DOC_VIG}')
                with open(archivo_ACTUALIZA_PDF_DOC_VIG, 'w') as bat_file_ACTUALIZA_PDF_DOC_VIG:
                # Itera a través de la hoja por cada linea
                    for linea in df_ACTUALIZA_PDF_DOC_VIG['RUTA']:
                        bat_file_ACTUALIZA_PDF_DOC_VIG.write(f'{linea}\n')

                # Abre el archivo BAT en modo de escritura
                log_file.write(f'Parcialidad: {parcialidad} BAT {archivo_ACTUALIZA_PDF_REV_LETRA}\n')
                print(f'BAT {archivo_ACTUALIZA_PDF_REV_LETRA}')
                with open(archivo_ACTUALIZA_PDF_REV_LETRA, 'w') as bat_file_ACTUALIZA_PDF_REV_LETRA:
                # Itera a través de la hoja por cada linea
                    for linea in df_ACTUALIZA_PDF_REV_LETRA['RUTA']:
                        bat_file_ACTUALIZA_PDF_REV_LETRA.write(f'{linea}\n')

                # Abre el archivo BAT en modo de escritura
                log_file.write(f'Parcialidad: {parcialidad} BAT {archivo_ACTUALIZA_PDF_REV_NUM}\n')
                print(f'BAT {archivo_ACTUALIZA_PDF_REV_NUM}')
                with open(archivo_ACTUALIZA_PDF_REV_NUM, 'w') as bat_file_ACTUALIZA_PDF_REV_NUM:
                # Itera a través de la hoja por cada linea
                    for linea in df_ACTUALIZA_PDF_REV_NUM['RUTA']:
                        bat_file_ACTUALIZA_PDF_REV_NUM.write(f'{linea}\n')
log_file.close
with open(archivo_log, 'a') as log_file:
    print("Proceso finalizado. Los resultados se han guardado en R:\01 PARCIALIDADES\0000-00 ADMINISTRACION\LOG en el archivo de log_ProcesarParcialidades.")
    log_file.write(f'Proceso finalizado. Los resultados se han guardado en R:\01 PARCIALIDADES\0000-00 ADMINISTRACION\LOG en el archivo de log_ProcesarParcialidades.\n')
log_file.close
bat_file.close

#try:
#    # Ejecuta el archivo batch
#    subprocess.run(archivo_bat, shell=True)
#except Exception as e:
#    print(f"Error al ejecutar el archivo batch: {e}")
#    log_file.write(f'Error en ejecucion de Bat_ProcesarParcialidades.bat\n')


