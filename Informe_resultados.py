from openpyxl import load_workbook

# Ruta del formato original y vacío del informe de resultados
ruta_formato_original = r"D:\Desktop\Martin_Anaya\03_Muestreo_Superficial\02_Visual_SC\01_Input\DPT-F23 Informe de resultados V03 - propuesta.xlsx"
# Ruta donde se guarda el informe modificado y diligenciado
ruta_formato_modificado = r"D:\Desktop\Martin_Anaya\03_Muestreo_Superficial\02_Visual_SC\02_Output"
# Ruta del archivo que contiene los datos para el diccionario
ruta_archivo_fuente_1 = r"D:\Desktop\Martin_Anaya\03_Muestreo_Superficial\02_Visual_SC\01_Input\DPT-F11MuestreoDeAguaSuperficial-034.xlsx"

# Cargar el formato de Excel original
libro_original = load_workbook(ruta_formato_original)
# Leer archivo fuente
libro_fuente_1 = load_workbook(ruta_archivo_fuente_1)

# Modificar archivo fuente
hoja_fuente_1 = libro_fuente_1.active 

#Contador para el número de filas llenas
filas_llenas = 0

#Iteración sobre las y conteo de las filas llenas
for fila in hoja_fuente_1.iter_rows():
    if not all(cell is None for cell in fila):
        filas_llenas += 1
print()
print("Número de puntos de muestreo:", filas_llenas - 1)
print("Desde el número 2 hasta el número", filas_llenas)
print()

datos_encabezado = {"Proyecto":"D","Otro_proyecto":"E","Centro_costos":"F","Numero_solicitud":"G","Plan_muestreo":"G","Responsable_1":"K","Responsable_2":"L"}
celdas_encabezado = ["I7","I8","I9","H3","I10","I11","I12"]
datos_parametros = {"Departamento":"H","Municipio":"I","Fecha":"M","Hora":"N","Codigo_punto":"O","Nombre_fuente":"R","pH":"AL","Temperatura":"AM","OD":"AP","Conductividad_E":"AV","Caudal":"BG"}
celdas_parametros = []

print(len(datos_encabezado))
print(len(celdas_encabezado))

nombre = "\Informe_llenado_"

for fila in range(2,filas_llenas + 1):
    
    # Crear copia del archivo original
    libro_modificado = load_workbook(ruta_formato_original)
    # Modificar copia del archivo original
    hoja_modificada = libro_modificado.active 

    lista_keys = list(datos_encabezado.keys())             #Extraer una lista con las "claves" del diccionario
    for dato in range(len(datos_encabezado)):
    
        key = lista_keys[dato]                  #Definir cada una de las claves, posición dato de la lista lista_keys
        pos_celda = datos_encabezado[key] + str(fila)      #Obtener la posición de la celda y cocnatenar con el número de la fila
        valor = hoja_fuente_1[pos_celda].value
    
        hoja_modificada[celdas_encabezado[dato]] = valor
    
    ruta_ultima = ruta_formato_modificado+nombre+str(fila)+".xlsx"
    libro_modificado.save(ruta_ultima)
    libro_modificado.close()
        
libro_original.close()        
    