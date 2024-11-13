# This Python script performs the following tasks:
import pandas as pd
import os

# Verificar si el archivo existe en la ruta
ruta_origen = r'C:\Users\Constanza Perez\Documents\Contabilizacones 2022\Contabilizacion 2024\2306 - CLP_USD Compra-Venta.xlsx'

if os.path.exists(ruta_origen):
    print("El archivo existe.")
else:
    print("El archivo NO existe en la ruta especificada.")

# Genera la ruta de salida en la misma carpeta del archivo de origen
carpeta_salida = os.path.dirname(ruta_origen)
nombre_salida = os.path.join(carpeta_salida, 'mbi manager compra dolares.xlsx')

# Configuración de filtro de fecha
mes_filtro = 10  # Mes en formato numérico, ej. 10 para octubre
anio_filtro = 2024  # Año en formato numérico

# Lee la hoja específica del archivo de origen
df_origen = pd.read_excel(ruta_origen, sheet_name='COMPRA USD')

# Imprime las primeras filas del DataFrame para verificar las columnas
print("Primeras filas del DataFrame:")
print(df_origen.head())

# Asegúrate de que la columna de fechas esté en formato datetime
df_origen['Fecha'] = pd.to_datetime(df_origen.iloc[:, 2], errors='coerce')  # Accede a la columna de fechas

# Filtra las filas según las condiciones de texto y fecha
df_filtrado = df_origen[
    (df_origen.iloc[:, 3] == 'MBI CORREDORES DE BOLSA') &  # Columna D: "MBI CORREDORES DE BOLSA"
    (df_origen.iloc[:, 5] == 'BANCO BICE USD') &  # Columna F: "BANCO BICE USD"
    (df_origen['Fecha'].dt.month == mes_filtro) &
    (df_origen['Fecha'].dt.year == anio_filtro)
]

# Verifica cuántas filas quedaron después del filtro
print(f"Filas después de aplicar el filtro: {df_filtrado.shape[0]}")

# Si no hay filas después del filtro, imprimimos un mensaje y detenemos el código
if df_filtrado.shape[0] == 0:
    print("No hay filas que coincidan con los criterios.")
else:
    # Inicializa una lista para acumular las filas
    filas = []

    # Variables para rellenar valores en el archivo de salida
    codigo_cuenta_contableh = '110296'
    codigo_cuenta_contabled = '110295'


    # Itera sobre el DataFrame filtrado para rellenar el nuevo archivo
    for i, row in df_filtrado.iterrows():
        fecha_contable = row['Fecha'].strftime('%d/%m/%Y')
        monto = row.iloc[4]  # Columna E
        # Abreviación: valores en "K" sin decimales, en "M" con un decimal
        if row.iloc[6] >= 1000000:
            abreviado = f"{row.iloc[6] / 1000000:.1f} M"  # Un decimal para millones
        else:
            abreviado = f"{int(row.iloc[6] / 1000)}K"  # Sin decimales para miles
        tipo_cambio = str(row.iloc[7]).replace('.', ',')   # Columna H
        glosa_texto = f"CPA Compra Divisas {abreviado} TC {tipo_cambio}"
        
        correlativo = 1

        # Crear dos líneas para cada fecha
        for j in range(2):  # Dos líneas por fecha

             # La lógica para cambiar la cuenta contable
            if monto > 0 and j == 0:  # Si monto > 0, primera línea lleva la cuenta '110296'
                codigo_cuenta_contable = codigo_cuenta_contableh
                monto_debe = monto
                monto_haber = ''
            else:  # Si monto == 0, segunda línea lleva la cuenta '110295'
                codigo_cuenta_contable = codigo_cuenta_contabled
                monto_debe = ''
                monto_haber = monto
            

            linea = {
                'Tipo de comprobante': 'T',
                'Esquema': '',
                'Glosa comprobante': glosa_texto,
                'Fecha contable': fecha_contable,
                'Total debe': '',
                'Total haber': '',
                'Ítem detalle comprobante': correlativo,
                'Código unidad de negocio': 1,
                'Glosa comprobante': glosa_texto,
                'RUT cliente': '',
                'RUT personal': '',
                'Código cuenta contable': codigo_cuenta_contable,
                'Código centro costos': '',
                'Monto debe': monto if j == 0 else '',
                'Monto haber': '' if j == 0 else monto,
                'Tipo de documento': '',
                'Fecha de documento': '',
                'Número de documento': '',
                'Concepto 1': '',
                'Concepto 2': '',
                'Concepto 3': '',
                'Concepto 4': '',
                'Tipo contabilidad': ''
            }
            
            # Agrega la línea a la lista
            filas.append(linea)
            
            # Incrementa el correlativo solo después de las dos líneas
            correlativo += 1

    # Crea el DataFrame final con las filas acumuladas
    df_salida = pd.DataFrame(filas, columns=[
        'Tipo de comprobante', 'Esquema', 'Glosa comprobante', 'Fecha contable', 
        'Total debe', 'Total haber', 'Ítem detalle comprobante', 'Código unidad de negocio', 
        'Glosa comprobante', 'RUT cliente', 'RUT personal', 'Código cuenta contable', 
        'Código centro costos', 'Monto debe', 'Monto haber', 'Tipo de documento', 
        'Fecha de documento', 'Número de documento', 'Concepto 1', 'Concepto 2', 
        'Concepto 3', 'Concepto 4', 'Tipo contabilidad'
    ])

    # Guarda el DataFrame de salida en un archivo Excel en la misma carpeta que el archivo de origen
    df_salida.to_excel(nombre_salida, index=False)
    print(f"Archivo '{nombre_salida}' generado correctamente en la misma carpeta que el archivo de origen.")
