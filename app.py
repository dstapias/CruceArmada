import streamlit as st
import pandas as pd
import re
import openpyxl
from openpyxl.styles import Alignment, Font

# Título de la aplicación
st.title('Consolidación de Archivos Financieros')

# Carga de archivos desde la interfaz de Streamlit
archivo_siif = st.file_uploader("📄 Cargar archivo SIIF (Excel)", type=["xlsx"])
archivo_reglas = st.file_uploader("📄 Cargar archivo de reglas (Excel)", type=["xlsx"])
archivo_directorio = st.file_uploader("📄 Cargar archivo del directorio (Excel)", type=["xlsx"])

# Continuar solo si se cargan todos los archivos
if archivo_siif and archivo_reglas and archivo_directorio:
    st.success("✅ ¡Archivos cargados con éxito! Procesando...")
    # Leer todas las hojas del archivo principal
    hojas = pd.ExcelFile(archivo_siif).sheet_names

    # Lista para almacenar los DataFrames de cada hoja
    dfs = []

    # Columnas esperadas del archivo principal
    columnas_esperadas = [
        'Identificacion', 'Descripcion', 'Saldo Anterior',
        'Movimientos Debito', 'Movimientos Credito', 'Saldo Final'
    ]

    # Expresión regular para extraer el número de Código Contable
    codigo_pattern = re.compile(r'Codigo Contable\s*(\d+)')

    # Iterar sobre cada hoja del archivo principal
    for hoja in hojas:
        # Cargar la hoja sin encabezado para detectar dónde comienzan los datos
        df_temp = pd.read_excel(archivo_siif, sheet_name=hoja, header=None)

        # Inicializar el valor de código contable
        codigo_contable = None

        # Buscar el número de Código Contable antes del encabezado
        for i, row in df_temp.iterrows():
            row_text = ' '.join(row.dropna().astype(str))  # Convertir fila a texto
            match = codigo_pattern.search(row_text)
            if match:
                codigo_contable = match.group(1)  # Extraer el código contable
                break

        # Buscar el encabezado correcto basado en las columnas esperadas
        for i, row in df_temp.iterrows():
            if set(columnas_esperadas).issubset(set(row.dropna().tolist())):
                header_index = i
                break

        # Leer nuevamente la hoja con el encabezado correcto
        df = pd.read_excel(archivo_siif, sheet_name=hoja, header=header_index)

        # Filtrar solo las columnas esperadas
        df = df[columnas_esperadas]

        # Eliminar la última fila (que tiene los totales)
        df = df[:-1]

        # Añadir el código contable como una columna
        df['Codigo Contable'] = codigo_contable

        df['Codigo Sin 001'] = codigo_contable[:-3]

        # Crear una nueva columna NIT eliminando el prefijo 'TER '
        df['NIT'] = df['Identificacion'].str.replace(r'^TER\s*', '', regex=True)

        # Agregar el DataFrame de la hoja a la lista
        dfs.append(df)

    # Concatenar todos los DataFrames en uno solo
    df_final = pd.concat(dfs, ignore_index=True)

    # Cargar el archivo con la hoja 'Cuentas al 100%'
    df_nuevo_temp = pd.read_excel(archivo_reglas, sheet_name='Cuentas al 100%', header=None)

    # Buscar el encabezado correcto del archivo de reglas
    for i, row in df_nuevo_temp.iterrows():
        if {'Código', 'Descripción', 'Reportable al 100%'}.issubset(set(row.dropna().tolist())):
            header_index_nuevo = i
            break

    # Leer nuevamente la hoja con el encabezado correcto
    df_nuevo = pd.read_excel(archivo_reglas, sheet_name='Cuentas al 100%', header=header_index_nuevo)

    # Eliminar puntos de la columna 'Código' para el cruce
    df_nuevo['Codigo Limpiado'] = df_nuevo['Código'].astype(str).str.replace('.', '', regex=False)

    # Cruzar los archivos por NIT y Código Limpiado
    df_final = pd.merge(
        df_final,
        df_nuevo[['Codigo Limpiado', 'Descripción', 'Código']],
        left_on='Codigo Sin 001',
        right_on='Codigo Limpiado',
        how='inner'
    )

    # Eliminar columna auxiliar después del cruce
    df_final.drop(columns=['Codigo Limpiado'], inplace=True)

    # Filtrar solo registros donde el código empiece con 1, 2, 4 o 5
    df_final = df_final[df_final['Código'].str.startswith(('1', '2', '4', '5'))]

    ### Cargar el nuevo archivo para hacer el cruce con Directorio
    hojas_directorio = pd.ExcelFile(archivo_directorio).sheet_names

    # Buscar la hoja que contenga la palabra 'Directorio'
    hoja_directorio = [hoja for hoja in hojas_directorio if 'Directorio' in hoja][0]

    # Leer el archivo de directorio sin encabezado para detectar el encabezado correcto
    df_directorio_temp = pd.read_excel(archivo_directorio, sheet_name=hoja_directorio, header=None)

    # Buscar el encabezado correcto para el directorio
    for i, row in df_directorio_temp.iterrows():
        if {'Id Entidad', 'Nit', 'Razón Social', 'Departamento', 'Municipio', 'Dirección', 'Código Postal', 'Teléfono', 'Fax', 'e-mail', 'Página Web', 'Ámbito Convergencia', 'Ambito SIIF'}.issubset(set(row.dropna().tolist())):
            header_index_directorio = i
            break

    # Leer nuevamente la hoja del directorio con el encabezado correcto
    df_directorio = pd.read_excel(archivo_directorio, sheet_name=hoja_directorio, header=header_index_directorio)

    # Eliminar todo después de ':' en la columna 'Nit'
    df_directorio['Nit Limpiado'] = df_directorio['Nit'].astype(str).str.split(':').str[0]

    # Cruzar por NIT
    df_final = pd.merge(
        df_final,
        df_directorio[['Nit Limpiado', 'Id Entidad', 'Razón Social']],
        left_on='NIT',
        right_on='Nit Limpiado',
        how='inner'
    )

    # Eliminar columna auxiliar después del cruce
    df_final.drop(columns=['Nit Limpiado'], inplace=True)


    ### Verificar duplicados solo de los NIT del resultado final
    nits_finales = df_final['NIT'].unique()  # Obtener los NIT del resultado final
    df_directorio_filtrado = df_directorio[df_directorio['Nit Limpiado'].isin(nits_finales)]  # Filtrar solo los NIT finales

    # Verificar si alguno de los NIT finales está repetido
    nits_repetidos = df_directorio_filtrado['Nit Limpiado'].value_counts()
    nits_duplicados = nits_repetidos[nits_repetidos > 1].index.tolist()
    
    ### Verificar duplicados solo de los NIT del resultado final
    cods_finales = df_final['Código'].unique()  # Obtener los NIT del resultado final
    df_reglas_filtrado = df_nuevo[df_nuevo['Código'].isin(cods_finales)]  # Filtrar solo los NIT finales

    # Verificar si alguno de los NIT finales está repetido
    cods_repetidos = df_reglas_filtrado['Código'].value_counts()
    cods_duplicados = cods_repetidos[cods_repetidos > 1].index.tolist()

    # Verificar si hay duplicados en NIT o Código
    errores_encontrados = False

    if nits_duplicados:
        errores_encontrados = True
        for nit in nits_duplicados:
            st.warning(f"⚠️ El NIT {nit} está repetido. Deja solo uno en el archivo de Directorio")

    if cods_duplicados:
        errores_encontrados = True
        for cod in cods_duplicados:
            st.warning(f"⚠️ El Código {cod} está repetido. Deja solo uno en el archivo de Reglas")

    # Ordenar el DataFrame por 'CODIGO CONTABLE' de forma ascendente
    df_final = df_final.sort_values(by="Codigo Sin 001", ascending=True)

    # Seleccionar y renombrar las columnas específicas
    columnas_seleccionadas = {
        "Código": "CODIGO CONTABLE",
        "Descripción": "NOMBRE",
        "Id Entidad": "CODIGO ENTIDAD",
        "Razón Social": "NOMBRE ENTIDAD",
        "Saldo Final": "SALDO FINAL"  # Renombramos temporalmente para procesar
    }

    # Filtrar las columnas deseadas y renombrar encabezados
    df_final = df_final[list(columnas_seleccionadas.keys())]  # Filtrar solo las columnas necesarias
    df_final.rename(columns=columnas_seleccionadas, inplace=True)

    # ======= ASIGNAR VALORES A LAS COLUMNAS =======
    # Crear columnas 'VALOR NO CORRIENTE' y 'VALOR CORRIENTE' inicializadas en 0
    df_final["VALOR NO CORRIENTE"] = 0.0
    df_final["VALOR CORRIENTE"] = 0.0

    # Asignar valores según el inicio del 'CODIGO CONTABLE'
    df_final.loc[df_final["CODIGO CONTABLE"].astype(str).str.startswith(("1", "2")), "VALOR NO CORRIENTE"] = df_final["SALDO FINAL"]
    df_final.loc[df_final["CODIGO CONTABLE"].astype(str).str.startswith(("4", "5")), "VALOR CORRIENTE"] = df_final["SALDO FINAL"]

    # Eliminar la columna 'SALDO FINAL' después de procesar
    df_final.drop(columns=["SALDO FINAL"], inplace=True)

    # Si no hay errores, permitir la descarga del archivo
    if not errores_encontrados:
        st.success("✅ No hay duplicados. El archivo está listo para descargar.")

        # Guardar el archivo Excel en disco
        archivo_salida = "/tmp/consolidado_final.xlsx"  # Ruta temporal

        with pd.ExcelWriter(archivo_salida, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False, startrow=8, sheet_name="Consolidado")  # Mover datos una fila hacia abajo
            # Obtener el libro y la hoja activa
            workbook = writer.book
            worksheet = writer.sheets["Consolidado"]

            # ======= COMBINAR CELDAS Y AGREGAR TEXTO =======
            title_text = "MODELO CGN2005_002_OPERACIONES_RECIPROCAS"
            worksheet.merge_cells('A1:F1')  # Combinar columnas A-F en la fila 1
            title_cell = worksheet['A1']
            title_cell.value = title_text

            # ======= CENTRAR TEXTO =======
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # ======= AGREGAR DATOS SIN FORMATO EN LAS FILAS 3 A 7 =======
            datos_adicionales = [
                ("DEPARTAMENTO", "CUNDINAMARCA"),
                ("MUNICIPIO", "BOGOTÁ D.C."),
                ("ENTIDAD", 'ARMADA NACIONAL DE COLOMBIA - BASE NAVAL No. 6 ARC "Bogotá"'),
                ("CODIGO", "11100000"),
                ("FECHA DE CORTE:", "31 de Marzo de 2021")
            ]

            # Insertar los datos en las filas 3 a 7
            fila_inicio = 3
            for i, (col1, col2) in enumerate(datos_adicionales, start=fila_inicio):
                worksheet[f"A{i}"] = col1
                worksheet[f"B{i}"] = col2
            
            # ======= AGREGAR TEXTO Y COMBINAR CELDAS EN LA FILA 8 =======
            texto_fila_8 = "INFORMACION SOBRE SALDOS DE OPERACIONES RECIPROCAS"
            worksheet.merge_cells('A8:C8')  # Combinar columnas A-C en la fila 8
            cell_fila_8 = worksheet['A8']
            cell_fila_8.value = texto_fila_8

            # Aplicar negrita al texto de la fila 8
            cell_fila_8.font = Font(bold=True)

        # Abrir el archivo para descarga
        with open(archivo_salida, "rb") as file:
            st.download_button(
                label="📥 Descargar Consolidado Final",
                data=file,
                file_name="consolidado_final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("❌ No se puede generar el archivo debido a duplicados. Corrige los errores antes de continuar.")


    
else:
    st.info("⏳ Por favor, carga los tres archivos para comenzar el proceso.")
