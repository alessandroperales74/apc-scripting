# Cómo exportar un archivo de Excel desde SAP con Python y cargarlo a una Base de Datos Local en SQLite

Cuando comencé a hacer análisis de datos con transacciones de SAP, necesitaba organizar y estructurar mis datos con el tiempo. Lamentablemente, no tenía un servidor disponible para cargar los datos, y las opciones en la nube de SharePoint eran bastante limitadas para mis necesidades.

Fue en ese momento que descubrí SQLite y decidí explorarla. Aunque no es ideal depender de una base de datos local a largo plazo, es una buena manera de comenzar en el campo.

Si no tienes la infraestructura necesaria debido a circunstancias fuera de tu control, aún puedes hacer cosas interesantes con SQLite. En mi caso, pude trabajar y crear informes en Power BI utilizando conectores ODBC (al menos hasta que tuviera un servidor para cargar mis datos) o manejar una cantidad de filas que no habría sido posible utilizando Excel.

Aun así, fue mucho más fácil migrar datos que ya estaban estructurados y organizados de antemano. Así que, si no tienes un servidor para manejar este tipo de datos, no es una mala idea usar SQLite (que es gratuita y muy fácil de usar) para comenzar a trabajar en tus proyectos.

En este caso, voy a mostrar cómo hice un ETL muy simple pero efectivo para extraer datos de SAP y cargarlos en SQLite.

## 1. Importando Librerías Requeridas

```python
import win32com.client # Para poder hacer la conexión con SAP y Excel
import os # No es obligatorio, pero te da algo de flexibilidad con las rutas de las carpetas
import pandas as pd # Clásico de Pandas
import sqlite3
```

## 2. Exportación desde SAP

Ahora, necesito hacer la conexión con SAP. No voy a entrar en detalles sobre este código en particular porque quiero explicarlo más adelante en otra parte del repositorio, ya que es un tema bastante interesante y quiero darle el foco correspondiente.

```python
def sap_connection():

    print('Conectando con SAP...')
    SapGuiAuto  = win32com.client.GetObject("SAPGUI") # Haciendo la conexión con SAP mediante Python
    application = SapGuiAuto.GetScriptingEngine 
    connection = application.Children(0) # Primera conexión disponible
    session    = connection.Children(0) # Primera ventana disponible

    return session
```

Luego, declararé algunas variables necesarias para nuestras funciones posteriores. Necesito las entradas requeridas para la transacción. En este caso, lo intentaré con FBL1H. Puedes hacerlo con cualquiera que desees. Solo asegúrate de tener TODOS LOS DATOS REQUERIDOS.
```python

fbl1h_variant = 'APC_DB_SQLITE' # Tiene la configuración requerida de las entradas no variables
fbl1h_layout = 'APC_DB_SQLITE' # Esta es la distribución de columnas de las transacciones. Es una buena práctica tener una específicamente para esta carga

fbl1h_filepath = 'D:\\sqlite_db\\inputs'
fbl1h_filename = '202407_FBL1H.xlsx'

first_date = '01.07.2024'
last_date = '31.07.2024'

def fbl1h_contabilizadas(session):
    print('Exportando FBL1H...')
    session.StartTransaction("FBL1H")
    session.findById("wnd[0]/tbar[1]/btn[17]").press() 
    session.findById("wnd[1]/usr/txtV-LOW").text = fbl1h_variant
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]").sendVKey(8)

    # Puedes hacer que estas fechas sean dinámicas. Puedes usar una entrada o una función para obtener la primera y última fecha del mes anterior
    session.findById("wnd[0]/usr/ctxtS_PDATE-LOW").text = first_date
    session.findById("wnd[0]/usr/ctxtS_PDATE-HIGH").text = last_date

    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/chkP_TY_SPG").selected = True
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").text = fbl1h_layout
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/shellcont/shell").pressToolbarButton("SHOWBUT")
    session.findById("wnd[0]/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")

    # En este caso particular,
    # Estoy bien con los archivos de Excel debido al tamaño del informe. No es demasiado grande, así que es manejable
    session.findById("wnd[0]/shellcont/shell").selectContextMenuItem("&XXL")

    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").text = fbl1h_filename
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/cmbGS_EXPORT-FORMAT").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[20]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = fbl1h_filepath
    session.findById("wnd[1]").sendVKey(11) 
```
¡ATENCIÓN❗❗❗

En este caso particular, estoy bien con los archivos de Excel debido al tamaño del informe PERO debo advertir:

- **Si deseas exportar en un archivo de Excel:** A menudo tarda un poco más y Excel tiende a abrirse automáticamente después de exportar, por lo que -a veces- tienes que hacer un paso adicional para cerrarlo antes de leerlo con Pandas -en este caso, no he tenido problemas pero los tuve en otros scripts-. Además, considera el tiempo de espera. Por ejemplo, si tu informe tiene 500k filas, puedes exportarlo técnicamente (ya que el límite de filas en Excel es alrededor de 1 millón de filas), pero SAP puede fallar debido al tiempo de espera.

- **Si deseas exportar en un archivo txt/csv:** Es más rápido y no se abre automáticamente, pero es un poco complicado si encuentras algunos saltos de línea o tienes una columna de texto con el carácter '|' ya que ese carácter es el delimitador predeterminado de SAP. En una columna de 'Texto', a menudo encuentro este '||' 😒 ... Dicho esto: Ten cuidado.

## 3. Transformación del Archivo de Excel y Carga en la Base de Datos

Ahora que tengo el archivo de Excel exportado, tengo que cargarlo en mi Base de Datos.

Para este caso, tengo una función para transformar mi informe FBL1H. Puedes hacer cualquier transformación que desees.

```python

# Declarar ruta del archivo db
sqlite_path = 'D:\\database\\test_database.db'

# Conexión a SQLite
conn = sqlite3.connect(sqlite_path)
table_name = 'fbl1h'

def file_transformation():

    def convert_float(column):
        # SAP exporta números con comas y hace un poco complicado transformar la columna
        column_str = str(column).replace(',','')
        column_num = pd.to_numeric(column_str,errors='coerce')
        return column_num
    
    def strip_columns(df):
        # Solo por si acaso. Si lees un txt o csv de SAP, es obligatorio limpiar las columnas
        for columna in df.columns:
            df[columna] = df[columna].str.strip()

    # Leyendo mi Excel
    # Aquí uso la función os para tener mi ruta completa del Excel
    df_excel = pd.read_excel(os.path.join(fbl1h_filepath,fbl1h_filename),dtype=str)

    # A partir de aquí, esto está adaptado a mis necesidades. Puedes reemplazar esto con cualquier transformación que requieras. Lo dejo aquí como referencia.

    strip_columns(df_excel)

    # Estos encabezados deben coincidir con el esquema de la tabla en SQLite.
    # Agregaré una columna de período más adelante (que no está en SAP), pero aparte de eso, así es como creé mi tabla.

    headers = [
    'cod_sociedad',
    'cod_cuenta_mayor',
    'cod_proveedor',
    'nombre_proveedor',
    'referencia',
    'cod_documento_compras',
    'posicion',
    'cod_moneda_sociedad',
    'saldo_moneda_documento',
    'saldo_moneda_sociedad',
    'fecha_contabilizacion',
    'fecha_documento',
    'fecha_base'
    ]

    # Cambiar encabezados de SAP 
    df_excel.columns = headers 

    # Convertir columnas de fecha
    for column in headers:
        if 'fecha' in column:
            df_excel[column] = pd.to_datetime(df_excel[column],format='%d.%m.%Y',errors='coerce').dt.date

    # Convertir columnas numéricas
    numerical_columns = (
        'saldo_moneda_documento','saldo_moneda_sociedad'
    )

    for column in numerical_columns:
        df_excel[column] = df_excel[column].apply(convert_float)

    df_excel['cod_periodo'] = fbl1h_filename[0:6]
    df_excel['cod_periodo'] = df_excel['cod_periodo'].apply(convert_float)

    # Reordenar columnas para que coincidan con el esquema de la tabla
    df_excel = df_excel[['cod_periodo'] + headers]

    # Cargando datos a la Base de Datos de SQLite
    df_excel.to_sql(table_name, conn, if_exists='append', index=False)

# Llamada a funciones
session = sap_connection()
fbl1h_export(session)
file_transformation()

# Cerrar conexión SQLite
conn.close()
```

Finalmente, debes tener en cuenta que SQLite no tiene todas las características que tienen SQL Server, PostgreSQL, etc., pero es una manera fácil de empezar. Cuando tengas más conocimiento, puedes migrar a un motor de base de datos más robusto.

Puedes consultar el archivo Python en este enlace: ![SQLite_Python](python_scripts/sqlite_python.py)

¡Espero que te resulte útil! 🙋‍♂️🙋‍♂️🙋‍♂️🙋‍♂️
