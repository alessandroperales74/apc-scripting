import win32com.client # For being able to make the connection with SAP and Excel
import os # Not mandatory, but it gives you some flexibility with folder paths
import pandas as pd # Classic Pandas stuff 
import sqlite3

def sap_connection():

    print('Connecting with SAP...')
    SapGuiAuto  = win32com.client.GetObject("SAPGUI") # Making SAP connection with Python
    application = SapGuiAuto.GetScriptingEngine 
    connection = application.Children(0) # First connection available
    session    = connection.Children(0) # First window available

    return session

fbl1h_variant = 'APC_DB_SQLITE' # It has the required configuration of the non variable inputs
fbl1h_layout = 'APC_DB_SQLITE' # This is the column distribution of the transactions. It's a good practice to have one specifically for this load

fbl1h_filepath = 'D:\\sqlite_db\\inputs'
fbl1h_filename = '202407_FBL1H.xlsx'

first_date = '01.07.2024'
last_date = '31.07.2024'

def fbl1h_export(session):
    print('Exporting FBL1H...')
    session.StartTransaction("FBL1H")
    session.findById("wnd[0]/tbar[1]/btn[17]").press() 
    session.findById("wnd[1]/usr/txtV-LOW").text = '/FBL1H_BBDD'
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]").sendVKey(8)

    # You can make these dates dynamic. You can use an input or a function to get the first and last date of the previous month
    session.findById("wnd[0]/usr/ctxtS_PDATE-LOW").text = first_date
    session.findById("wnd[0]/usr/ctxtS_PDATE-HIGH").text = last_date

    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/chkP_TY_SPG").selected = True
    #session.findById("wnd[0]/usr/ctxtP_LAYOUT").text = fbl1h_layout
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/shellcont/shell").pressToolbarButton("SHOWBUT")
    session.findById("wnd[0]/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")

    # In this particular case,
    # I'm fine with Excel files due to the size of the report. It's not too large so it's cool
    session.findById("wnd[0]/shellcont/shell").selectContextMenuItem("&XXL")

    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").text = fbl1h_filename
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/cmbGS_EXPORT-FORMAT").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[20]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = fbl1h_filepath
    session.findById("wnd[1]").sendVKey(11) 


# Declare db file path
sqlite_path = 'D:\\database\\test_database.db'

# SQLite connection
conn = sqlite3.connect(sqlite_path)

table_name = 'fbl1h'

def file_transformation():

    def convert_float(column):
        # SAP export numbers with commas and make a bit tricky to transform the column
        column_str = str(column).replace(',','')
        column_num = pd.to_numeric(column_str,errors='coerce')
        return column_num
    
    def strip_columns(df):
        # Just in case. If you read a SAP txt or csv, is mandatory to strip the columns
        for columna in df.columns:
            df[columna] = df[columna].str.strip()

    # Reading my Excel
    # Here I use the os function join to have my full Excel path
    df_excel = pd.read_excel(os.path.join(fbl1h_filepath,fbl1h_filename),dtype=str)

    # From now on, this is my particular need. You can replace this code with anything you need. I'm going to leave it, so you have an idea

    strip_columns(df_excel)

    # These headers have to be exactly the same as your table in SQLite.
    # I'll add later a period column (which it's not in SAP) but, besides that, this is how I created my table

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

    # Change SAP headers 
    df_excel.columns = headers 

    # Converting date columns
    for column in headers:
        if 'fecha' in column:
            df_excel[column] = pd.to_datetime(df_excel[column],format='%d/%m/%Y',errors='coerce').dt.date

    # Converting numerical columns
    numerical_columns = (
        'saldo_moneda_documento','saldo_moneda_sociedad'
    )

    for column in numerical_columns:
        df_excel[column] = df_excel[column].apply(convert_float)

    df_excel['cod_periodo'] = fbl1h_filename[0:6]
    df_excel['cod_periodo'] = df_excel['cod_periodo'].apply(convert_float)

    # Reorder my columns to be exactly the same as the schema of my table
    df_excel = df_excel[['cod_periodo'] + headers]

    print('Cargando SQL')
    # Loading Database to SQLite
    df_excel.to_sql(table_name, conn, if_exists='append', index=False)

# Call functions
session = sap_connection()
fbl1h_export(session)
file_transformation()

# Close connection SQL
conn.close()
