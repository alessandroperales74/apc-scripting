# How to Export an Excel File from SAP with Python and Load It into a SQLite Local Database

When I started doing data analysis with SAP transactions, I needed to organize and structure my data over time. Unfortunately, I didn’t have a server available to load the data, and SharePoint's cloud options were quite limited for my needs.

It was at that moment that I discovered SQLite and decided to explore it. While it's not ideal to rely on a local database in the long term, it's a good way to get started in the field.

If you don’t have the necessary infrastructure due to circumstances beyond your control, you can still do interesting things with SQLite. In my case, I was able to work and create reports in Power BI using ODBC connectors (at least until I had a server to load my data) or manage a number of rows that would not have been possible using Excel.

Even so, it was much easier to migrate data that was already structured and organized beforehand. So, if you don’t have a server to handle this kind of data, it’s not a bad idea to use SQLite (which is free and very easy to use) to start working on your projects.

In this case, I’m going to show how I did a very simple - yet effective - ETL to extract data from SAP and load it into SQLite.

## 1. Importing Required Libraries

```python
import win32com.client # For being able to make the connection with SAP and Excel
import os # Not mandatory, but it gives you some flexibility with folder paths
import pandas as pd # Classic Pandas stuff
import sqlite3
```

## 2. SAP Exporting

And now, I need to do the SAP connection. I'm not going into details about these particular code because I want to explain it later.

```python
def sap_connection():

    print('Connecting with SAP...')
    SapGuiAuto  = win32com.client.GetObject("SAPGUI") # Making SAP connection with Python
    application = SapGuiAuto.GetScriptingEngine 
    connection = application.Children(0) # First connection available
    session    = connection.Children(0) # First window available

    return session
```

Then, I'll declare some necessary variables for our later functions. I need the inputs required for the transaction. In this case, I'll try with FBL1H. You can do it with whichever you want. Just make sure you have ALL THE INPUTS.

```python

fbl1h_variant = 'APC_DB_SQLITE' # It has the required configuration of the non variable inputs
fbl1h_layout = 'APC_DB_SQLITE' # This is the column distribution of the transactions. It's a good practice to have one specifically for this load

fbl1h_filepath = 'D:\\sqlite_db\\inputs'
fbl1h_filename = '202407_FBL1H.xlsx'

first_date = '01.07.2024'
last_date = '31.07.2024'

def fbl1h_contabilizadas(session):
    print('Exporting FBL1H...')
    session.StartTransaction("FBL1H")
    session.findById("wnd[0]/tbar[1]/btn[17]").press() 
    session.findById("wnd[1]/usr/txtV-LOW").text = fbl1h_variant
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]").sendVKey(8)

    # You can make these dates dynamic. You can use an input or a function to get the first and last date of the previous month
    session.findById("wnd[0]/usr/ctxtS_PDATE-LOW").text = first_date
    session.findById("wnd[0]/usr/ctxtS_PDATE-HIGH").text = last_date

    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/chkP_TY_SPG").selected = True
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").text = fbl1h_layout
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/shellcont/shell").pressToolbarButton("SHOWBUT")
    session.findById("wnd[0]/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")

    # In this particular case,
    # I'm fine with Excel files due to the size of the report. It's not too large, so it's manageable
    session.findById("wnd[0]/shellcont/shell").selectContextMenuItem("&XXL")

    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").text = fbl1h_filename
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/cmbGS_EXPORT-FORMAT").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[20]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = fbl1h_filepath
    session.findById("wnd[1]").sendVKey(11) 
```
ACHTUNG❗❗❗

In this particular case, I'm fine with Excel files due to the size of the report BUT I need to warn you:

- **If you want to export in an Excel File:** It often takes a bit more time and Excel tends to opens automatically after exporting, so -sometimes- you have to do an additional step to close it before reading it with Pandas -in this case, I haven't had any problems but I did have in another scripts-. Also, consider the wait time. For example, if your report has 500k rows, you can technically export it (since the row limit in Excel is around 1 million rows), but SAP may crash due to the wait time.

- **If you want to export in an txt/csv File:**: Is faster and it doesn't open automatically but it's a bit tricky if you find some line breaks or you have a text column with '|' character since that character is default SAP delimiter. In a 'Text' column, I often find this '||' 😒 ... With that being said: Be careful.

## 3. Transforming the Excel File and Loading to Database

Now I have the Excel file exported, so I have to load it to my Database.

For this case, I have a function to transform my FBL1H report. You can do any transformation you want.

```python

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
        # Just in case. If you read a SAP txt or csv, it’s mandatory to strip the columns
        for columna in df.columns:
            df[columna] = df[columna].str.strip()

    # Reading my Excel
    # Here I use the os function join to have my full Excel path
    df_excel = pd.read_excel(os.path.join(fbl1h_filepath,fbl1h_filename),dtype=str)

    # From here on, this is tailored to my needs. You can replace this with any transformation you require. I’m leaving it here for reference.

    strip_columns(df_excel)

    # These headers need to match the table schema in SQLite.
    # I’ll add a period column later (which isn’t in SAP), but aside from that, this is how I created my table.

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
            df_excel[column] = pd.to_datetime(df_excel[column],format='%d.%m.%Y',errors='coerce').dt.date

    # Converting numerical columns
    numerical_columns = (
        'saldo_moneda_documento','saldo_moneda_sociedad'
    )

    for column in numerical_columns:
        df_excel[column] = df_excel[column].apply(convert_float)

    df_excel['cod_periodo'] = fbl1h_filename[0:6]
    df_excel['cod_periodo'] = df_excel['cod_periodo'].apply(convert_float)

    # Reorder columns to match the schema of the table
    df_excel = df_excel[['cod_periodo'] + headers]

    # Loading Database to SQLite
    df_excel.to_sql(table_name, conn, if_exists='append', index=False)

# Call functions
session = sap_connection()
fbl1h_export(session)
file_transformation()

# Close SQLite connection
conn.close()
```

Finally, you have to be aware that SQLite doesn't have all the features that SQL Server, PostgreSQL, etc., have, but it's an easy way to start. When you have more knowledge, you can migrate to a more robust database engine.

You can check the python file in this link: ![SQLite_Python](python_scripts/sqlite_python.py)

Hope you find it useful! 🙋‍♂️🙋‍♂️🙋‍♂️🙋‍♂️
