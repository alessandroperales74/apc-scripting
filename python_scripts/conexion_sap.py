import win32com.client as win32
import subprocess
from time import sleep

# Extraer los valores de user y password
# No es una buena práctica tener el usuario dentro del código, pero para este ejemplo lo dejo aquí.
# Lo puedes extraer como una variable de entorno o referenciando a un JSON
user = "admin"
password = "abc123"

# Conexión 
sap_s4 = '1.01 - SAP PRD - S/4'
 
def sap():
 
    def open_sap():
        # Ruta SAP
        sapgui_executable = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        try:
            subprocess.Popen([sapgui_executable])
            sleep(5)
        except Exception as e:
            print("Error OPEN_SAP:", str(e))
 
    def sap_connection():
 
        connection = None
 
        # Revisar si hay conexiones activas de SAP
        for i,conn in enumerate(application.Children):
 
            if conn.Description == sap_s4:
                print(f'Conexión {i + 1} correcta')
                connection = conn
                break
 
        if connection:
            # Si la conexión activa es igual a la conexión deseada, definir esa conexión
            print(f'Conexión SAP: {connection.Description}')
        else:
            # Si no está abierta la conexión deseada, abrir una nueva
            print('Conexión no encontrada. Abriendo conexión')
            connection = application.OpenConnection(sap_s4,True)
            session = connection.Children(0)
 
            # Llenar usuario y contraseña
            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
            session.findById("wnd[0]").sendVKey(0)
 
        return connection
 
    def sap_session():
 
        session = None
 
        # Total de sesiones activas
        sessions = connection.Children.Count
 
        # Si solo hay una sesión activa. Crear una nueva
        if sessions == 1:
            connection.Children(0).CreateSession()
 
        # Primero chequeamos si hay alguna ventana en el SESSION_MANAGER, lo que quiere decir que está libre
        for i, ses in enumerate(connection.Children):
 
            print(ses.Info.Transaction)
 
            if ses.Info.Transaction == 'SESSION_MANAGER':
                # Si la transacción activa es el SESSION MANAGER, usar esa
                print(f'Sesion {i + 1} libre')
                session = ses
                break
            elif ses.Info.Transaction == 'S000':
                # Estamos en la ventana de login, pero aún no hemos accedido a nuestra cuenta
                print(f'Sesion {i + 1} en ventana de login. Iniciando sesión...')
                # Llenar usuario y contraseña
                session = ses
                session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
                session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
                session.findById("wnd[0]").sendVKey(0)
                break

        # Si no hay ninguna en el SESSION_MANAGER, tomamos la primera que no esté ocupada
        if not session:
            for i, ses in enumerate(connection.Children):
 
                if ses.Busy == True:
                    # Si la sesión está ocupada. Es decir, en proceso de carga de otra transacción, pasar a la siguiente.
                    print(f'Sesion {i + 1} ocupada. Ir a la siguiente')
                else:
                    # Si la sesión no está ocupada, tomar esa.
                    print(f'Sesion {i + 1} activa. Utilizar')
                    session = ses
                    break
 
        return session
 
    print('Generando conexión SAP...')
    try:
        SapGuiAuto  = win32.GetObject("SAPGUI")
    except:
        print('Abriendo SAP...')
        open_sap()
        SapGuiAuto  = win32.GetObject("SAPGUI")
       
    application = SapGuiAuto.GetScriptingEngine
    connection = sap_connection()
    session = sap_session()
 
    return session

session = sap()
session.StartTransaction("FBL1N")

print('Finalizado!')