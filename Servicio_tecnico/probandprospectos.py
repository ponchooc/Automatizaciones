# ®Carlos Alfonso Ortega Molina®
# importacion de modulos requeridos para el proceso
import os
import math
import json
import time
import ftplib
import requests
import paramiko
from datetime import datetime, timedelta
from requests.auth import HTTPBasicAuth

# variables de configuracion manual de fechas opcionales
# si se dejan vacias el sistema tomara la fecha de ayer de forma automatica
FECHA_DESDE = "" # formato aaaa-mm-dd ej 2026-05-24
FECHA_HASTA = "" # formato aaaa-mm-dd ej 2026-06-03

# destinatarios del correo de notificacion editable por el usuario
CORREO_DESTINATARIOS = "ortegac@sanborns.com.mx,gresendiz@sears.com.mx,eacastillo@sears.com.mx,mdominguezc@sears.com.mx"

# definicion de la direccion base de la api de wms oracle cloud
BASE_URL = "https://e6.wms.ocs.oraclecloud.com/sears2/wms/lgfapi/v10/entity/order_hdr"

# definicion de parametros del segundo servidor ftp remotos y quemados en el codigo
FTP_HOST_SECUNDARIO = "140.240.11.6"
FTP_USER_SECUNDARIO = "ftpusr01"
FTP_PASS_SECUNDARIO = "ftpgral1"
FTP_REMOTE_DIR_SECUNDARIO = "/syp_data"
FTP_REMOTE_FILE_SECUNDARIO = "plomeria.txt"

# direccion ip del primer servidor remoto gnx
GNX_HOST = "140.240.11.1"

def cargar_credenciales():
    # variables globales para almacenar las credenciales recuperadas de archivos locales
    global WMS_USER, WMS_PASS, GNX_USER, GNX_PASS
    WMS_USER, WMS_PASS = "", ""
    GNX_USER, GNX_PASS = "", ""
    
    # lectura del archivo json de acceso para la api de wms oracle
    if os.path.exists(".acceso_wms"):
        try:
            with open(".acceso_wms", "r", encoding="utf-8") as f:
                datos = json.load(f)
                WMS_USER = datos.get("usuario", "")
                WMS_PASS = datos.get("contrasena", "")
        except:
            pass

    # lectura del archivo json de acceso para el servidor gnx
    if os.path.exists(".acceso_gnx"):
        try:
            with open(".acceso_gnx", "r", encoding="utf-8") as f:
                datos = json.load(f)
                GNX_USER = datos.get("usuario", "")
                GNX_PASS = datos.get("contrasena", "")
        except:
            pass

def exportar_prospectos_wms():
    # ®Carlos Alfonso Ortega Molina®
    # inicializacion de credenciales dinamicas
    cargar_credenciales()
    
    if not WMS_USER or not WMS_PASS or not GNX_USER or not GNX_PASS:
        print("[error] no se pudieron cargar las credenciales de los archivos de acceso")
        return

    # calculo de rango de fechas de forma dinamica o parametrizada
    if FECHA_DESDE and FECHA_HASTA:
        ayer_formato_api_inicio = FECHA_DESDE
        ayer_formato_api_fin = FECHA_HASTA
        try:
            f_temp = datetime.strptime(FECHA_HASTA, "%Y-%m-%d")
            ayer_formato_archivo = f_temp.strftime("%d%m%Y")
        except:
            ayer_formato_archivo = datetime.now().strftime("%d%m%Y")
    else:
        fecha_actual = datetime.now()
        fecha_ayer = fecha_actual - timedelta(days=1)
        ayer_formato_api_inicio = fecha_ayer.strftime("%Y-%m-%d")
        ayer_formato_api_fin = fecha_ayer.strftime("%Y-%m-%d")
        ayer_formato_archivo = fecha_ayer.strftime("%d%m%Y")
    
    nombre_archivo_local = f"reporte_ordenes_{ayer_formato_archivo}.txt"
    
    # configuracion de filtros para la consulta de la api wms
    params_api = {
        "create_ts__gte": f"{ayer_formato_api_inicio}T00:00:00",
        "create_ts__lte": f"{ayer_formato_api_fin}T23:59:59",
        "mod_ts__gte": f"{ayer_formato_api_inicio}T00:00:00",
        "mod_ts__lte": f"{ayer_formato_api_fin}T23:59:59",
        "status_id": "90",
        "order_type_id__in": "83,103",
        "fields": "cust_field_1,mod_ts"
    }

    page_size = 500
    params_api["page_size"] = page_size
    params_api["page"] = 1
    
    print(f"estableciendo conexion con oracle wms...")
    
    # utilizacion de la variable carlos como el objeto de sesion de consultas http
    carlos = requests.Session()
    carlos.auth = HTTPBasicAuth(WMS_USER, WMS_PASS)
    
    try:
        respuesta = carlos.get(BASE_URL, params=params_api)
        respuesta.raise_for_status()
        data = respuesta.json()
        
        total_resultados = data.get("result_count", 0)
        
        if total_resultados > 0:
            total_paginas = math.ceil(total_resultados / page_size)
            print(f"\n==================================================")
            print(f" total de registros encontrados : {total_resultados}")
            print(f" total de paginas a procesar     : {total_paginas}")
            print(f"==================================================\n")
        else:
            print("no se encontraron registros en el rango de fechas")
            carlos.close()
            total_resultados = 0
            primeros_registros = []
        
        if total_resultados > 0:
            primeros_registros = data.get("results", [])

    except Exception as e:
        print(f"error al inicializar la consulta de wms: {e}")
        carlos.close()
        return

    page_number = 1
    total_registros_escritos = 0
    
    print(f"creando archivo '{nombre_archivo_local}'...")
    
    with open(nombre_archivo_local, "w", encoding="utf-8") as f:
        while True:
            if not primeros_registros:
                break
                
            if page_number == 1:
                registros_pagina = primeros_registros
            else:
                params_api["page"] = page_number
                print(f"consultando pagina {page_number} de {total_paginas}...")
                try:
                    respuesta = carlos.get(BASE_URL, params=params_api)
                    respuesta.raise_for_status()
                    registros_pagina = respuesta.json().get("results", [])
                except Exception as e:
                    print(f"error al obtener la pagina {page_number}: {e}")
                    break

            if not registros_pagina:
                break
            
            # escribir linea por linea en el archivo plano
            for item in registros_pagina:
                sales_check = item.get("cust_field_1", "")
                raw_mod_ts = item.get("mod_ts", "")
                
                mod_ts_formatted = ""
                if raw_mod_ts and len(raw_mod_ts) >= 16:
                    anio = raw_mod_ts[0:4]
                    mes = raw_mod_ts[5:7]
                    dia = raw_mod_ts[8:10]
                    hora_min = raw_mod_ts[11:16]
                    
                    mod_ts_formatted = f"{dia}/{mes}/{anio} {hora_min}"
                
                linea = f"{sales_check}|{mod_ts_formatted}\n"
                f.write(linea)
            
            cantidad_actual = len(registros_pagina)
            total_registros_escritos += cantidad_actual
            print(f" -> pagina {page_number} procesada ({cantidad_actual} registros guardados)")
            
            if cantidad_actual < page_size:
                break
                
            page_number += 1

    carlos.close()
    print(f"\n[exito] archivo local '{nombre_archivo_local}' generado con {total_registros_escritos} registros")

    # fase de envio por ftp al primer servidor remoto gnx depositandolo como fase1.txt
    print(f"conectando a {GNX_HOST} via ftp para subir el reporte como fase1.txt...")
    try:
        with ftplib.FTP(GNX_HOST) as ftp_gnx:
            ftp_gnx.login(GNX_USER, GNX_PASS)
            ftp_gnx.voidcmd('TYPE I')
            ruta_remota_gnx = "/gnx_prod/manto/desa/trabajo/sears/carlos_ortega/prospectos"
            ftp_gnx.cwd(ruta_remota_gnx)
            
            with open(nombre_archivo_local, "rb") as f_subir:
                ftp_gnx.storbinary("STOR fase1.txt", f_subir)
            print(f"[exito] archivo subido a gnx en {ruta_remota_gnx} bajo el nombre fase1.txt")
    except Exception as e:
        print(f"error en la subida ftp al servidor gnx: {e}")
        return

    # fase de ejecucion de comandos por ssh en el servidor gnx
    print("estableciendo conexion ssh con paramiko para ejecutar dbaccess...")
    try:
        ssh_gnx = paramiko.SSHClient()
        ssh_gnx.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh_gnx.connect(GNX_HOST, username=GNX_USER, password=GNX_PASS)
        
        sh = ssh_gnx.invoke_shell(term='vt100', width=220, height=50)
        sh.settimeout(300)

        # funcion de apoyo para leer del buffer del shell interactivo
        def leer_buffer(max_seg=10):
            salida = ""
            inicio_tiempo = time.time()
            while True:
                if sh.recv_ready():
                    bloque = sh.recv(4096).decode('latin-1', 'ignore')
                    print(bloque, end="", flush=True)
                    salida += bloque
                    inicio_tiempo = time.time()
                elif time.time() - inicio_tiempo > max_seg:
                    break
                else:
                    time.sleep(0.3)
            return salida

        # funcion de apoyo para enviar comandos de terminal
        def ejecutar_comando_terminal(comando, descripcion="", espera=5):
            if descripcion:
                print(descripcion)
            if comando:
                time.sleep(1)
                sh.send(comando + "\n")
            leer_buffer(espera)

        # navegacion del menu interactivo inicial de gnx
        ejecutar_comando_terminal("1", "\n>>> seleccionando ambiente 1 sears...", 5)
        ejecutar_comando_terminal("1", ">>> seleccionando opcion 1 sears...", 5)
        
        # cambio de ruta al directorio de prospectos de carlos ortega
        ejecutar_comando_terminal("cd /gnx_prod/manto/desa/trabajo/sears/carlos_ortega/prospectos", ">>> ingresando a la ruta de prospectos...", 5)
        
        # ejecucion del comando dbaccess esperando su conclusion total
        print(">>> ejecutando dbaccess gen obtenedorplomeria.sql...")
        sh.send("dbaccess gen obtenedorplomeria.sql\n")
        leer_buffer(max_seg=300)
        
        print("\n>>> esperando liberacion del prompt...")
        leer_buffer(max_seg=30)
        
        # salida limpia de la terminal de comandos de ssh
        print(">>> cerrando el shell de comandos...")
        sh.send("exit\n")
        time.sleep(3)
        
        print(">>> finalizando sesion del menu de gnx...")
        sh.send("f\n")
        time.sleep(3)
        
        leer_buffer(max_seg=5)
        ssh_gnx.close()
        print("[exito] proceso de ejecucion remota sql completado con exito")
    except Exception as e:
        print(f"error critico en la ejecucion ssh: {e}")
        return

    # fase de descarga del archivo generado final999.txt desde gnx
    nombre_final_descargado = "final999.txt"
    print("conectando de nuevo via ftp a gnx para descargar final999.txt...")
    try:
        with ftplib.FTP(GNX_HOST) as ftp_descarga:
            ftp_descarga.login(GNX_USER, GNX_PASS)
            ftp_descarga.voidcmd('TYPE I')
            ftp_descarga.cwd("/gnx_prod/manto/desa/trabajo/sears/carlos_ortega/prospectos")
            
            with open(nombre_final_descargado, "wb") as f_bajar:
                ftp_descarga.retrbinary("RETR final999.txt", f_bajar.write)
            print(f"[exito] archivo {nombre_final_descargado} descargado a local")
    except Exception as e:
        print(f"error al descargar el archivo final999.txt de gnx: {e}")
        return

    # fase de subida al segundo servidor ftp remoto con credenciales harcodeadas renombrando como plomeria.txt
    print(f"conectando al segundo servidor {FTP_HOST_SECUNDARIO} via ftp...")
    exito_final = False
    try:
        with ftplib.FTP(FTP_HOST_SECUNDARIO) as ftp_secundario:
            ftp_secundario.login(FTP_USER_SECUNDARIO, FTP_PASS_SECUNDARIO)
            ftp_secundario.voidcmd('TYPE I')
            ftp_secundario.cwd(FTP_REMOTE_DIR_SECUNDARIO)
            
            with open(nombre_final_descargado, "rb") as f_subir_sec:
                ftp_secundario.storbinary(f"STOR {FTP_REMOTE_FILE_SECUNDARIO}", f_subir_sec)
            print(f"[exito] archivo subido al segundo servidor como {FTP_REMOTE_FILE_SECUNDARIO} en {FTP_REMOTE_DIR_SECUNDARIO}")
            exito_final = True
    except Exception as e:
        print(f"error en la subida ftp al segundo servidor: {e}")

    # limpieza de archivos temporales del primer servidor gnx
    print("eliminando archivos temporales fase1.txt y final999.txt del servidor gnx...")
    try:
        with ftplib.FTP(GNX_HOST) as ftp_limpieza:
            ftp_limpieza.login(GNX_USER, GNX_PASS)
            ftp_limpieza.cwd("/gnx_prod/manto/desa/trabajo/sears/carlos_ortega/prospectos")
            try:
                ftp_limpieza.delete("fase1.txt")
                print("[exito] archivo temporales fase1.txt eliminado de gnx")
            except Exception as e_del1:
                print(f"aviso no se pudo eliminar fase1.txt de gnx: {e_del1}")
            try:
                ftp_limpieza.delete("final999.txt")
                print("[exito] archivo temporales final999.txt eliminado de gnx")
            except Exception as e_del2:
                print(f"aviso no se pudo eliminar final999.txt de gnx: {e_del2}")
    except Exception as e_clean:
        print(f"aviso no se pudo establecer conexion ftp para limpieza de gnx: {e_clean}")

    # limpieza opcional de archivos locales generados
    try:
        if os.path.exists(nombre_archivo_local):
            os.remove(nombre_archivo_local)
        if os.path.exists(nombre_final_descargado):
            os.remove(nombre_final_descargado)
        print("limpieza de archivos temporales locales completada")
    except Exception as e:
        print(f"aviso error al limpiar archivos locales: {e}")

    # fase de envio de notificacion de correo usando el sistema smtp mailx de petota_gnx.py via ssh
    print("estableciendo conexion ssh para el envio de correo smtp de aviso...")
    try:
        ssh_correo = paramiko.SSHClient()
        ssh_correo.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh_correo.connect(GNX_HOST, username=GNX_USER, password=GNX_PASS)
        
        sh_c = ssh_correo.invoke_shell(term='vt100', width=220, height=50)
        sh_c.settimeout(300)
        
        # funcion de apoyo para leer del buffer del shell interactivo
        def leer_buffer_correo(max_seg=10):
            salida = ""
            inicio_tiempo = time.time()
            while True:
                if sh_c.recv_ready():
                    bloque = sh_c.recv(4096).decode('latin-1', 'ignore')
                    print(bloque, end="", flush=True)
                    salida += bloque
                    inicio_tiempo = time.time()
                elif time.time() - inicio_tiempo > max_seg:
                    break
                else:
                    time.sleep(0.3)
            return salida

        # funcion de apoyo para enviar comandos de terminal
        def ejecutar_comando_correo(comando, descripcion="", espera=5):
            if descripcion:
                print(descripcion)
            if comando:
                time.sleep(1)
                sh_c.send(comando + "\n")
            leer_buffer_correo(espera)

        # navegacion del menu interactivo inicial de gnx
        ejecutar_comando_correo("1", "\n>>> seleccionando ambiente 1 sears...", 5)
        ejecutar_comando_correo("1", ">>> seleccionando opcion 1 sears...", 5)
        
        # cambio de directorio remoto
        ejecutar_comando_correo("cd /gnx_prod/manto/desa/trabajo/sears/carlos_ortega/prospectos", ">>> ingresando a ruta de prospectos...", 5)
        
        # creacion del archivo de texto temporal para el cuerpo del correo de forma segura linea por linea
        # esto evita problemas con saltos de linea interactivos en la terminal ssh
        estatus_t = "EXITO" if exito_final else "ERROR"
        ejecutar_comando_correo("echo 'reporte de validacion de prospectos' > mail_body.txt", ">>> construyendo reporte txt...", 2)
        ejecutar_comando_correo("echo '' >> mail_body.txt", "", 1)
        ejecutar_comando_correo(f"echo 'fecha desde:              {ayer_formato_api_inicio}' >> mail_body.txt", "", 1)
        ejecutar_comando_correo(f"echo 'fecha hasta:              {ayer_formato_api_fin}' >> mail_body.txt", "", 1)
        ejecutar_comando_correo(f"echo 'lineas encontradas:       {total_registros_escritos}' >> mail_body.txt", "", 1)
        ejecutar_comando_correo(f"echo 'servidor de deposito:     {FTP_HOST_SECUNDARIO}' >> mail_body.txt", "", 1)
        ejecutar_comando_correo(f"echo 'ruta de deposito:         {FTP_REMOTE_DIR_SECUNDARIO}' >> mail_body.txt", "", 1)
        ejecutar_comando_correo(f"echo 'archivo depositado:       {FTP_REMOTE_FILE_SECUNDARIO}' >> mail_body.txt", "", 1)
        ejecutar_comando_correo(f"echo 'estatus de transferencia: {estatus_t}' >> mail_body.txt", "", 1)

        # ejecucion de la herramienta enviador_correos.sh pasando el contenido del archivo generado
        print(">>> enviando correo de notificacion via smtp...")
        comando_envio_correo = f'/gnx_prod/manto/desa/trabajo/sears/carlos_ortega/solo_petota/enviador_correos.sh "{CORREO_DESTINATARIOS}" "Prospectos" "$(cat mail_body.txt)"'
        sh_c.send(comando_envio_correo + "\n")
        leer_buffer_correo(max_seg=30)
        
        # eliminacion del archivo temporal de texto del correo en gnx
        ejecutar_comando_correo("rm -f mail_body.txt", ">>> eliminando archivo temporal de correo...", 2)
        
        # salida de la terminal de comandos de ssh
        ejecutar_comando_correo("exit", ">>> cerrando el shell de comandos de correo...", 3)
        ejecutar_comando_correo("f", ">>> finalizando sesion del menu de gnx...", 3)
        
        ssh_correo.close()
        print("[exito] proceso de envio de correo completado")
    except Exception as e:
        print(f"error al enviar el correo smtp via ssh: {e}")

if __name__ == "__main__":
    exportar_prospectos_wms()
# ®Carlos Alfonso Ortega Molina®
