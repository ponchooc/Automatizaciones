# ®Carlos Alfonso Ortega Molina®
# importacion de modulos del sistema y bibliotecas requeridas para el proceso
import requests
import pandas as pd
import os
import json
import threading
import queue
import time
import ftplib
import paramiko
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import sys
import win32com.client
import pythoncom

# formatear la fecha actual para nombrar archivos de salida de forma unica
fecha_hoy = datetime.now().strftime("%d%m%Y")

# definicion de las direcciones web base para consultar los servicios de wms y otm
URL_BASE = "https://e6.wms.ocs.oraclecloud.com/sears2/wms/lgfapi/v10/entity/order_hdr/"
USUARIO = ""
CONTRASENA = ""
CORREO_WMS = ""

OTM_URL_BASE = "https://otmgtm-gsanborns202311wms.otmgtm.us-ashburn-1.ocs.oraclecloud.com/logisticsRestApi/resources-int/v2/orderReleases"
OTM_USUARIO = ""
OTM_CONTRASENA = ""

# direccion ip y credenciales por defecto para el servidor gnx
GNX_HOST = "140.240.11.1"
GNX_USER = ""
GNX_PASS = ""

# definicion de la clase encargada de duplicar la salida estandar tanto a la consola como a la bitacora de texto
class duplicador_salida:
    # constructor que inicializa la consola y abre el archivo de bitacora
    def __init__(self, nombre_archivo):
        self.consola = sys.stdout
        self.archivo_log = open(nombre_archivo, "a", encoding="utf-8")
    # escribe el mensaje recibido en ambos flujos de salida
    def write(self, mensaje):
        self.consola.write(mensaje)
        self.archivo_log.write(mensaje)
        # forzar escritura inmediata en disco para log en tiempo real
        self.flush()
    # asegura que todo el buffer se escriba directamente en disco
    def flush(self):
        self.consola.flush()
        self.archivo_log.flush()
    # cierra el archivo log de salida al finalizar el proceso
    def close(self):
        self.archivo_log.close()

# definicion de las rutas locales para los archivos generados durante el procesamiento
archivo_log = f"log_{fecha_hoy}.txt"
archivo_entrada = "cliyclitie.txt"
archivo_salida = f"reporte_ordenes_sears_{fecha_hoy}.xlsx"
archivo_control = f"control_i68_{fecha_hoy}.json"
archivo_texto_procesado = f"cliyclitie_procesado_{fecha_hoy}.txt"
archivo_errores = f"errores_{fecha_hoy}.txt"

# objetos de sincronizacion para procesos multihilo y manejo de colas de resultados
bloqueo_fichero = threading.Lock()
cola_resultados = queue.Queue()
# bloqueo ferreo de seguridad para evitar colisiones de ordenes entre hilos
bloqueo_ordenes_activas = threading.Lock()
ordenes_activas = set()

# funcion para recuperar y cargar las credenciales de wms otm y gnx de los archivos de configuracion
def obtener_credenciales():
    global USUARIO, CONTRASENA, CORREO_WMS, GNX_USER, GNX_PASS, OTM_USUARIO, OTM_CONTRASENA
    wms_ok = False
    otm_ok = False
    # lectura de las credenciales para la plataforma otm si existe el archivo
    if os.path.exists(".acceso_otm"):
        try:
            with open(".acceso_otm", "r") as f:
                d = json.load(f)
                OTM_USUARIO = d.get("usuario")
                OTM_CONTRASENA = d.get("contrasena")
                if OTM_USUARIO and OTM_CONTRASENA: otm_ok = True
        except: pass
    # lectura de las credenciales para la plataforma wms si existe el archivo
    if os.path.exists(".acceso_wms"):
        try:
            with open(".acceso_wms", "r") as f:
                d = json.load(f)
                USUARIO = d.get("usuario")
                CONTRASENA = d.get("contrasena")
                CORREO_WMS = d.get("correo")
                if USUARIO and CONTRASENA: wms_ok = True
        except: pass

    gnx_ok = False
    # lectura de las credenciales para el servidor gnx si existe el archivo
    if os.path.exists(".acceso_gnx"):
        try:
            with open(".acceso_gnx", "r") as f:
                d = json.load(f)
                GNX_USER = d.get("usuario")
                GNX_PASS = d.get("contrasena")
                if GNX_USER and GNX_PASS: gnx_ok = True
        except: pass
    return wms_ok and gnx_ok and otm_ok
# funcion para descargar el archivo de entrada cliyclitie txt desde el servidor ftp de gnx
def fase_cero_ftp():
    # imprimir encabezado informativo de la fase cero del flujo
    print(f"\n ══════════════════════════════════════════════════════════════════════")
    print(f"  [FASE 0] AUTO-ABASTECIMIENTO DE MATERIA PRIMA (FTP)")
    print(f" ══════════════════════════════════════════════════════════════════════")
    print(f" [≫] Conectando a {GNX_HOST}...")
    try:
        # iniciar conexion ftp usando la ip del servidor gnx y codificacion latin1
        with ftplib.FTP(GNX_HOST, encoding='latin-1') as ftp:
            # realizar el inicio de sesion con las credenciales cargadas de gnx
            ftp.login(GNX_USER, GNX_PASS)
            # definir la ruta del directorio remoto de oracle
            ruta = "/gnx_prod/manto/desa/trabajo/sears/ORACLE"
            # cambiar el directorio de trabajo activo del ftp a la ruta especificada
            ftp.cwd(ruta)
            print(f" [i] Ruta FTP activa: {ruta}")
            try:
                # comprobar si el archivo de entrada se encuentra en la lista remota
                if archivo_entrada in ftp.nlst():
                    # obtener la fecha de modificacion del archivo remoto
                    res = ftp.voidcmd(f"MDTM {archivo_entrada}")
                    # recortar y procesar la fecha de respuesta del comando mdtm
                    fecha_srv = res.split()[1][:8]
                    print(f" [✓] Archivo remoto valido: {archivo_entrada} ({fecha_srv})")
                    # verificar si la fecha del archivo remoto coincide con el dia de hoy
                    if fecha_srv == datetime.now().strftime("%Y%m%d"):
                        print(" [≫]  Descargando...")
                        # abrir el archivo local cliyclitie txt para escritura binaria
                        with open(archivo_entrada, "wb") as f_loc:
                            # transferir y escribir el archivo del servidor en el archivo local
                            ftp.retrbinary(f"RETR {archivo_entrada}", f_loc.write)
                        print(f" [✓] Descarga completada: {os.path.getsize(archivo_entrada)} bytes.")
                        # retornar verdadero al finalizar exitosamente la descarga
                        return True
                    else:
                        print(" [!] ALERTA: Archivo remoto no es de hoy.")
            except: pass
    except: pass
    # retornar falso si ocurrio cualquier excepcion o no se pudo descargar
    return False

# funcion para subir las incidencias pendientes a gnx y disparar el script sh via ssh
def subir_ftp_y_ejecutar_ssh(errores_pendientes):
    # imprimir el encabezado informativo del protocolo ssh y ftp
    print(f" ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print(f"  [GNX] INICIANDO PROTOCOLO REMOTO PARA {len(errores_pendientes)} INCIDENCIAS")
    print(f" ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    
    # definir el nombre del archivo unl de carga temporal para gnx
    archivo_carga = "cargai15.unl"
    try:
        # abrir el archivo unl con codificacion latin1 y salto de linea de tipo unix
        with open(archivo_carga, "w", newline='\n', encoding='latin-1') as f_unl:
            # conjunto para almacenar los registros de sales ya procesados y evitar duplicidad
            vistos_carlos = set()
            # iterar por cada uno de los errores identificados en el lote
            for error in errores_pendientes:
                # obtener el identificador de sales check libre de espacios vacios
                sales = error.get('sales_check', '').strip()
                # verificar que la cadena sea valida y no se haya procesado previamente
                if sales and sales not in vistos_carlos:
                    # escribir el sales check seguido de un salto de linea unix
                    f_unl.write(f"{sales}\n")
                    # agregar el registro al conjunto de elementos unicos procesados
                    vistos_carlos.add(sales)
        print(f" [✓] Archivo de carga generado: {archivo_carga} (Unix/Latin-1)")
    except Exception as e:
        print(f"error generando archivo de carga: {e}")
        # retornar falso si falla la generacion local del archivo de carga unl
        return False
        
    print(f" [≫] Subiendo {archivo_carga} a {GNX_HOST} via FTP...")
    try:
        # establecer conexion ftp con el servidor gnx
        with ftplib.FTP(GNX_HOST) as ftp:
            # iniciar sesion en el ftp con el usuario y clave cargados
            ftp.login(GNX_USER, GNX_PASS)
            # forzar la transferencia en modo de tipo binario para no alterar saltos de linea
            ftp.voidcmd('TYPE I')
            # cambiar el directorio de trabajo del servidor ftp a la carpeta de carlos ortega
            ftp.cwd("/gnx_prod/manto/desa/trabajo/sears/carlos_ortega")
            # abrir el archivo de carga local en modo de lectura binaria
            with open(archivo_carga, "rb") as f_sub: 
                # subir el archivo al directorio remoto usando el comando stor del ftp
                ftp.storbinary(f"STOR {archivo_carga}", f_sub)
            print(f" [✓] Transferencia ferrea de {archivo_carga} exitosa.")
        
        # renombrado local inmediato para guardar historico diario del archivo de carga unl
        try:
            # generar el nombre del archivo de bitacora historica para cargai15
            archivo_i15_hist = f"cargai15_{fecha_hoy}.unl"
            # eliminar el archivo historico si ya existia previamente
            if os.path.exists(archivo_i15_hist): os.remove(archivo_i15_hist)
            # renombrar el archivo cargai15 unl al nombre de respaldo historico
            os.rename(archivo_carga, archivo_i15_hist)
            print(f" [i] Archivo local {archivo_carga} archivado como {archivo_i15_hist}.")
        except Exception as e_ren:
            print(f" [!] Aviso: No se pudo archivar localmente: {e_ren}")
            
    except Exception as e:
        print(f" [X] Error en subida FTP: {e}")
        # retornar falso si falla la subida por ftp al servidor
        return False
    
    print(" [≫] Estableciendo conexion SSH (Paramiko) para reproceso...")
    try:
        # inicializar el cliente ssh del modulo paramiko
        ssh = paramiko.SSHClient()
        # cargar la politica de adicion automatica de llaves de host desconocidas
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        # conectar al servidor gnx usando el host puerto usuario y contraseña
        ssh.connect(GNX_HOST, username=GNX_USER, password=GNX_PASS)
        print(" [✓] Conexion establecida. Iniciando eco del servidor.")
        # invocar shell interactiva vt100 para enviar comandos de terminal
        sh = ssh.invoke_shell(term='vt100', width=220, height=50)
        # definir tiempo de espera maximo para operaciones del canal de lectura
        sh.settimeout(300)
        
        # funcion interna para leer los datos que llegan del shell hasta encontrar inactividad
        def leer_hasta_prompt(max_seg=10):
            salida = ""
            inicio = time.time()
            while True:
                # comprobar si hay bytes disponibles en el canal del shell
                if sh.recv_ready():
                    # decodificar el bloque de bytes usando codificacion latin1 ignorando fallos
                    bloque = sh.recv(4096).decode('latin-1', 'ignore')
                    # imprimir el bloque leido de forma directa en pantalla sin saltos extras
                    print(bloque, end="", flush=True)
                    # concatenar el bloque decodificado al resultado total de la lectura
                    salida += bloque
                    # restablecer el tiempo inicial para control de inactividad
                    inicio = time.time()
                # romper el ciclo si transcurre mas del tiempo maximo de inactividad
                elif time.time() - inicio > max_seg:
                    break
                else:
                    # pausa corta para dar tiempo al canal de recibir mas informacion
                    time.sleep(0.3)
            return salida

        # funcion interna para enviar un comando al shell interactivo y esperar su ejecucion
        def run_cmd(cmd, desc="", espera_seg=5):
            if desc: print(desc)
            if cmd: 
                # pausa prudencial antes del envio para el buffer del shell
                time.sleep(1) 
                # enviar el comando de texto con salto de linea
                sh.send(cmd + "\n")
            # drenar y leer la salida del shell interactivo tras enviar el comando
            leer_hasta_prompt(espera_seg)
            
        # ejecutar comando para seleccionar la opcion uno del ambiente sears
        run_cmd("1", "\n>>> [AUTO] Seleccionando Ambiente 1 (SEARS)...", 5)
        # ejecutar comando para seleccionar la opcion uno del menu principal
        run_cmd("1", ">>> [AUTO] Seleccionando Opcion 1 (SEARS)...", 5)
        # cambiar de directorio remoto hacia la carpeta absoluta de carlos ortega
        run_cmd("cd /gnx_prod/manto/desa/trabajo/sears/carlos_ortega", ">>> [AUTO] Seleccionando ruta de trabajo absoluta...", 5)

        print(">>> [AUTO] Ejecutando script de reproceso i68... (esperando conclusion total)")
        # enviar el comando de ejecucion del script reprocesai68 sh
        sh.send("./reprocesai68.sh\n")
        # esperar a que termine el script dando un tiempo de inactividad largo de diez minutos
        leer_hasta_prompt(max_seg=600)

        print("\n>>> [AUTO] Esperando a que el prompt se libere ('>')...")
        # esperar hasta treinta segundos para asegurar recibir el caracter del prompt del shell
        leer_hasta_prompt(max_seg=30)

        print(">>> [AUTO] Saliendo del shell ('exit')...")
        # enviar comando de salida del shell interactivo
        sh.send("exit\n")
        # pausa corta de tres segundos para que el shell termine
        time.sleep(3)

        print(">>> [AUTO] Enviando 'f' para finalizar sesion en el menu principal...")
        # enviar la tecla f para cerrar el menu heredado de gnx
        sh.send("f\n")
        # pausa de tres segundos para cerrar la sesion principal del menu
        time.sleep(3)

        # drenar cualquier salida pendiente del canal
        leer_hasta_prompt(max_seg=5)
        # cerrar la sesion ssh de forma segura
        ssh.close()
        # retornar verdadero al finalizar exitosamente la ejecucion de la secuencia ssh
        return True
    except Exception as e:
        print(f"error critico ssh: {e}")
        # retornar falso si ocurrio algun error en la secuencia de comandos ssh
        return False

def cargar_historial_ayer():
    print(" [≫] Buscando historial de validacion...")
    ayer = datetime.now() - timedelta(days=1)
    str_yesterday = ayer.strftime("%d%m%Y")
    archivo_ayer = f"control_i68_{str_yesterday}.json"
    if not os.path.exists(archivo_ayer):
        return {}
    
    print(f"cargando memoria historica desde: {archivo_ayer}")
    cache_historica = {}
    try:
        with open(archivo_ayer, "r") as f:
            for carlos in f:
                try:
                    carlos = carlos.strip()
                    if not carlos or carlos in ["[", "]", ","]: continue
                    obj = json.loads(carlos)
                    if obj.get('etiqueta_final') == "OK VERIFICADO API" or "YA VALIDADA ANTES" in obj.get('etiqueta_final', ''):
                        carlos = str(obj.get('orden_buscada'))
                        fecha_val = obj.get('fecha_validacion_original', obj.get('fecha_hora_proceso', 'desconocida'))
                        if len(carlos) == 11: carlos = "0" + carlos
                        cache_historica[carlos] = {'data': obj, 'fecha_validacion': fecha_val}
                except: continue
        print(f"memoria cargada: {len(cache_historica)} ordenes ya validadas previamente.")
    except Exception as e:
        print(f"aviso: error al leer historial {archivo_ayer}: {e}")
    return cache_historica

def cargar_estado():
    mapa_recuperado = {}
    if os.path.exists(archivo_control):
        try:
            with open(archivo_control, "r") as f:
                for num_linea, carlos in enumerate(f, 1):
                    contenido = carlos.strip()
                    if not contenido or contenido in ["[", "]", ","]: continue
                    try:
                        obj = json.loads(contenido)
                        if 'orden_buscada' in obj:
                            o_key = str(obj['orden_buscada'])
                            if len(o_key) == 11: o_key = "0" + o_key
                            mapa_recuperado[o_key] = obj
                    except: continue 
        except: pass
    return list(mapa_recuperado.values())

def hilo_guardado_continuo():
    while True:
        resultado = cola_resultados.get()
        if resultado is None: 
            cola_resultados.task_done()
            break
        try:
            with bloqueo_fichero:
                with open(archivo_control, "a") as f_json:
                    f_json.write(json.dumps(resultado) + "\n")
                
                with open(archivo_texto_procesado, "a") as f_txt:
                    carlos = resultado.get('orden_buscada', 'n/a')
                    sales = resultado.get('sales_check', 'n/a')
                    estado = resultado.get('etiqueta_final', 'NO MAL ERROR')
                    linea_txt = f"{carlos}|{sales}|{estado}"
                    if "YA VALIDADA ANTES" in estado:
                        fecha_orig = resultado.get('fecha_validacion_original', 'N/A')
                        if " " in fecha_orig: fecha_orig = fecha_orig.split(" ")[0]
                        linea_txt += f"|VALIDADO EL DIA {fecha_orig}"
                    f_txt.write(linea_txt + "\n")
        except: pass
        cola_resultados.task_done()

def consultar_lotes_ordenes(lista_datos_entrada, sesion=None, sesion_otm=None):
    # ®Carlos Alfonso Ortega Molina®
    if not lista_datos_entrada: return []
    
    # control ferreo de exclusion mutua para evitar verificacion simultanea de las mismas ordenes
    with bloqueo_ordenes_activas:
        lista_datos_entrada = [carlos for carlos in lista_datos_entrada if carlos[0] not in ordenes_activas]
        for carlos in lista_datos_entrada:
            ordenes_activas.add(carlos[0])
            
    if not lista_datos_entrada: return []
    
    mapa_lote = {carlos: sales for carlos, sales in lista_datos_entrada}
    ordenes_str = ",".join(mapa_lote.keys())
    resultados_finales_lote = []
    
    carlos_ses_creada = False
    carlos_otm_ses_creada = False
    
    try:
        if sesion is None:
            sesion = requests.Session()
            estrategia_reintento = Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
            adaptador = HTTPAdapter(pool_connections=1, pool_maxsize=1, max_retries=estrategia_reintento)
            sesion.mount("https://", adaptador)
            sesion.auth = (USUARIO, CONTRASENA)
            carlos_ses_creada = True
            
        if sesion_otm is None:
            sesion_otm = requests.Session()
            estrategia_reintento = Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
            adaptador = HTTPAdapter(pool_connections=1, pool_maxsize=1, max_retries=estrategia_reintento)
            sesion_otm.mount("https://", adaptador)
            sesion_otm.auth = (OTM_USUARIO, OTM_CONTRASENA)
            carlos_otm_ses_creada = True
            
        respuesta = sesion.get(URL_BASE, params={'order_nbr__in': ordenes_str, 'page_size': 40}, timeout=25)
        ahora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        encontrados = {}
        if respuesta.status_code == 200:
            datos = respuesta.json()
            for info in datos.get('results', []):
                o_api = info.get('order_nbr')
                if o_api: encontrados[o_api] = info
                
        # validacion api otm
        encontrados_otm = set()
        otm_error = False
        if sesion_otm:
            query_parts = [f'orderReleaseXid eq "{carlos}"' for carlos in mapa_lote.keys()]
            q_param = " or ".join(query_parts)
            otm_params = {'q': q_param, 'fields': 'orderReleaseXid', 'limit': 40}
            try:
                resp_otm = sesion_otm.get(OTM_URL_BASE, params=otm_params, timeout=25)
                if resp_otm.status_code == 200:
                    datos_otm = resp_otm.json()
                    for item in datos_otm.get('items', []):
                        o_api = item.get('orderReleaseXid')
                        if o_api: encontrados_otm.add(str(o_api))
                else:
                    otm_error = True
            except:
                otm_error = True
        else:
            # si no hay sesion otm asumimos error para forzar
            otm_error = True

        for carlos, sales in lista_datos_entrada:
            res_obj = {'orden_buscada': carlos, 'sales_check': sales, 'fecha_hora_proceso': ahora}
            
            existe_wms = carlos in encontrados
            existe_otm = carlos in encontrados_otm
            if otm_error: existe_otm = False
            
            if existe_wms and existe_otm:
                info = encontrados[carlos]
                res_obj.update({
                    'wms_id_interno': info.get('id'),
                    'wms_order_nbr_api': info.get('order_nbr'),
                    'wms_estatus_id': info.get('status_id'),
                    'wms_bodega': info.get('facility_id', {}).get('key') if info.get('facility_id') else "n/a",
                    'otm_order_release_xid': carlos,
                    'resultado': "exito en wms y otm",
                    'etiqueta_final': "OK VERIFICADO API",
                    'fecha_validacion_original': ahora 
                })
            else:
                if not existe_wms and not existe_otm:
                    motivo = "no existe en wms ni otm"
                elif existe_wms and not existe_otm:
                    # wms si respondio pero otm no tiene la orden - guardar datos wms igual
                    motivo = "existe en wms pero no otm"
                    info = encontrados[carlos]
                    res_obj.update({
                        'wms_id_interno': info.get('id'),
                        'wms_order_nbr_api': info.get('order_nbr'),
                        'wms_estatus_id': info.get('status_id'),
                        'wms_bodega': info.get('facility_id', {}).get('key') if info.get('facility_id') else "n/a",
                        'otm_order_release_xid': "NO ENCONTRADO EN OTM",
                    })
                else:
                    # otm si respondio pero wms no tiene la orden
                    motivo = "existe en otm pero no wms"
                    res_obj.update({
                        'wms_id_interno': "NO ENCONTRADO EN WMS",
                        'wms_estatus_id': "NO ENCONTRADO EN WMS",
                        'wms_bodega': "NO ENCONTRADO EN WMS",
                        'otm_order_release_xid': carlos,
                    })

                if respuesta.status_code != 200:
                    motivo = "error api wms"
                if otm_error:
                    motivo += " | error api otm"

                res_obj.update({'resultado': motivo, 'etiqueta_final': "NO MAL ERROR"})
            resultados_finales_lote.append(res_obj)
        return resultados_finales_lote
    except Exception as e:
        return [{'orden_buscada': c, 'sales_check': s, 'resultado': f"fallo lote: {str(e)}", 'etiqueta_final': "NO MAL ERROR", 'fecha_hora_proceso': datetime.now().strftime("%d/%m/%Y %H:%M:%S")} for c, s in lista_datos_entrada]
    finally:
        if carlos_ses_creada:
            try: sesion.close()
            except: pass
        if carlos_otm_ses_creada:
            try: sesion_otm.close()
            except: pass
        with bloqueo_ordenes_activas:
            for carlos in lista_datos_entrada:
                if carlos[0] in ordenes_activas:
                    ordenes_activas.remove(carlos[0])

def verificar_reentrada(resultados_previos, sesion, sesion_otm):
    if not resultados_previos: return resultados_previos
    candidatos = [r for r in resultados_previos if "YA VALIDADA ANTES" not in r.get('etiqueta_final', '')]
    if not candidatos: return resultados_previos
    num_a_verificar = min(3, len(candidatos))
    ultimas_3 = candidatos[-num_a_verificar:]
    indices_map = {r['orden_buscada']: i for i, r in enumerate(resultados_previos) if r in ultimas_3}
    for orden_res in ultimas_3:
        idx = indices_map.get(orden_res['orden_buscada'])
        if idx is not None:
             datos_entrada = [(orden_res['orden_buscada'], orden_res['sales_check'])]
             re_chequeo = consultar_lotes_ordenes(datos_entrada, sesion, sesion_otm)[0]
             if re_chequeo['etiqueta_final'] != orden_res.get('etiqueta_final'):
                 resultados_previos[idx] = re_chequeo
    return resultados_previos

def esperar_y_revalidar(sesion, sesion_otm, fase_actual, errores_pendientes, minutos):
    archivo_checkpoint = f".fase{fase_actual}_iniciada_i68_{fecha_hoy}"
    segundos = minutos * 60
    ahora = datetime.now()
    luego = ahora + timedelta(seconds=segundos)
    
    if fase_actual == 3:
        titulo = "VALIDACION POST-REPROCESO i68"
        label_X = "Ordenes a monitorear"
    else:
        titulo = "REVISION FINAL DEFINITIVA i68"
        label_X = "Ordenes en Revision Final"

    print(f"\n════════════════════════════════════════════════════════════")
    print(f"  [FASE {fase_actual}] {titulo} ({minutos} min)")
    print(f"════════════════════════════════════════════════════════════")
    print(f" | {label_X}:    {len(errores_pendientes)}")
    
    # inteligencia verificar si esta espera ya se cumplio en una ejecucion previa
    if os.path.exists(archivo_checkpoint):
        print(f" | [!] AVISO: El tiempo de espera de la Fase {fase_actual} ya se cumplio anteriormente.")
        print(f" | [!] Saltando pausa de {minutos} min para proceder directo a la validacion API.")
    else:
        print(f" | Pausa de seguridad:      {segundos} seg ({minutos} min)")
        print(f" | Inicio: {ahora.strftime('%H:%M:%S')} | Re-consulta: {luego.strftime('%H:%M:%S')}")
        print(f"------------------------------------------------------------\n")
        # crear checkpoint antes de dormir
        try: open(archivo_checkpoint, "w").close()
        except: pass
        time.sleep(segundos)
        
    print(f"------------------------------------------------------------")
    print(f"[API WMS] Consultando WMS Oracle Cloud - Fase {fase_actual}...")
    print(f"[API OTM] Consultando OTM Oracle Transportation - Fase {fase_actual}...\n")

    lote_trabajo = []
    for err in errores_pendientes:
        orden_orig = err['orden_buscada']
        # mantener regla de los 12 digitos si aplica
        if fase_actual >= 2 and len(orden_orig) == 11:
            lote_trabajo.append(("0" + orden_orig, err['sales_check']))
        else:
            lote_trabajo.append((orden_orig, err['sales_check']))

    tamanio_lote = 40
    total_lotes_fase = (len(lote_trabajo) + tamanio_lote - 1) // tamanio_lote
    for i in range(0, len(lote_trabajo), tamanio_lote):
        num_lote = (i // tamanio_lote) + 1
        segmento = lote_trabajo[i:i + tamanio_lote]
        print(f" [F{fase_actual}] lote wms/otm [{num_lote}/{total_lotes_fase}] - consultando {len(segmento)} ordenes...")
        resultados = consultar_lotes_ordenes(segmento, sesion, sesion_otm)
        ok_lote   = sum(1 for r in resultados if r.get('etiqueta_final') == 'OK VERIFICADO API')
        mal_lote  = len(resultados) - ok_lote
        print(f"          wms+otm ok: {ok_lote} | pendientes: {mal_lote}")
        with bloqueo_fichero:
            with open(archivo_control, "a") as f_json:
                for res in resultados:
                    f_json.write(json.dumps(res) + "\n")

    datos_finales = cargar_estado()
    remanentes_post = [item for item in datos_finales if item.get('etiqueta_final') == "NO MAL ERROR"]
    exitos_post = len(errores_pendientes) - len(remanentes_post)
    if exitos_post < 0: exitos_post = 0

    # desglose por sistema de los que siguen fallando
    solo_otm  = sum(1 for r in remanentes_post if "existe en wms pero no otm" in r.get('resultado', ''))
    solo_wms  = sum(1 for r in remanentes_post if "existe en otm pero no wms" in r.get('resultado', ''))
    ambos     = sum(1 for r in remanentes_post if "no existe en wms ni otm"  in r.get('resultado', ''))

    if fase_actual == 3:
        print(f"----------------------------------------")
        print(f" [RESULTADOS FASE {fase_actual}] WMS & OTM")
        print(f" - Ordenes analizadas:       {len(errores_pendientes)}")
        print(f" - Validadas con exito:      {exitos_post}  (ok en WMS Y en OTM)")
        print(f" - Pendientes de correccion: {len(remanentes_post)}")
        print(f"   |- Solo falla OTM:        {solo_otm}")
        print(f"   |- Solo falla WMS:        {solo_wms}")
        print(f"   |- Falla en ambos:        {ambos}")
        print(f"----------------------------------------\n")
    else:
        print(f"[RESUMEN FINAL] WMS & OTM")
        print(f" * Recuperadas en cierre: {exitos_post}  (ok en WMS Y en OTM)")
        print(f" * Fallidas definitivas:  {len(remanentes_post)}")
        print(f"   |- Solo falla OTM:     {solo_otm}")
        print(f"   |- Solo falla WMS:     {solo_wms}")
        print(f"   |- Falla en ambos:     {ambos}")
        print(f"----------------------------------------\n")
    generar_reporte_final(datos_finales, fase=fase_actual)
    return remanentes_post

def procesar_lote_errores(sesion, sesion_otm, fase_actual):
    print(f"\n ════════════════════════════════════════════════════════════")
    print(f"  [FASE {fase_actual}] REVALIDACION Y MONITOREO")
    print(f" ════════════════════════════════════════════════════════════")
    datos_actuales = cargar_estado()
    errores_pendientes = [item for item in datos_actuales if item.get('etiqueta_final') == "NO MAL ERROR"]
    
    if not errores_pendientes:
        generar_reporte_final(datos_actuales, fase=fase_actual)
        return []

    lote_trabajo = []
    # ®Carlos Alfonso Ortega Molina®
    autor_firma_mid = "®Carlos Alfonso Ortega Molina®"
    for err in errores_pendientes:
        orden_orig = err['orden_buscada']
        if fase_actual == 2 and len(orden_orig) == 11:
            orden_nueva = "0" + orden_orig
            print(f"regla 11: transformando {orden_orig} -> {orden_nueva}")
            lote_trabajo.append((orden_nueva, err['sales_check']))
        else:
            lote_trabajo.append((orden_orig, err['sales_check']))

    print(f"re-validando {len(lote_trabajo)} ordenes...")
    tamanio_lote = 40
    for i in range(0, len(lote_trabajo), tamanio_lote):
        segmento = lote_trabajo[i:i + tamanio_lote]
        resultados = consultar_lotes_ordenes(segmento, sesion, sesion_otm)
        with bloqueo_fichero:
            with open(archivo_control, "a") as f_json:
                for res in resultados:
                    f_json.write(json.dumps(res) + "\n")

    datos_finales = cargar_estado()
    generar_reporte_final(datos_finales, fase=fase_actual)
    return [item for item in datos_finales if item.get('etiqueta_final') == "NO MAL ERROR"]

def generar_reporte_final(datos_completos, fase=1):
    if not datos_completos: return
    leyendas = {
        1: "Finalizacion Verificada y validada",
        2: "Segunda validacion de errores",
        3: "Tercera validacion tras reproceso GNX",
        4: "Cuarta y ultima validacion"
    }
    leyenda = leyendas.get(fase, f"Finalizacion Fase {fase}")
    
    print(f"\n[REPORTE] --- Iniciando protocolo de cierre (Fase {fase}) ---")
    
    df = pd.DataFrame(datos_completos)
    columnas_orden = [
        # datos de identificacion de la orden
        'orden_buscada', 'sales_check', 'etiqueta_final', 'resultado',
        # datos de respuesta de wms oracle cloud
        'wms_estatus_id', 'wms_id_interno', 'wms_order_nbr_api', 'wms_bodega',
        # datos de respuesta de otm oracle transportation
        'otm_order_release_xid',
        # datos de auditoria y fechas
        'fecha_hora_proceso', 'fecha_validacion_original'
    ]
    df = df.reindex(columns=[c for c in columnas_orden if c in df.columns])
    # renombrar columnas para que el excel sea legible y profesional
    mapa_nombres = {
        'orden_buscada':           'Order Release (GNX)',
        'sales_check':             'Sales Check (GNX)',
        'etiqueta_final':          'Estatus Final',
        'resultado':               'Resultado Validacion',
        'wms_estatus_id':          '[WMS] Status ID',
        'wms_id_interno':          '[WMS] ID Interno',
        'wms_order_nbr_api':       '[WMS] Order Nbr API',
        'wms_bodega':              '[WMS] Bodega / Facility',
        'otm_order_release_xid':   '[OTM] Order Release XID',
        'fecha_hora_proceso':      'Fecha/Hora Proceso',
        'fecha_validacion_original': 'Fecha Validacion Original'
    }
    df.rename(columns={k: v for k, v in mapa_nombres.items() if k in df.columns}, inplace=True)
    df.to_excel(archivo_salida, index=False)
    
    lista_errores = [item for item in datos_completos if item.get('etiqueta_final') == "NO MAL ERROR"]
    conteo_errores_maestro = len(lista_errores)
    
    modo_error = "a" 
    if fase == 1: modo_error = "w"

    with open(archivo_errores, modo_error) as f_err:
        f_err.write(f"\n--- reporte de fase {fase} ({datetime.now().strftime('%H:%M:%S')}) ---\n")
        for err in lista_errores:
            carlos = err.get('orden_buscada', 'n/a')
            motivo = err.get('resultado', 'error')
            orden_limpia = str(carlos).strip()
            if len(orden_limpia) == 11 and fase >= 2: carlos = "0" + orden_limpia
            f_err.write(f"carlos: {carlos} | motivo: {motivo}\n")
            
    with open(archivo_errores, "r") as f_ver:
        texto = f_ver.read()
        bloque = texto.split(f"--- reporte de fase {fase}")[-1]
        lineas_reporte = len([l for l in bloque.splitlines() if "carlos:" in l])
        
    print(f"verificacion dual fase {fase}: maestro({conteo_errores_maestro}) vs reporte({lineas_reporte})")

    if conteo_errores_maestro == lineas_reporte:
        sh_final = f"\n{datetime.now().strftime('%d/%m/%Y %H:%M:%S')} {leyenda}\n"
        for arch in [archivo_texto_procesado, archivo_errores]:
            if os.path.exists(arch):
                with open(arch, "a") as f_app: f_app.write(sh_final)

def enviar_correo(total_gnx, total_hoy, validados, lista_pendientes, datos_todos=None):
    archivo_chk_correo = f".correo_enviado_i68_{fecha_hoy}"
    if os.path.exists(archivo_chk_correo):
        print(f" [i] El correo ya fue enviado previamente el dia de hoy ({archivo_chk_correo}).")
        return

    pendientes_fin = len(lista_pendientes)
    fecha_format = datetime.now().strftime("%d/%m/%Y")
    hora_format = datetime.now().strftime("%H:%M:%S")
    # desglose de pendientes por sistema para mostrar en correo
    cnt_solo_otm = sum(1 for r in lista_pendientes if "existe en wms pero no otm" in r.get('resultado', ''))
    cnt_solo_wms = sum(1 for r in lista_pendientes if "existe en otm pero no wms" in r.get('resultado', ''))
    cnt_ambos    = sum(1 for r in lista_pendientes if "no existe en wms ni otm"  in r.get('resultado', ''))
    # calcular errores de conexion o fallos de las apis de oracle
    cnt_tecnicos = pendientes_fin - (cnt_solo_otm + cnt_solo_wms + cnt_ambos)
    _ = pendientes_fin  # referenciado en html
    
    intentos_mail = 0
    max_intentos_mail = 3
    exito_mail = False

    while intentos_mail < max_intentos_mail and not exito_mail:
        intentos_mail += 1
        print(f"\n [≫] Preparando notificacion i68 por correo (Intento {intentos_mail}/{max_intentos_mail})...")
        try:
            if not os.path.exists("destinatarios.txt"):
                print(" [!] Error: No existe destinatarios.txt")
                return

            dests = open("destinatarios.txt").read().strip().replace("\n", ";")
            if not dests:
                print(" [!] No se encontraron destinatarios en destinatarios.txt")
                return

            pythoncom.CoInitialize()
            time.sleep(2)
            out = win32com.client.Dispatch("Outlook.Application")
            mail = out.CreateItem(0)
            mail.Subject = f"[REPORTE] Match Automatizado i68 (WMS & OTM) - {fecha_format}"
            mail.To = dests
            mail.SentOnBehalfOfName = "ortegac@sanborns.com.mx"

            color_pendientes = "color:red;" if pendientes_fin > 0 else "color:green;"
            bg_header = "#003366" 

            html_body = f"""<html>
<head>
<style>
    body {{ font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 11pt; color: #333; }}
    .container {{ width: 80%; margin: auto; }}
    .header {{ background-color: {bg_header}; color: white; padding: 15px; text-align: center; border-radius: 5px 5px 0 0; }}
    .content {{ padding: 20px; border: 1px solid #ddd; border-top: none; background-color: #f9f9f9; }}
    table {{ width: 100%; border-collapse: collapse; margin-top: 15px; background-color: white; }}
    th, td {{ border: 1px solid #ddd; padding: 10px; text-align: left; }}
    th {{ background-color: #f2f2f2; font-weight: bold; color: {bg_header}; }}
    .resumen {{ width: 60%; }}
    .alert {{ font-weight: bold; {color_pendientes} }}
    .footer {{ font-size: 9pt; color: #777; margin-top: 20px; text-align: center; }}
    .fila-par {{ background-color: #fcfcfc; }}
</style>
</head>
<body>
<div class="container">
    <div class="header">
        <h2>Reporte de Validacion de Pedidos i68</h2>
        <p>Fecha: {fecha_format} | Hora: {hora_format}</p>
    </div>
    <div class="content">
        <p>Se ha completado el ciclo de validación automatizada entre <b>GNX</b>, <b>WMS Oracle Cloud</b> y <b>OTM</b>.</p>
        
        <h3>Resumen Ejecutivo</h3>
        <table class="resumen">
          <tr><th>Concepto</th><th>Cantidad</th></tr>
          <tr><td>Total registros analizados (GNX)</td><td>{total_gnx:,}</td></tr>
          <tr><td>Pedidos a validar hoy</td><td>{total_hoy:,}</td></tr>
          <tr><td>Validados correctamente (WMS y OTM)</td><td style="color:green;font-weight:bold">{validados:,}</td></tr>
          <tr><td>Pendientes sin validacion final</td><td class="alert">{pendientes_fin}</td></tr>
        </table>
        <br>
        <h3>Desglose por Sistema</h3>
        <table class="resumen">
          <tr><th>Sistema</th><th>Fallidos</th><th>Descripcion</th></tr>
          <tr><td><b>Solo OTM</b></td><td style="color:#cc6600;font-weight:bold">{cnt_solo_otm}</td><td>Existe en WMS pero no en OTM</td></tr>
          <tr><td><b>Solo WMS</b></td><td style="color:#cc6600;font-weight:bold">{cnt_solo_wms}</td><td>Existe en OTM pero no en WMS</td></tr>
          <tr><td><b>Ambos (WMS y OTM)</b></td><td style="color:#990000;font-weight:bold">{cnt_ambos}</td><td>No existe en ninguno de los dos</td></tr>
        </table>"""

            if pendientes_fin > 0:
                html_body += f"""<br>
        <h3 style="color: #990000;">Detalle de Pendientes Definitivos</h3>
        <p>Los siguientes pedidos no pudieron ser validados tras el protocolo de reproceso:</p>
        <table>
          <tr>
            <th style="width: 30%;">Sales Check</th>
            <th style="width: 30%;">Order Release</th>
            <th style="width: 40%;">Causa del Fallo</th>
          </tr>"""
                for i, item in enumerate(lista_pendientes):
                    clase_fila = 'class="fila-par"' if i % 2 == 0 else ""
                    sales_p = item.get('sales_check', 'n/a')
                    release_p = item.get('orden_buscada', 'n/a')
                    motivo_p = item.get('resultado', 'error desconocido')
                    if "no existe en wms ni otm" in motivo_p:
                        causa = "Faltante en WMS y OTM"
                    elif "existe en wms pero no otm" in motivo_p:
                        causa = "Faltante en OTM"
                    elif "existe en otm pero no wms" in motivo_p:
                        causa = "Faltante en WMS"
                    else:
                        causa = motivo_p.upper()
                        
                    html_body += f"<tr {clase_fila}><td>{sales_p}</td><td>{release_p}</td><td><span style='color: #990000; font-weight: 500;'>{causa}</span></td></tr>"
                html_body += "</table>"
            else:
                html_body += """<br><div style="padding: 15px; background-color: #e6fffa; border: 1px solid #38b2ac; color: #234e52; border-radius: 5px;">
                    <b>¡Éxito total!</b> No se encontraron pedidos pendientes para el día de hoy. Todos validados en WMS y OTM.
                </div>"""

            html_body += f"""<br>
        <div style="padding: 15px; background-color: #f9f9f9; border-left: 4px solid {bg_header}; font-size: 0.95em;">
            <strong>Nota Informativa:</strong> El proceso ha concluido todas sus fases (0 a 4). 
            Se adjunta el log detallado para auditoría de tiempos de respuesta SSH/API.
        </div>
        <div class="footer">
            <p>Este es un correo generado automaticamente por el Sistema de Match i68.<br>
            Autor: Carlos Alfonso Ortega Molina</p>
        </div>
    </div>
</div>
</body>
</html>"""
            mail.HTMLBody = html_body
            sys.stdout.flush()
            if os.path.exists(archivo_log): mail.Attachments.Add(os.path.abspath(archivo_log))
            mail.Send()
            print(" [✓] Mail enviado satisfactoriamente con diseño HTML 2.0.")
            exito_mail = True
            open(archivo_chk_correo, "w").close()
            
        except Exception as e:
            print(f" [X] Error en Intento {intentos_mail}: {e}")
            if intentos_mail < max_intentos_mail:
                print(" [i] Reintentando en 10 segundos...")
                time.sleep(10)
            else:
                print(" [!] Se agotaron los intentos de envio de correo.")
        finally:
            try: pythoncom.CoUninitialize()
            except: pass

def ejecutar_verificacion():
    if not os.path.exists(archivo_entrada): return

    with open(archivo_entrada, "r") as f:
        f_lineas = f.readlines()
        mapa_ordenes_completas = {}
        mapa_almacen_origen = {}
        for carlos in f_lineas:
            linea_limpia = carlos.strip()
            if linea_limpia:
                partes = linea_limpia.split('|')
                if len(partes) >= 2:
                    carlos = partes[0].strip()
                    mapa_ordenes_completas[carlos] = partes[1].strip()
                    # si el archivo tiene info de bodega la guardamos formato ordensalesbodega o similar
                    if len(partes) >= 3:
                        mapa_almacen_origen[carlos] = partes[2].strip()

    print("\n ══════════════════════════════════════════════════════════════════════")
    print("  [SISTEMA] INICIALIZANDO EJECUCION DE VERIFICACION")
    print(" ══════════════════════════════════════════════════════════════════════\n")

    # control ferreo renombrado local inmediato del archivo de entrada tras carga exitosa en memoria
    try:
        arch_rename = f"cliyclitie_{fecha_hoy}.txt"
        if os.path.exists(archivo_entrada):
            if os.path.exists(arch_rename): os.remove(arch_rename)
            os.rename(archivo_entrada, arch_rename)
            print(f" [✓] Control Ferreo: {archivo_entrada} archivado de inmediato como {arch_rename}.")
    except Exception as e:
        print(f" [!] Aviso: No se pudo archivar localmente el archivo de entrada: {e}")

    todas_keys = list(mapa_ordenes_completas.keys())
    total_gnx = len(todas_keys)
    res_previos = cargar_estado()
    hechos_hoy = {i['orden_buscada'] for i in res_previos}
    pendientes = [o for o in todas_keys if o not in hechos_hoy]

    cache = cargar_historial_ayer()
    para_api = []
    
    escritor = threading.Thread(target=hilo_guardado_continuo, daemon=True)
    escritor.start()
    
    for carlos in pendientes:
        k12 = "0" + carlos if len(carlos) == 11 else carlos
        hit = cache.get(carlos) or cache.get(k12)
        
        # recuperar bodega origen si existe
        bodega_orig = mapa_almacen_origen.get(carlos, "n/a")

        if hit:
            d = hit['data'].copy()
            d['orden_buscada'] = carlos
            d['sales_check'] = mapa_ordenes_completas[carlos]
            if d.get('bodega') == 'n/a': d['bodega'] = bodega_orig
            d['fecha_hora_proceso'] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            d['etiqueta_final'] = f"YA VALIDADA ANTES [{hit['fecha_validacion']}]"
            d['fecha_validacion_original'] = hit['fecha_validacion']
            cola_resultados.put(d)
        else:
            para_api.append(carlos)
            
    if para_api:
        print(f"\n ════════════════════════════════════════════════════════════")
        print(f"  [FASE 1] VALIDACION BATCH WMS & OTM (Lotes de 40)")
        print(f"  Pendientes: {len(para_api)} | Total hoy: {len(todas_keys)}")
        print(f" ════════════════════════════════════════════════════════════")
        with requests.Session() as sesion, requests.Session() as sesion_otm:
             sesion_otm.auth = (OTM_USUARIO, OTM_CONTRASENA)
             estrategia_reintento = Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
             adaptador = HTTPAdapter(pool_connections=5, pool_maxsize=5, max_retries=estrategia_reintento)
             sesion.mount("https://", adaptador)
             sesion.auth = (USUARIO, CONTRASENA)
             
             verificar_reentrada(res_previos, sesion, sesion_otm)
             
             lotes = []
             for i in range(0, len(para_api), 40):
                 seg = para_api[i:i+40]
                 lotes.append([(o, mapa_ordenes_completas[o]) for o in seg])
                 
             procesadas = len(hechos_hoy) + len(pendientes) - len(para_api)
             
             with ThreadPoolExecutor(max_workers=5) as ex:
                 futuros = {ex.submit(consultar_lotes_ordenes, l, None, None): i for i, l in enumerate(lotes)}
                 idx = 1
                 for f in as_completed(futuros):
                     res = f.result()
                     for r in res: cola_resultados.put(r)
                     procesadas += len(res)
                     print(f"lotes wms/otm: [{idx}/{len(lotes)}] | ordenes: {procesadas}/{len(todas_keys)}...")
                     idx += 1
                     
    cola_resultados.put(None)
    escritor.join()
    
    # total_hoy  ordenes que realmente fueron a la api las de cache no cuentan
    total_hoy = len(para_api)
    
    def despachar_correo():
        d_fin = cargar_estado()
        remanentes_hoy = [i for i in d_fin if i.get('etiqueta_final') == "NO MAL ERROR"]
        validados = total_hoy - len(remanentes_hoy)
        if validados < 0: validados = 0
        enviar_correo(total_gnx, total_hoy, validados, remanentes_hoy, datos_todos=d_fin)
    
    d_f1 = cargar_estado()
    generar_reporte_final(d_f1, fase=1)
    
    errores_f2 = []
    with requests.Session() as s, requests.Session() as s_otm:
        s_otm.auth = (OTM_USUARIO, OTM_CONTRASENA)
        s.auth = (USUARIO, CONTRASENA)
        errores_f2 = procesar_lote_errores(s, s_otm, fase_actual=2)
        
    if errores_f2:
        print(f" [i] {len(errores_f2)} incidencias detectadas. Iniciando protocolo FTP/SSH...")
        exito_gnx = subir_ftp_y_ejecutar_ssh(errores_f2)
            
        if exito_gnx:
            errores_f3 = []
            with requests.Session() as s3, requests.Session() as s3_otm:
                s3_otm.auth = (OTM_USUARIO, OTM_CONTRASENA)
                s3.auth = (USUARIO, CONTRASENA)
                errores_f3 = esperar_y_revalidar(s3, s3_otm, 3, errores_f2, 20)
                
            if errores_f3:
                with requests.Session() as s4, requests.Session() as s4_otm:
                    s4_otm.auth = (OTM_USUARIO, OTM_CONTRASENA)
                    s4.auth = (USUARIO, CONTRASENA)
                    errores_f4 = esperar_y_revalidar(s4, s4_otm, 4, errores_f3, 30)
                    
                if errores_f4:
                    print("\nAVISO: El proceso termino pero persisten errores.")
                    despachar_correo()
                else:
                    despachar_correo()
            else:
                despachar_correo()
        else:
            despachar_correo()
    else:
        arch_rename = f"cliyclitie_{fecha_hoy}.txt"
        if os.path.exists(archivo_entrada): 
            if os.path.exists(arch_rename): os.remove(arch_rename)
            try: os.rename(archivo_entrada, arch_rename)
            except: pass
        despachar_correo()

if __name__ == "__main__":
    # limpiar semaforos antiguos si existen para evitar bloqueos
    if os.path.exists(".chk_ssh_i68"):
        try: os.remove(".chk_ssh_i68")
        except: pass

    servicio_log = duplicador_salida(archivo_log)
    sys.stdout = servicio_log
    print(f"--- inicializando sistema automatizado i68 (v11 GNX) [{fecha_hoy}] ---\n")
    if obtener_credenciales():
        if not os.path.exists(archivo_entrada):
            if not fase_cero_ftp():
                print(" [!] ALERTA: No se pudo descargar archivo de hoy por FTP.")
        
        if os.path.exists(archivo_entrada):
            # flujo normal hay archivo de entrada procesar completo
            ejecutar_verificacion()
        else:
            # recuperacion revisar si el proceso termino pero fallo el correo
            d_prev = cargar_estado()
            remanentes = [i for i in d_prev if i.get('etiqueta_final') == "NO MAL ERROR"]
            archivo_chk_correo = f".correo_enviado_i68_{fecha_hoy}"

            if len(d_prev) > 0 and not os.path.exists(archivo_chk_correo):
                print(f" [!] ALERTA DE RECUPERACION: Detectados datos procesados sin confirmacion de correo enviado.")
                total_gnx = len(d_prev)
                total_hoy = len([i for i in d_prev if i.get('etiqueta_final') != "YA VALIDADA ANTES"])
                validados = total_hoy - len(remanentes)
                if validados < 0: validados = 0
                print(f" [!] Re-intentando despacho de notificacion pendiente...")
                d_rec = cargar_estado()
                remanentes = [i for i in d_rec if i.get('etiqueta_final') == "NO MAL ERROR"]
                enviar_correo(total_gnx, total_hoy, validados, remanentes, datos_todos=d_rec)
                print(" [✓] Recuperacion de correo completada.")

            # sin archivo de entrada revisar si hay remanentes pendientes de fases anteriores
            elif len(remanentes) > 0 and os.path.exists(".chk_ssh_i68"):
                # ya se ejecuto el ssh solo falta esperar y revalidar fase 34
                print(f" [!] ALERTA DE RECUPERACION: {len(remanentes)} ordenes pendientes tras reproceso GNX.")
                total_gnx = len(d_prev)
                total_hoy = len([i for i in d_prev if i.get('etiqueta_final') != "YA VALIDADA ANTES"])
                def despachar_correo():
                    d_fin = cargar_estado()
                    remanentes_hoy = [i for i in d_fin if i.get('etiqueta_final') == "NO MAL ERROR"]
                    validados = total_hoy - len(remanentes_hoy)
                    if validados < 0: validados = 0
                    enviar_correo(total_gnx, total_hoy, validados, remanentes_hoy, datos_todos=d_fin)
                with requests.Session() as s_rec, requests.Session() as s_rec_otm:
                    s_rec_otm.auth = (OTM_USUARIO, OTM_CONTRASENA)
                    s_rec.auth = (USUARIO, CONTRASENA)
                    errores_rec3 = esperar_y_revalidar(s_rec, s_rec_otm, 3, remanentes, 20)
                    if errores_rec3:
                        esperar_y_revalidar(s_rec, s_rec_otm, 4, errores_rec3, 30)
                despachar_correo()
            elif len(remanentes) > 0 and not os.path.exists(".chk_ssh_i68"):
                # nunca se llego al ssh intentar reproceso completo
                print(f" [!] ALERTA DE RECUPERACION: {len(remanentes)} ordenes pendientes sin reproceso GNX. Ejecutando...")
                exito_gnx = subir_ftp_y_ejecutar_ssh(remanentes)
                if exito_gnx:
                    open(".chk_ssh_i68", "w").close()
                    total_gnx = len(d_prev)
                    total_hoy = len([i for i in d_prev if i.get('etiqueta_final') != "YA VALIDADA ANTES"])
                    def despachar_correo():
                        d_fin = cargar_estado()
                        remanentes_hoy = [i for i in d_fin if i.get('etiqueta_final') == "NO MAL ERROR"]
                        validados = total_hoy - len(remanentes_hoy)
                        if validados < 0: validados = 0
                        enviar_correo(total_gnx, total_hoy, validados, remanentes_hoy, datos_todos=d_fin)
                    with requests.Session() as s_rec, requests.Session() as s_rec_otm:
                        s_rec_otm.auth = (OTM_USUARIO, OTM_CONTRASENA)
                        s_rec.auth = (USUARIO, CONTRASENA)
                        errores_rec3 = esperar_y_revalidar(s_rec, s_rec_otm, 3, remanentes, 20)
                        if errores_rec3:
                            esperar_y_revalidar(s_rec, s_rec_otm, 4, errores_rec3, 30)
                    despachar_correo()
            else:
                print("proceso del dia completado previamente. nada pendiente.")
    else:
        print("Error: Faltan credenciales validas en json (.acceso_wms o .acceso_gnx)")
    sys.stdout = servicio_log.consola
    servicio_log.close()
# ®Carlos Alfonso Ortega Molina®


