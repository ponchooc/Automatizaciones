# ®Carlos Alfonso Ortega Molina®
# importacion de modulos del sistema y bibliotecas para el procesamiento de ordenes de compra
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

# establecer la fecha actual formateada para los nombres de archivo
fecha_hoy = datetime.now().strftime("%d%m%Y")

# definicion de la direccion base de la api de wms para shipment
URL_BASE = "https://e6.wms.ocs.oraclecloud.com/sears2/wms/lgfapi/v10/entity/ib_shipment/"
USUARIO = ""
CONTRASENA = ""

# servidor y accesos para gnx remoto
GNX_HOST = "140.240.11.1"
GNX_USER = ""
GNX_PASS = ""

# flujo operativo ftp de archivos oc validacion api por bodega reproceso gnx cierre
# este script procesa ordenes de compra del dia
# combina dos archivos de entrada
# valida por bodega y luego cierra con correo

class duplicador_salida:
    def __init__(self, nombre_archivo):
        self.consola = sys.stdout
        self.archivo_log = open(nombre_archivo, "a", encoding="utf-8")
    def write(self, mensaje):
        self.consola.write(mensaje)
        self.archivo_log.write(mensaje)
        # Forzar escritura inmediata en disco para log en tiempo real
        self.flush()
    def flush(self):
        self.consola.flush()
        self.archivo_log.flush()
    def close(self):
        self.archivo_log.close()

archivo_log = f"log_oc_{fecha_hoy}.txt"
archivo_entrada_val = "OC_val_N.txt"
archivo_entrada_ins = "OC_val_I.txt"
archivo_salida = f"reporte_oc_{fecha_hoy}.xlsx"
archivo_control = f"control_oc_{fecha_hoy}.json"
archivo_texto_procesado = f"oc_procesado_{fecha_hoy}.txt"
archivo_errores = f"errores_oc_{fecha_hoy}.txt"

# Salidas diarias del proceso OC para auditoria de estado y rastreo de incidencias
bloqueo_fichero = threading.Lock()
cola_resultados = queue.Queue()

# Carga credenciales de ambos mundos WMS y GNX como prerrequisito de ejecucion
def obtener_credenciales():
    # objetivo
    # preparar acceso para api ftp y ssh
    # salida
    # verdadero cuando existen ambas credenciales
    global USUARIO, CONTRASENA, GNX_USER, GNX_PASS
    wms_ok = False
    if os.path.exists(".acceso_wms"):
        try:
            with open(".acceso_wms", "r") as f:
                d = json.load(f)
                USUARIO = d.get("usuario")
                CONTRASENA = d.get("contrasena")
                if USUARIO and CONTRASENA: wms_ok = True
        except: pass

    gnx_ok = False
    if os.path.exists(".acceso_gnx"):
        try:
            with open(".acceso_gnx", "r") as f:
                d = json.load(f)
                GNX_USER = d.get("usuario")
                GNX_PASS = d.get("contrasena")
                if GNX_USER and GNX_PASS: gnx_ok = True
        except: pass
    # Solo se habilita el flujo cuando hay acceso tanto a fuente GNX como a destino WMS
    return wms_ok and gnx_ok

def fase_cero_ftp():
    # Intenta descargar ambas fuentes OC del dia basta con recuperar al menos una
    # busca dos archivos de entrada
    # si llega al menos uno el flujo puede iniciar
    print(f"\n ══════════════════════════════════════════════════════════════════════")
    print(f"  [FASE 0] AUTO-ABASTECIMIENTO DE MATERIA PRIMA OC (FTP)")
    print(f" ══════════════════════════════════════════════════════════════════════")
    print(f" [≫] Conectando a {GNX_HOST}...")
    descargados = 0
    try:
        with ftplib.FTP(GNX_HOST, encoding='latin-1') as ftp:
            ftp.login(GNX_USER, GNX_PASS)
            ruta = "/gnx_prod/manto/desa/trabajo/sears/ORACLE"
            ftp.cwd(ruta)
            for archivo_oc in [archivo_entrada_val, archivo_entrada_ins]:
                try:
                    if archivo_oc in ftp.nlst():
                        res = ftp.voidcmd(f"MDTM {archivo_oc}")
                        fecha_srv = res.split()[1][:8]
                        if fecha_srv == datetime.now().strftime("%Y%m%d"):
                            with open(archivo_oc, "wb") as f_loc:
                                ftp.retrbinary(f"RETR {archivo_oc}", f_loc.write)
                            tamanio = os.path.getsize(archivo_oc)
                            print(f" [✓] Archivo descargado: {archivo_oc} ({tamanio} bytes)")
                            descargados += 1
                        else:
                            print(f" [!] ALERTA: {archivo_oc} no es de hoy.")
                except: pass
    except Exception as e:
        print(f" [X] Error en FTP: {e}")
    return descargados > 0

def subir_ftp_y_ejecutar_ssh(errores_pendientes):
    # Genera i04 txt con incidencias lo publica en GNX y lanza el reproceso OC remoto
    # entrada
    # lista de ordenes con error
    # salida
    # verdadero cuando la ejecucion remota termina
    print(f" ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print(f"  [GNX] INICIANDO PROTOCOLO REMOTO PARA {len(errores_pendientes)} INCIDENCIAS OC")
    print(f" ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

    archivo_carga = "i04.txt"
    try:
        with open(archivo_carga, "w", newline='\n') as f_unl:
            for error in errores_pendientes:
                sales = error.get('sales_check', '').strip()
                if sales:
                    f_unl.write(f"{sales}\n")
        
        # Historial de incidencias local
        archivo_backup = f"i04_{fecha_hoy}.txt"
        import shutil
        shutil.copy(archivo_carga, archivo_backup)
        print(f" [✓] Archivo de carga generado: {archivo_carga} (Backup: {archivo_backup})")
    except Exception as e:
        print(f"error generando archivo de carga: {e}")
        return False

    print(f" [≫] Subiendo {archivo_carga} a {GNX_HOST} via FTP Python (Modo Binario)...")
    try:
        with ftplib.FTP(GNX_HOST) as ftp:
            ftp.login(GNX_USER, GNX_PASS)
            ftp.voidcmd("TYPE I")  # Transferencia férrea binaria
            # Ruta donde GNX espera el archivo de entrada i04 txt
            ftp.cwd("/respaldo_migracion/reportes_gnx")
            with open(archivo_carga, "rb") as f_sub:
                ftp.storbinary(f"STOR {archivo_carga}", f_sub)
            print(f" [✓] Transferencia confirmada en /respaldo_migracion/reportes_gnx.")
    except Exception as e:
        print(f" [X] Error en transferencia FTP: {e}")
        return False

    print(f" [≫] Estableciendo conexion SSH (Paramiko) para reproceso OC...")
    try:
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(GNX_HOST, username=GNX_USER, password=GNX_PASS)
        print(" [✓] Conexion establecida. Iniciando eco del servidor.")
        sh = ssh.invoke_shell(term='vt100', width=220, height=50)
        sh.settimeout(300)

        def leer_hasta_prompt(max_seg=10):
            """lee todo lo que llega del canal hasta que no haya mas datos por max_seg segundos"""
            # captura salida de shell remoto para log local
            salida = ""
            inicio = __import__('time').time()
            while True:
                if sh.recv_ready():
                    bloque = sh.recv(4096).decode('latin-1', 'ignore')
                    print(bloque, end="", flush=True)
                    salida += bloque
                    inicio = __import__('time').time()
                elif __import__('time').time() - inicio > max_seg:
                    break
                else:
                    __import__('time').sleep(0.3)
            return salida

        def run_cmd(cmd, desc="", espera_seg=5):
            if desc: print(desc)
            if cmd:
                time.sleep(1) # Pausa de seguridad para el buffer
                sh.send(cmd + "\n")
            return leer_hasta_prompt(espera_seg)

        run_cmd("1", "\n>>> [AUTO] Seleccionando Ambiente 1 (SEARS)...", 5)
        run_cmd("1", ">>> [AUTO] Seleccionando Opcion 1 (SEARS)...", 5)
        # Ruta donde reside el ejecutable sh
        run_cmd("cd /gnx_prod/manto/desa/trabajo/sears/carlos_ortega", ">>> [AUTO] Seleccionando ruta de trabajo absoluta...", 5)

        print(">>> [AUTO] Ejecutando script de reproceso OC... (esperando conclusion total)")
        sh.send("./automatizador_ordenes_compra.sh\n")
        # Aumentamos el tiempo de espera a 600 segundos 10 min para asegurar que procesos pesados de DB terminen
        leer_hasta_prompt(max_seg=600)

        print("\n>>> [AUTO] Esperando a que el prompt se libere ('>')...")
        leer_hasta_prompt(max_seg=30) # Asegura la recepcion del prompt final del script .sh

        print(">>> [AUTO] Saliendo del shell ('exit')...")
        sh.send("exit\n")
        time.sleep(2) # Pausa rigurosa para que el sistema procese la salida y muestre el menu

        print(">>> [AUTO] Enviando 'f' para finalizar sesion en el menu principal...")
        sh.send("f\n")
        time.sleep(1)

        leer_hasta_prompt(max_seg=5)
        ssh.close()
        return True
    except Exception as e:
        print(f"error critico ssh: {e}")
        return False

def cargar_historial_ayer():
    # Habilita cache de exitos previos para evitar reprocesar ordenes ya confirmadas
    # solo guarda registros ya validados
    # sirve para saltar trabajo repetido
    ayer = datetime.now() - timedelta(days=1)
    str_yesterday = ayer.strftime("%d%m%Y")
    archivo_ayer = f"control_oc_{str_yesterday}.json"
    print(f" [≫] Buscando historial de validacion: {archivo_ayer}")
    if not os.path.exists(archivo_ayer):
        return {}

    print(f"cargando memoria historica desde: {archivo_ayer}")
    cache_historica = {}
    try:
        with open(archivo_ayer, "r") as f:
            for linea in f:
                try:
                    linea = linea.strip()
                    if not linea or linea in ["[", "]", ","]: continue
                    obj = json.loads(linea)
                    if obj.get('etiqueta_final') == "OK VERIFICADO API" or "YA VALIDADA ANTES" in obj.get('etiqueta_final', ''):
                        orden = str(obj.get('orden_buscada'))
                        fecha_val = obj.get('fecha_validacion_original', obj.get('fecha_hora_proceso', 'desconocida'))
                        if len(orden) == 11: orden = "0" + orden
                        cache_historica[orden] = {'data': obj, 'fecha_validacion': fecha_val}
                except: continue
        print(f"memoria cargada: {len(cache_historica)} ordenes ya validadas previamente.")
    except Exception as e:
        print(f"aviso: error al leer historial {archivo_ayer}: {e}")
    return cache_historica

def cargar_estado():
    # Devuelve estado consolidado del dia deduplicado por orden
    # recorre control diario y conserva ultima version por orden
    mapa_recuperado = {}
    if os.path.exists(archivo_control):
        try:
            with open(archivo_control, "r") as f:
                for num_linea, linea in enumerate(f, 1):
                    # Compatibilidad con historicos no estrictamente NDJSON
                    contenido = linea.strip()
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
    # Persiste de forma serializada los resultados asincronos del procesamiento por lotes
    # una sola rutina escribe en disco
    # evita conflictos por escritura simultanea
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
                    orden = resultado.get('orden_buscada', 'n/a')
                    sales = resultado.get('sales_check', 'n/a')
                    estado = resultado.get('etiqueta_final', 'NO MAL ERROR')
                    linea_txt = f"{orden}|{sales}|{estado}"
                    if "YA VALIDADA ANTES" in estado:
                        fecha_orig = resultado.get('fecha_validacion_original', 'N/A')
                        if " " in fecha_orig: fecha_orig = fecha_orig.split(" ")[0]
                        linea_txt += f"|VALIDADO EL DIA {fecha_orig}"
                    f_txt.write(linea_txt + "\n")
        except: pass
        cola_resultados.task_done()

def consultar_lotes_ordenes(lista_datos_entrada, sesion):
    # Consulta API de ib shipment por bodega y transforma respuesta a formato interno
    # entrada
    # lote con orden sales y bodega
    # salida
    # estructura uniforme para todas las fases
    # lista datos entrada orden sales bodega
    if not lista_datos_entrada: return []
    
    bodega_del_lote = lista_datos_entrada[0][2]
    ordenes_list = [o[0] for o in lista_datos_entrada]
    ordenes_str = ",".join(ordenes_list)
    resultados_finales_lote = []

    try:
        params = {
            # En OC la bodega es parte del criterio de busqueda no solo dato informativo
            'facility_id__code': bodega_del_lote,
            'company_id__code': 'GPOSAN',
            'shipment_nbr__in': ordenes_str,
            'fields': 'shipment_nbr,status_id',
            'page_size': 100
        }
        respuesta = sesion.get(URL_BASE, params=params, timeout=25)
        ahora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        encontrados = {}
        if respuesta.status_code == 200:
            datos = respuesta.json()
            for info in datos.get('results', []):
                o_api = info.get('shipment_nbr')
                if o_api: encontrados[o_api] = info

        for orden, sales, bodega in lista_datos_entrada:
            res_obj = {'orden_buscada': orden, 'sales_check': sales, 'bodega': bodega, 'fecha_hora_proceso': ahora}
            if orden in encontrados:
                info = encontrados[orden]
                res_obj.update({
                    'shipment_nbr_api': info.get('shipment_nbr'),
                    'estatus_id': info.get('status_id'),
                    'resultado': "exito en wms",
                    'etiqueta_final': "OK VERIFICADO API",
                    'fecha_validacion_original': ahora
                })
            else:
                # Estandar de remanente usado por fases 2 3 y 4
                res_obj.update({'resultado': "no existe en wms" if respuesta.status_code == 200 else "error api", 'etiqueta_final': "NO MAL ERROR"})
            resultados_finales_lote.append(res_obj)
        return resultados_finales_lote
    except Exception as e:
        return [{'orden_buscada': o[0], 'sales_check': o[1], 'bodega': o[2], 'resultado': f"fallo lote: {str(e)}", 'etiqueta_final': "NO MAL ERROR", 'fecha_hora_proceso': datetime.now().strftime("%d/%m/%Y %H:%M:%S")} for o in lista_datos_entrada]

def verificar_reentrada(resultados_previos, sesion):
    # Control de consistencia revalida un subconjunto reciente antes de nuevos lotes
    # evita arrastrar estados viejos en arranques con recuperacion
    if not resultados_previos: return resultados_previos
    candidatos = [r for r in resultados_previos if "YA VALIDADA ANTES" not in r.get('etiqueta_final', '')]
    if not candidatos: return resultados_previos
    num_a_verificar = min(3, len(candidatos))
    ultimas_3 = candidatos[-num_a_verificar:]
    indices_map = {r['orden_buscada']: i for i, r in enumerate(resultados_previos) if r in ultimas_3}
    for orden_res in ultimas_3:
        idx = indices_map.get(orden_res['orden_buscada'])
        if idx is not None:
            # Necesitamos la bodega para re consultar
            datos_entrada = [(orden_res['orden_buscada'], orden_res['sales_check'], orden_res.get('bodega', 'n/a'))]
            re_chequeo = consultar_lotes_ordenes(datos_entrada, sesion)[0]
            if re_chequeo['etiqueta_final'] != orden_res.get('etiqueta_final'):
                resultados_previos[idx] = re_chequeo
    return resultados_previos

def esperar_y_revalidar(sesion, fase_actual, errores_pendientes, minutos):
    # Ejecuta ventana de espera revalidacion final con checkpoints para recuperacion idempotente
    # fase tres y cuatro usan la misma base
    # cambia el tiempo y el encabezado
    archivo_checkpoint = f".fase{fase_actual}_iniciada_oc_{fecha_hoy}"
    segundos = minutos * 60
    ahora = datetime.now()
    luego = ahora + timedelta(seconds=segundos)
    
    if fase_actual == 3:
        titulo = "VALIDACION POST-REPROCESO OC"
        label_X = "Ordenes a monitorear"
    else:
        titulo = "REVISION FINAL DEFINITIVA OC"
        label_X = "Ordenes en Revision Final"

    print(f"\n════════════════════════════════════════════════════════════")
    print(f"  [FASE {fase_actual}] {titulo} ({minutos} min)")
    print(f"════════════════════════════════════════════════════════════")
    print(f" | {label_X}:    {len(errores_pendientes)}")
    
    # INTELIGENCIA Verificar si esta espera ya se cumplio en una ejecucion previa
    if os.path.exists(archivo_checkpoint):
        print(f" | [!] AVISO: El tiempo de espera de la Fase {fase_actual} ya se cumplio anteriormente.")
        print(f" | [!] Saltando pausa de {minutos} min para proceder directo a la validacion API.")
    else:
        print(f" | Pausa de seguridad:      {segundos} seg ({minutos} min)")
        print(f" | Inicio: {ahora.strftime('%H:%M:%S')} | Re-consulta: {luego.strftime('%H:%M:%S')}")
        print(f"------------------------------------------------------------\n")
        # Crear checkpoint antes de dormir
        try: open(archivo_checkpoint, "w").close()
        except: pass
        time.sleep(segundos)
        
    print(f"------------------------------------------------------------")
    print(f"[API] Ejecutando consulta de verificacion Fase {fase_actual}...\n")

    # Mantiene la misma estrategia de particion por bodega utilizada en fase 1
    lotes_por_bodega = {}
    for err in errores_pendientes:
        b = err.get('bodega', 'n/a')
        if b not in lotes_por_bodega: lotes_por_bodega[b] = []
        lotes_por_bodega[b].append((err['orden_buscada'], err['sales_check'], b))

    for b_key, items_bodega in lotes_por_bodega.items():
        for i in range(0, len(items_bodega), 100):
            segmento = items_bodega[i:i + 100]
            resultados = consultar_lotes_ordenes(segmento, sesion)
            with bloqueo_fichero:
                with open(archivo_control, "a") as f_json:
                    for res in resultados:
                        f_json.write(json.dumps(res) + "\n")

    datos_finales = cargar_estado()
    remanentes_post = [item for item in datos_finales if item.get('etiqueta_final') == "NO MAL ERROR"]
    exitos_post = len(errores_pendientes) - len(remanentes_post)
    if exitos_post < 0: exitos_post = 0

    if fase_actual == 3:
        print(f"----------------------------------------")
        print(f" [RESULTADOS FASE 3]")
        print(f" - Ordenes analizadas:      {len(errores_pendientes)}")
        print(f" - Validadas con exito:     {exitos_post}")
        print(f" - Pendientes de correccion: {len(remanentes_post)}")
        print(f"----------------------------------------\n")
    else:
        print(f"[RESUMEN FINAL]")
        print(f" * Recuperadas en cierre: {exitos_post}")
        print(f" * Fallidas definitivas:   {len(remanentes_post)}")
        print(f"----------------------------------------\n")
    generar_reporte_final(datos_finales, fase=fase_actual)
    return remanentes_post

def procesar_lote_errores(sesion, fase_actual):
    # Reintenta exclusivamente ordenes con etiqueta de error agrupando por bodega
    # esta fase reduce remanentes antes del reproceso remoto
    titulo = "REFUERZO DE ERRORES OC" if fase_actual == 2 else "REVALIDACION Y MONITOREO OC"
    print(f"\n ════════════════════════════════════════════════════════════")
    print(f"  [FASE {fase_actual}] {titulo}")
    print(f" ════════════════════════════════════════════════════════════")

    if fase_actual == 2:
        print("recolectando ordenes con error para segunda validacion...")

    datos_actuales = cargar_estado()
    errores_pendientes = [item for item in datos_actuales if item.get('etiqueta_final') == "NO MAL ERROR"]

    if not errores_pendientes:
        generar_reporte_final(datos_actuales, fase=fase_actual)
        return []

    print(f"re-validando {len(errores_pendientes)} ordenes...")
    
    # Reagrupacion para evitar mezclar codigos de facility en una misma consulta
    lotes_por_bodega = {}
    for err in errores_pendientes:
        b = err.get('bodega', 'n/a')
        if b not in lotes_por_bodega: lotes_por_bodega[b] = []
        lotes_por_bodega[b].append((err['orden_buscada'], err['sales_check'], b))

    for b_key, items_bodega in lotes_por_bodega.items():
        for i in range(0, len(items_bodega), 100):
            segmento = items_bodega[i:i + 100]
            resultados = consultar_lotes_ordenes(segmento, sesion)
            with bloqueo_fichero:
                with open(archivo_control, "a") as f_json:
                    for res in resultados:
                        f_json.write(json.dumps(res) + "\n")

    datos_finales = cargar_estado()
    generar_reporte_final(datos_finales, fase=fase_actual)
    return [item for item in datos_finales if item.get('etiqueta_final') == "NO MAL ERROR"]

def generar_reporte_final(datos_completos, fase=1):
    # ®Carlos Alfonso Ortega Molina®
    # cierre de fase actualiza reporte maestro bitacora de errores y marcas de finalizacion
    # deja evidencia de cada fase
    # permite comparar resultado entre memoria y reporte escrito
    if not datos_completos: return
    leyendas = {
        1: "Finalizacion Verificada y validada",
        2: "Segunda validacion de errores",
        3: "Tercera validacion tras reproceso GNX",
        4: "Cuarta y ultima validacion"
    }
    leyenda = leyendas.get(fase, f"Finalizacion Fase {fase}")

    if fase <= 2:
        print(f"\nGenerando reporte cierre fase {fase}")
    else:
        print(f"\n[REPORTE] --- Iniciando protocolo de cierre OC (Fase {fase}) ---")

    df = pd.DataFrame(datos_completos)
    columnas_orden = [
        'orden_buscada', 'sales_check', 'etiqueta_final', 'resultado',
        'estatus_id', 'id_interno', 'bodega', 'fecha_hora_proceso', 'fecha_validacion_original'
    ]
    df = df.reindex(columns=[c for c in columnas_orden if c in df.columns])
    df.to_excel(archivo_salida, index=False)

    lista_errores = [item for item in datos_completos if item.get('etiqueta_final') == "NO MAL ERROR"]
    conteo_errores_maestro = len(lista_errores)

    modo_error = "a"
    if fase == 1: modo_error = "w"

    with open(archivo_errores, modo_error) as f_err:
        f_err.write(f"\n--- reporte de fase {fase} ({datetime.now().strftime('%H:%M:%S')}) ---\n")
        for err in lista_errores:
            orden = err.get('orden_buscada', 'n/a')
            motivo = err.get('resultado', 'error')
            f_err.write(f"orden: {orden} | motivo: {motivo}\n")

    # Doble validacion de consistencia antes de sellar el cierre de fase
    with open(archivo_errores, "r") as f_ver:
        texto = f_ver.read()
        bloque = texto.split(f"--- reporte de fase {fase}")[-1]
        lineas_reporte = len([l for l in bloque.splitlines() if "orden:" in l])

    print(f"verificacion dual fase {fase}: maestro({conteo_errores_maestro}) vs reporte({lineas_reporte})")

    if conteo_errores_maestro == lineas_reporte:
        sh_final = f"\n{datetime.now().strftime('%d/%m/%Y %H:%M:%S')} {leyenda}\n"
        for arch in [archivo_texto_procesado, archivo_errores]:
            if os.path.exists(arch):
                with open(arch, "a") as f_app: f_app.write(sh_final)

    if conteo_errores_maestro == 0 or fase == 4:
        pass

def enviar_correo(total_gnx, total_memoria, total_hoy, validados, lista_pendientes):
    # Correo ejecutivo con semaforo diario para evitar duplicidad de notificaciones
    # incluye total de memoria y total validado hoy
    # esto separa historial de trabajo real del dia
    archivo_chk_correo = f".correo_enviado_oc_{fecha_hoy}"
    if os.path.exists(archivo_chk_correo):
        print(f" [i] El correo ya fue enviado previamente el dia de hoy ({archivo_chk_correo}).")
        return

    pendientes_fin = len(lista_pendientes)
    fecha_format = datetime.now().strftime("%d/%m/%Y")
    hora_format = datetime.now().strftime("%H:%M:%S")
    
    intentos_mail = 0
    max_intentos_mail = 3
    exito_mail = False

    # Reintento para fallas transitorias de cliente Outlook o perfil MAPI
    while intentos_mail < max_intentos_mail and not exito_mail:
        intentos_mail += 1
        print(f"\n [≫] Preparando notificacion OC por correo (Intento {intentos_mail}/{max_intentos_mail})...")
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
            mail.Subject = f"[REPORTE] Match Automatizado de Ordenes de Compra - {fecha_format}"
            mail.To = dests
            mail.SentOnBehalfOfName = "ortegac@sanborns.com.mx"

            color_pendientes = "color:red;" if pendientes_fin > 0 else "color:green;"
            bg_header = "#003366" # Azul institucional

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
        <h2>Reporte de Validacion de Ordenes de Compra</h2>
        <p>Fecha: {fecha_format} | Hora: {hora_format}</p>
    </div>
    <div class="content">
        <p>Se ha completado el ciclo de validación automatizada entre <b>GNX</b> y <b>WMS Oracle Cloud</b>.</p>
        
        <h3>Resumen Ejecutivo</h3>
        <table class="resumen">
          <tr><th>Concepto</th><th>Cantidad</th></tr>
          <tr><td>Total registros analizados (GNX)</td><td>{total_gnx:,}</td></tr>
          <tr><td>Ordenes recuperadas de memoria (Caché)</td><td>{total_memoria:,}</td></tr>
          <tr><td>Ordenes validadas hoy via API</td><td>{total_hoy:,}</td></tr>
          <tr><td>Validadas correctamente en WMS</td><td>{validados:,}</td></tr>
          <tr><td>Pendientes sin validacion final</td><td class="alert">{pendientes_fin}</td></tr>
        </table>"""

            if pendientes_fin > 0:
                html_body += f"""<br>
        <h3 style="color: #990000;">Detalle de Pendientes Definitivos</h3>
        <p>Estas ordenes no fueron encontradas en el WMS tras el protocolo de reproceso:</p>
        <table>
          <tr>
            <th style="width: 50%;">Numero de Orden</th>
            <th style="width: 50%;">Almacen / Bodega</th>
          </tr>"""
                for i, item in enumerate(lista_pendientes):
                    clase_fila = 'class="fila-par"' if i % 2 == 0 else ""
                    ord_p = item.get('orden_buscada', 'n/a')
                    alm_p = item.get('bodega', 'No detectado')
                    # Mapeo visual SIL para el correo
                    if alm_p == "110403": alm_p = "SIL"
                    html_body += f"<tr {clase_fila}><td>{ord_p}</td><td>{alm_p}</td></tr>"
                html_body += "</table>"
            else:
                html_body += """<br><div style="padding: 15px; background-color: #e6fffa; border: 1px solid #38b2ac; color: #234e52; border-radius: 5px;">
                    <b>¡Éxito total!</b> No se encontraron ordenes pendientes para el día de hoy.
                </div>"""

            html_body += f"""
        <div class="footer">
            <p>Este es un correo generado automaticamente por el Sistema de Match de Ordenes de Compra.<br>
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
    # Orquesta parseo de fuentes N I cache historica validaciones y eventual reproceso GNX
    # resumen de flujo
    # leer y combinar archivos
    # mapear bodega por facility
    # resolver cache
    # ejecutar lotes por bodega
    # cerrar fases y correo
    hay_archivos = os.path.exists(archivo_entrada_val) or os.path.exists(archivo_entrada_ins)
    if not hay_archivos: return

    mapa_ordenes_completas = {}
    duplicados_detectados = []
    
    # Merge de dos fuentes OC en un unico mapa de trabajo del dia
    for arch_oc in [archivo_entrada_val, archivo_entrada_ins]:
        # cada archivo aporta ordenes al mismo mapa maestro
        if os.path.exists(arch_oc):
            conteo_local = 0
            with open(arch_oc, "r") as f:
                for linea in f:
                    linea_limpia = linea.strip()
                    if linea_limpia:
                        conteo_local += 1
                        partes = linea_limpia.split('|')
                        if len(partes) >= 2:
                            # Columna 0 Facility Code 840 0 o 870 0 o 850 0 limpiar 0
                            fac_code = partes[0].strip()
                            if fac_code.endswith(".0"): fac_code = fac_code[:-2]
                            # Mapeo de bodegas 840 CIG 870 VAL 850 SIL 110403
                            bodega = "CIG" if fac_code == "840" else "VAL" if fac_code == "870" else "110403" if fac_code == "850" else "n/a"
                            
                            # Columna 1 Numero de Orden 17324936 0 limpiar 0
                            orden = partes[1].strip()
                            if orden.endswith(".0"): orden = orden[:-2]
                            
                            if orden:
                                if orden in mapa_ordenes_completas:
                                    duplicados_detectados.append(orden)
                                # Guardamos orden sales check y bodega
                                mapa_ordenes_completas[orden] = {'sales_check': orden, 'bodega': bodega}
            print(f" [i] Registros en {arch_oc}: {conteo_local}")
    
    if duplicados_detectados:
        set_dups = list(set(duplicados_detectados))
        print(f" [i] ALERTA: Se detectaron {len(duplicados_detectados)} registros duplicados en origen (seran colapsados).")
        print(f" [i] Detalle de algunas ordenes duplicadas: {set_dups[:20]}{'...' if len(set_dups)>20 else ''}")

    print("\n ══════════════════════════════════════════════════════════════════════")
    print("  [SISTEMA] INICIALIZANDO EJECUCION DE VERIFICACION OC")
    print(" ══════════════════════════════════════════════════════════════════════\n")

    # CONTROL FERREO Renombrado LOCAL inmediato de los archivos de entrada tras carga exitosa en memoria
    for arch_oc in [archivo_entrada_val, archivo_entrada_ins]:
        if os.path.exists(arch_oc):
            try:
                arch_rename = f"{arch_oc.replace('.txt', '')}_{fecha_hoy}.txt"
                if os.path.exists(arch_rename): os.remove(arch_rename)
                os.rename(arch_oc, arch_rename)
                print(f" [✓] Control Ferreo: {arch_oc} archivado de inmediato como {arch_rename}.")
            except Exception as e:
                print(f" [!] Aviso: No se pudo archivar localmente {arch_oc}: {e}")

    todas_keys = list(mapa_ordenes_completas.keys())
    total_gnx = len(todas_keys)
    res_previos = cargar_estado()
    hechos_hoy = {i['orden_buscada'] for i in res_previos}
    pendientes = [o for o in todas_keys if o not in hechos_hoy]

    cache = cargar_historial_ayer()
    para_api = []
    recuperadas_cache = []

    escritor = threading.Thread(target=hilo_guardado_continuo, daemon=True)
    escritor.start()

    # Separa trabajo nuevo de ordenes resueltas via memoria historica
    for orden in pendientes:
        # decide si la orden se resuelve por cache o por api
        k12 = "0" + orden if len(orden) == 11 else orden
        hit = cache.get(orden) or cache.get(k12)
        if hit:
            recuperadas_cache.append(orden)
            d = hit['data'].copy()
            d['orden_buscada'] = orden
            d['sales_check'] = mapa_ordenes_completas[orden]['sales_check']
            d['bodega'] = mapa_ordenes_completas[orden]['bodega']
            d['fecha_hora_proceso'] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            d['etiqueta_final'] = f"YA VALIDADA ANTES [{hit['fecha_validacion']}]"
            d['fecha_validacion_original'] = hit['fecha_validacion']
            cola_resultados.put(d)
        else:
            para_api.append(orden)

    if cache:
        print(f" [i] Resumen de Memoria: {len(cache)} cargadas -> {len(recuperadas_cache)} encontradas en el archivo de hoy.")

    if para_api:
        print(f"\n ════════════════════════════════════════════════════════════")
        print(f"  [FASE 1] VALIDACION BATCH API OC (Lotes de 100)")
        print(f"  Pendientes: {len(para_api)} | Total hoy: {len(todas_keys)}")
        print(f" ════════════════════════════════════════════════════════════")
        with requests.Session() as sesion:
            estrategia_reintento = Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
            adaptador = HTTPAdapter(pool_connections=5, pool_maxsize=5, max_retries=estrategia_reintento)
            sesion.mount("https://", adaptador)
            sesion.auth = (USUARIO, CONTRASENA)

            verificar_reentrada(res_previos, sesion)

            # Agrupar por bodega para Fase 1
            ordenes_por_bodega = {}
            for o in para_api:
                b = mapa_ordenes_completas[o]['bodega']
                if b not in ordenes_por_bodega: ordenes_por_bodega[b] = []
                ordenes_por_bodega[b].append((o, mapa_ordenes_completas[o]['sales_check'], b))

            # Construccion final de lotes por bodega para ejecucion paralela
            lotes_finales = []
            for b_key, items in ordenes_por_bodega.items():
                for i in range(0, len(items), 100):
                    lotes_finales.append(items[i:i + 100])

            procesadas = len(todas_keys) - len(para_api)

            with ThreadPoolExecutor(max_workers=5) as ex:
                # ejecucion paralela para acelerar validacion por lotes
                futuros = {ex.submit(consultar_lotes_ordenes, l, sesion): i for i, l in enumerate(lotes_finales)}
                idx = 1
                for f in as_completed(futuros):
                    res = f.result()
                    for r in res: cola_resultados.put(r)
                    procesadas += len(res)
                    print(f"lotes: [{idx}/{len(lotes_finales)}] | progreso: {procesadas}/{len(todas_keys)}...")
                    idx += 1
            print("finalizando escritura en disco...")

    cola_resultados.put(None)
    escritor.join()

    # total hoy Las ordenes que realmente fueron a la API
    total_hoy = len(para_api)

    # Cierre unificado para que todas las ramas calculen igual memoria trabajo remanente
    def despachar_correo():
        d_fin = cargar_estado()
        # Segregar memoria de hoy
        memoria_hoy = [i for i in d_fin if "YA VALIDADA ANTES" in i.get('etiqueta_final', '')]
        trabajo_hoy = [i for i in d_fin if "YA VALIDADA ANTES" not in i.get('etiqueta_final', '')]
        
        remanentes_hoy = [i for i in trabajo_hoy if i.get('etiqueta_final') == "NO MAL ERROR"]
        # validados hoy total hoy menos errores reales de hoy
        validados_hoy = len(trabajo_hoy) - len(remanentes_hoy)
        if validados_hoy < 0: validados_hoy = 0
        
        enviar_correo(total_gnx, len(memoria_hoy), len(trabajo_hoy), validados_hoy, remanentes_hoy)

    d_f1 = cargar_estado()
    generar_reporte_final(d_f1, fase=1)

    errores_f2 = []
    with requests.Session() as s:
        s.auth = (USUARIO, CONTRASENA)
        errores_f2 = procesar_lote_errores(s, fase_actual=2)

    if errores_f2:
        print(f"\n [!] ATENCION: Detectados {len(errores_f2)} errores tras Fase 2.")
        print(f" [!] Iniciando protocolo de Auto-Reparacion GNX...")
        chk_ssh_dia = f".chk_ssh_oc_{fecha_hoy}"
        if not os.path.exists(chk_ssh_dia):
            exito_gnx = subir_ftp_y_ejecutar_ssh(errores_f2)
            open(chk_ssh_dia, "w").close()
        else:
            exito_gnx = True

        if exito_gnx:
            errores_f3 = []
            with requests.Session() as s3:
                s3.auth = (USUARIO, CONTRASENA)
                errores_f3 = esperar_y_revalidar(s3, 3, errores_f2, 20)

            if errores_f3:
                print(f"\n ============================================================")
                print(f"  [ESTADO] --- INICIANDO FASE 4: REVISION FINAL DEFINITIVA OC ---")
                print(f" ============================================================")
                with requests.Session() as s4:
                    s4.auth = (USUARIO, CONTRASENA)
                    errores_f4 = esperar_y_revalidar(s4, 4, errores_f3, 30)

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
        despachar_correo()

if __name__ == "__main__":
    # Punto de entrada corre flujo normal y si aplica retoma procesos incompletos del mismo dia
    # el main decide entre ruta normal o ruta de recuperacion
    # si no hay archivos intenta cerrar pendientes guardados
    servicio_log = duplicador_salida(archivo_log)
    sys.stdout = servicio_log
    print(f"--- inicializando sistema automatizado OC [{fecha_hoy}] ---\n")
    if obtener_credenciales():
        hay_archivos = os.path.exists(archivo_entrada_val) or os.path.exists(archivo_entrada_ins)
        if not hay_archivos:
            if not fase_cero_ftp():
                print(" [!] ALERTA: No se pudo descargar archivo de hoy por FTP.")

        hay_archivos = os.path.exists(archivo_entrada_val) or os.path.exists(archivo_entrada_ins)
        if hay_archivos:
            # flujo normal hay archivo s de entrada procesar completo
            ejecutar_verificacion()
        else:
            # RECUPERACION Revisar si el proceso termino pero fallo el correo
            d_prev = cargar_estado()
            remanentes = [i for i in d_prev if i.get('etiqueta_final') == "NO MAL ERROR"]
            archivo_chk_correo = f".correo_enviado_oc_{fecha_hoy}"

            if len(d_prev) > 0 and not os.path.exists(archivo_chk_correo):
                # Caso 1 proceso previo completo sin confirmacion de notificacion
                print(f" [!] ALERTA DE RECUPERACION: Detectados datos procesados sin confirmacion de correo enviado.")
                total_gnx = len(d_prev)
                memoria_rec = len([i for i in d_prev if "YA VALIDADA ANTES" in i.get('etiqueta_final', '')])
                trabajo_rec = total_gnx - memoria_rec
                remanentes_rec = [i for i in d_prev if i.get('etiqueta_final') == "NO MAL ERROR"]
                validados_rec = trabajo_rec - len(remanentes_rec)
                if validados_rec < 0: validados_rec = 0
                print(f" [!] Re-intentando despacho de notificacion pendiente...")
                enviar_correo(total_gnx, memoria_rec, trabajo_rec, validados_rec, remanentes_rec)
                print(" [✓] Recuperacion de correo completada.")

            # sin archivo de entrada revisar si hay remanentes pendientes
            elif len(remanentes) > 0 and os.path.exists(f".chk_ssh_oc_{fecha_hoy}"):
                # ya se ejecuto el ssh solo falta revalidar fase 3 4
                print(f" [!] ALERTA DE RECUPERACION: {len(remanentes)} ordenes pendientes tras reproceso GNX.")
                total_gnx = len(d_prev)
                memoria_rec = len([i for i in d_prev if "YA VALIDADA ANTES" in i.get('etiqueta_final', '')])
                trabajo_rec = total_gnx - memoria_rec
                def despachar_correo():
                    d_fin = cargar_estado()
                    remanentes_fin = [i for i in d_fin if i.get('etiqueta_final') == "NO MAL ERROR"]
                    validados_fin = trabajo_rec - len(remanentes_fin)
                    if validados_fin < 0: validados_fin = 0
                    enviar_correo(total_gnx, memoria_rec, trabajo_rec, validados_fin, remanentes_fin)
                with requests.Session() as s_rec:
                    s_rec.auth = (USUARIO, CONTRASENA)
                    errores_rec3 = esperar_y_revalidar(s_rec, 3, remanentes, 20)
                    if errores_rec3:
                        esperar_y_revalidar(s_rec, 4, errores_rec3, 30)
                despachar_correo()
            elif len(remanentes) > 0 and not os.path.exists(f".chk_ssh_oc_{fecha_hoy}"):
                # nunca se llego al ssh intentar reproceso completo
                print(f" [!] ALERTA DE RECUPERACION: {len(remanentes)} ordenes pendientes sin reproceso GNX. Ejecutando...")
                exito_gnx = subir_ftp_y_ejecutar_ssh(remanentes)
                if exito_gnx:
                    open(f".chk_ssh_oc_{fecha_hoy}", "w").close()
                    total_gnx = len(d_prev)
                    memoria_rec = len([i for i in d_prev if "YA VALIDADA ANTES" in i.get('etiqueta_final', '')])
                    trabajo_rec = total_gnx - memoria_rec
                    def despachar_correo():
                        d_fin = cargar_estado()
                        remanentes_fin = [i for i in d_fin if i.get('etiqueta_final') == "NO MAL ERROR"]
                        validados_fin = trabajo_rec - len(remanentes_fin)
                        if validados_fin < 0: validados_fin = 0
                        enviar_correo(total_gnx, memoria_rec, trabajo_rec, validados_fin, remanentes_fin)
                    with requests.Session() as s_rec:
                        s_rec.auth = (USUARIO, CONTRASENA)
                        errores_rec3 = esperar_y_revalidar(s_rec, 3, remanentes, 20)
                        if errores_rec3:
                            esperar_y_revalidar(s_rec, 4, errores_rec3, 30)
                    despachar_correo()
            else:
                # Caso final no existe materia prima ni pendientes no hay acciones adicionales
                print("proceso del dia completado previamente. nada pendiente.")
    else:
        print("Error: Faltan credenciales validas en json (.acceso_wms o .acceso_gnx)")
    sys.stdout = servicio_log.consola
    servicio_log.close()
# ®Carlos Alfonso Ortega Molina®
