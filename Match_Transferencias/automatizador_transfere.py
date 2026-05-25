# ®Carlos Alfonso Ortega Molina®
# importacion de modulos del sistema y bibliotecas para la automatizacion de transferencias
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

# inicializar fecha actual formateada para nombres de archivo unico
fecha_hoy = datetime.now().strftime("%d%m%Y")

# url base del wms para ordenes
URL_BASE = "https://e6.wms.ocs.oraclecloud.com/sears2/wms/lgfapi/v10/entity/order_hdr/"
USUARIO = ""
CONTRASENA = ""

# configuracion del host gnx para ftp y ssh
GNX_HOST = "140.240.11.1"
GNX_USER = ""
GNX_PASS = ""

# flujo operativo recupera match diario valida en wms reprocesa en gnx y notifica cierre
# este script procesa transferencias del dia
# puede retomar desde archivo archivado del mismo dia
# cierra con reporte y correo final

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

archivo_log = f"log_transfere_{fecha_hoy}.txt"
archivo_entrada = "match_13_14.txt"
archivo_salida = f"reporte_transfere_{fecha_hoy}.xlsx"
archivo_control = f"control_transfere_{fecha_hoy}.json"
archivo_texto_procesado = f"transfere_procesado_{fecha_hoy}.txt"
archivo_errores = f"errores_transfere_{fecha_hoy}.txt"

# Artefactos diarios para control de ejecucion consolidacion y soporte operativo
bloqueo_fichero = threading.Lock()
cola_resultados = queue.Queue()

# Carga credenciales necesarias para acceder a API WMS y servidor GNX
def obtener_credenciales():
    # objetivo
    # activar conexiones de trabajo
    # salida
    # verdadero cuando hay acceso completo
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
    # La ejecucion queda bloqueada si falta cualquiera de las dos credenciales
    return wms_ok and gnx_ok

def fase_cero_ftp():
    # Obtiene el archivo match 13 14 del dia y valida fecha de origen
    # esta fase se usa cuando no existe insumo local
    # verifica que el archivo remoto sea del dia
    print(f"\n ══════════════════════════════════════════════════════════════════════")
    print(f"  [FASE 0] AUTO-ABASTECIMIENTO DE MATERIA PRIMA TRANSFERE (FTP)")
    print(f" ══════════════════════════════════════════════════════════════════════")
    print(f" [≫] Conectando a {GNX_HOST}...")
    descargados = 0
    try:
        with ftplib.FTP(GNX_HOST, encoding='latin-1') as ftp:
            ftp.login(GNX_USER, GNX_PASS)
            ruta = "/gnx_prod/manto/desa/trabajo/sears/ORACLE"
            ftp.cwd(ruta)
            
            if archivo_entrada in ftp.nlst():
                res = ftp.voidcmd(f"MDTM {archivo_entrada}")
                fecha_srv = res.split()[1][:8]
                if fecha_srv == datetime.now().strftime("%Y%m%d"):
                    with open(archivo_entrada, "wb") as f_loc:
                        ftp.retrbinary(f"RETR {archivo_entrada}", f_loc.write)
                    tamanio = os.path.getsize(archivo_entrada)
                    print(f" [✓] Archivo descargado: {archivo_entrada} ({tamanio} bytes)")
                    descargados += 1
                else:
                    print(f" [!] ALERTA: {archivo_entrada} no es de hoy ({fecha_srv}).")
            else:
                print(f" [!] Error: No se encontro {archivo_entrada} en el servidor.")
    except Exception as e:
        print(f" [X] Error en FTP: {e}")
    return descargados > 0

def subir_ftp_y_ejecutar_ssh(errores_pendientes):
    # Sube archivo i04 con incidencias y ejecuta reproceso remoto de transferencias
    # entrada
    # transferencias pendientes despues de fase dos
    # salida
    # verdadero cuando el script remoto termina
    print(f" ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print(f"  [GNX] INICIANDO PROTOCOLO REMOTO PARA {len(errores_pendientes)} INCIDENCIAS TRANSFERE")
    print(f" ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

    archivo_carga = "i04.txt"
    try:
        # Usamos newline n y encoding latin 1 para asegurar formato Unix y compatibilidad con GNX Informix
        with open(archivo_carga, "w", newline='\n', encoding='latin-1') as f_unl:
            for error in errores_pendientes:
                sales = error.get('sales_check', '').strip()
                if sales:
                    f_unl.write(f"{sales}\n")
        print(f" [✓] Archivo de carga generado: {archivo_carga} (Formato Unix/Latin-1)")
    except Exception as e:
        print(f"error generando archivo de carga: {e}")
        return False

    print(f" [≫] Subiendo {archivo_carga} a {GNX_HOST} via FTP Python...")
    try:
        with ftplib.FTP(GNX_HOST) as ftp:
            ftp.login(GNX_USER, GNX_PASS)
            # Forzamos modo binario para evitar alteraciones de fin de linea por el cliente FTP
            ftp.voidcmd('TYPE I')
            # Ruta donde GNX espera el archivo de entrada i04 txt Nombre remoto inmutable
            ftp.cwd("/respaldo_migracion/reportes_gnx")
            with open(archivo_carga, "rb") as f_sub:
                ftp.storbinary(f"STOR {archivo_carga}", f_sub)
            print(f" [✓] Transferencia ferrea de {archivo_carga} confirmada en /respaldo_migracion/reportes_gnx.")
        
        # Renombrado LOCAL inmediato tras subida exitosa para control de proceso
        try:
            archivo_i04_hist = f"i04_{fecha_hoy}.txt"
            if os.path.exists(archivo_i04_hist): os.remove(archivo_i04_hist)
            os.rename(archivo_carga, archivo_i04_hist)
            print(f" [i] Archivo local {archivo_carga} renombrado a {archivo_i04_hist}.")
        except Exception as e:
            print(f" [!] Aviso: No se pudo renombrar localmente {archivo_carga}: {e}")
            
    except Exception as e:
        print(f" [X] Error en transferencia FTP: {e}")
        return False
    except: return False

    print(f" [≫] Estableciendo conexion SSH (Paramiko) para reproceso TRANSFERE...")
    try:
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(GNX_HOST, username=GNX_USER, password=GNX_PASS)
        print(" [✓] Conexion establecida. Iniciando eco del servidor.")
        sh = ssh.invoke_shell(term='vt100', width=220, height=50)
        sh.settimeout(300)

        def leer_hasta_prompt(max_seg=10):
            # lectura continua de shell remoto para dejar traza en log
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
                # Pequena pausa antes de enviar para evitar saturar el buffer de entrada del servidor
                time.sleep(1)
                sh.send(cmd + "\n")
            leer_hasta_prompt(espera_seg)

        run_cmd("1", "\n>>> [AUTO] Seleccionando Ambiente 1 (SEARS)...", 5)
        run_cmd("1", ">>> [AUTO] Seleccionando Opcion 1 (SEARS)...", 5)
        # Ruta donde reside el ejecutable sh
        run_cmd("cd /gnx_prod/manto/desa/trabajo/sears/carlos_ortega", ">>> [AUTO] Seleccionando ruta de trabajo absoluta...", 5)

        print(">>> [AUTO] Ejecutando script de reproceso TRANSFERE... (esperando conclusion total)")
        sh.send("./automatizador_transferencias.sh\n")
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
    # Reusa validaciones exitosas previas para disminuir carga de consulta diaria
    # evita repetir validaciones que ya fueron exitosas
    ayer = datetime.now() - timedelta(days=1)
    str_yesterday = ayer.strftime("%d%m%Y")
    archivo_ayer = f"control_transfere_{str_yesterday}.json"
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
        print(f"memoria cargada: {len(cache_historica)} transferencias ya validadas previamente.")
    except Exception as e:
        print(f"aviso: error al leer historial {archivo_ayer}: {e}")
    return cache_historica

def cargar_estado():
    # Reconstruye el snapshot vigente del control diario por numero de transferencia
    # estado consolidado para soporte de recuperacion
    mapa_recuperado = {}
    if os.path.exists(archivo_control):
        try:
            with open(archivo_control, "r") as f:
                for num_linea, linea in enumerate(f, 1):
                    # Admite lectura robusta de historicos con lineas de ruido
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
    # Escritor dedicado evita colisiones de I O entre hilos de consulta
    # escribe json tecnico y txt operativo en cada resultado
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
    # Consulta transferencias en WMS y conserva bodega origen para trazabilidad de reporte
    # contrato comun
    # etiqueta ok verificado api para exitos
    # etiqueta no mal error para pendientes
    # lista datos entrada transferencia sales bodega orig
    if not lista_datos_entrada: return []
    
    ordenes_list = [o[0] for o in lista_datos_entrada]
    ordenes_str = ",".join(ordenes_list)
    resultados_finales_lote = []

    try:
        params = {
            'order_nbr__in': ordenes_str,
            'page_size': 100
        }
        respuesta = sesion.get(URL_BASE, params=params, timeout=25)
        ahora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        encontrados = {}
        if respuesta.status_code == 200:
            datos = respuesta.json()
            for info in datos.get('results', []):
                o_api = info.get('order_nbr')
                if o_api: encontrados[o_api] = info

        for orden, sales, bodega_orig in lista_datos_entrada:
            res_obj = {
                'orden_buscada': orden, 
                'sales_check': sales, 
                'bodega': bodega_orig, # Mantenemos la bodega original por defecto
                'fecha_hora_proceso': ahora
            }
            if orden in encontrados:
                info = encontrados[orden]
                res_obj.update({
                    'id_interno': info.get('id'),
                    'order_nbr_api': info.get('order_nbr'),
                    'estatus_id': info.get('status_id'),
                    # Si el WMS nos da una bodega la usamos si no dejamos la original
                    'bodega': info.get('facility_id', {}).get('key') if info.get('facility_id') else bodega_orig,
                    'resultado': "exito en wms",
                    'etiqueta_final': "OK VERIFICADO API",
                    'fecha_validacion_original': ahora
                })
            else:
                # Clasificacion uniforme de pendientes para facilitar reproceso posterior
                res_obj.update({'resultado': "no existe en wms" if respuesta.status_code == 200 else "error api", 'etiqueta_final': "NO MAL ERROR"})
            resultados_finales_lote.append(res_obj)
        return resultados_finales_lote
    except Exception as e:
        return [{'orden_buscada': o[0], 'sales_check': o[1], 'bodega': o[2], 'resultado': f"fallo lote: {str(e)}", 'etiqueta_final': "NO MAL ERROR", 'fecha_hora_proceso': datetime.now().strftime("%d/%m/%Y %H:%M:%S")} for o in lista_datos_entrada]

def verificar_reentrada(resultados_previos, sesion):
    # Verificacion rapida de estabilidad de estado sobre los ultimos resultados no cacheados
    # esta verificacion ayuda cuando el proceso se relanza en el mismo dia
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
            re_chequeo = consultar_lotes_ordenes(datos_entrada, sesion)[0]
            if re_chequeo['etiqueta_final'] != orden_res.get('etiqueta_final'):
                resultados_previos[idx] = re_chequeo
    return resultados_previos

def esperar_y_revalidar(sesion, fase_actual, errores_pendientes, minutos):
    # Revalidacion post reproceso con checkpoint para soportar reinicios sin repetir pausas
    # la funcion se usa para fase tres y fase cuatro
    # guarda checkpoint para no dormir dos veces
    archivo_checkpoint = f".fase{fase_actual}_iniciada_transfere_{fecha_hoy}"
    segundos = minutos * 60
    ahora = datetime.now()
    luego = ahora + timedelta(seconds=segundos)
    
    if fase_actual == 3:
        titulo = "VALIDACION POST-REPROCESO TRANSFERE"
        label_X = "Transferencias a monitorear"
    else:
        titulo = "REVISION FINAL DEFINITIVA TRANSFERE"
        label_X = "Transferencias en Revision Final"

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
        # Crear checkpoint antes de dormir para que si se interrumpe durante el sleep al reiniciar sepa que ya paso por aqui
        open(archivo_checkpoint, "w").close()
        time.sleep(segundos)
        
    print(f"------------------------------------------------------------")
    print(f"[API] Ejecutando consulta de verificacion Fase {fase_actual}...\n")

    # Se conserva la bodega de origen en el re chequeo para mantener contexto en reporte final
    lote_trabajo = [(err['orden_buscada'], err['sales_check'], err.get('bodega', 'n/a')) for err in errores_pendientes]
    tamanio_lote = 100
    for i in range(0, len(lote_trabajo), tamanio_lote):
        segmento = lote_trabajo[i:i + tamanio_lote]
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
    # Segunda validacion enfocada en remanentes de error
    # solo revisa transferencias en error
    # prepara terreno antes del reproceso gnx
    titulo = "REFUERZO DE ERRORES TRANSFERE" if fase_actual == 2 else "REVALIDACION Y MONITOREO TRANSFERE"
    print(f"\n ════════════════════════════════════════════════════════════")
    print(f"  [FASE {fase_actual}] {titulo}")
    print(f" ════════════════════════════════════════════════════════════")

    if fase_actual == 2:
        print("recolectando transferencias con error para segunda validacion...")

    datos_actuales = cargar_estado()
    errores_pendientes = [item for item in datos_actuales if item.get('etiqueta_final') == "NO MAL ERROR"]

    if not errores_pendientes:
        generar_reporte_final(datos_actuales, fase=fase_actual)
        return []

    print(f"re-validando {len(errores_pendientes)} transferencias...")
    # Solo reintenta transferencias etiquetadas en error dentro del estado consolidado
    lote_trabajo = [(err['orden_buscada'], err['sales_check'], err.get('bodega', 'n/a')) for err in errores_pendientes]
    
    tamanio_lote = 100
    for i in range(0, len(lote_trabajo), tamanio_lote):
        segmento = lote_trabajo[i:i + tamanio_lote]
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
    # Consolidacion de cierre Excel txt de errores sellos de finalizacion por fase
    # genera salidas para operacion diaria y soporte
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
        print(f"\n[REPORTE] --- Iniciando protocolo de cierre TRANSFERE (Fase {fase}) ---")

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

    # Contraste de consistencia para asegurar que el txt refleje el maestro consolidado
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
        # Renombrar archivo de entrada si todo fue exitoso o es la fase final
        if os.path.exists(archivo_entrada) and (conteo_errores_maestro == 0 or fase == 4):
            if len(datos_completos) > 0:
                try:
                    arch_rename = f"{archivo_entrada.replace('.txt', '')}_{fecha_hoy}.txt"
                    if os.path.exists(arch_rename): os.remove(arch_rename)
                    os.rename(archivo_entrada, arch_rename)
                    print(f"limpieza exitosa: {archivo_entrada} archivado.")
                except: pass
        
        # Asegurar que i04 txt se archive si existe usado en reproceso GNX
        archivo_i04 = "i04.txt"
        if os.path.exists(archivo_i04):
            try:
                arch_i04_rename = f"i04_{fecha_hoy}.txt"
                if os.path.exists(arch_i04_rename): os.remove(arch_i04_rename)
                os.rename(archivo_i04, arch_i04_rename)
                print(f"archivado: {archivo_i04} movido a {arch_i04_rename}")
            except: pass

def enviar_correo(total_gnx, total_hoy, validados, lista_pendientes):
    # Envio de resumen final con marca diaria para prevenir reenvios
    # envia resumen con detalle de pendientes
    # usa semaforo para evitar duplicados
    archivo_chk_correo = f".correo_enviado_transfere_{fecha_hoy}"
    if os.path.exists(archivo_chk_correo):
        print(f" [i] El correo ya fue enviado previamente el dia de hoy ({archivo_chk_correo}).")
        return

    pendientes_fin = len(lista_pendientes)
    fecha_format = datetime.now().strftime("%d/%m/%Y")
    hora_format = datetime.now().strftime("%H:%M:%S")
    
    intentos_mail = 0
    max_intentos_mail = 3
    exito_mail = False

    # Reintento de notificacion para fallos temporales de automatizacion de Outlook
    while intentos_mail < max_intentos_mail and not exito_mail:
        intentos_mail += 1
        print(f"\n [≫] Preparando notificacion TRANSFERE por correo (Intento {intentos_mail}/{max_intentos_mail})...")
        try:
            if not os.path.exists("destinatarios.txt"):
                print(" [!] Error: No existe destinatarios.txt")
                return

            dests = open("destinatarios.txt").read().strip().replace("\n", ";")
            if not dests:
                print(" [!] No se encontraron destinatarios en destinatarios.txt")
                return

            pythoncom.CoInitialize()
            # Pequena pausa para asegurar que el servidor COM este listo
            time.sleep(2)
            out = win32com.client.Dispatch("Outlook.Application")
            mail = out.CreateItem(0)
            mail.Subject = f"[REPORTE] Match Automatizado de Transferencias - {fecha_format}"
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
        <h2>Reporte de Validacion de Transferencias</h2>
        <p>Fecha: {fecha_format} | Hora: {hora_format}</p>
    </div>
    <div class="content">
        <p>Se ha completado el ciclo de validación automatizada entre <b>GNX</b> y <b>WMS Oracle Cloud</b>.</p>
        
        <h3>Resumen Ejecutivo</h3>
        <table class="resumen">
          <tr><th>Concepto</th><th>Cantidad</th></tr>
          <tr><td>Total registros analizados (GNX)</td><td>{total_gnx:,}</td></tr>
          <tr><td>Transferencias a validar hoy</td><td>{total_hoy:,}</td></tr>
          <tr><td>Validadas correctamente en WMS</td><td>{validados:,}</td></tr>
          <tr><td>Pendientes sin validacion final</td><td class="alert">{pendientes_fin}</td></tr>
        </table>"""

            if pendientes_fin > 0:
                html_body += f"""<br>
        <h3 style="color: #990000;">Detalle de Pendientes Definitivos</h3>
        <p>Estas transferencias no fueron encontradas en el WMS tras el protocolo de reproceso:</p>
        <table>
          <tr>
            <th style="width: 50%;">Numero de Transferencia</th>
            <th style="width: 50%;">Almacen Origen (GNX)</th>
          </tr>"""
                for i, item in enumerate(lista_pendientes):
                    clase_fila = 'class="fila-par"' if i % 2 == 0 else ""
                    ord_p = item.get('orden_buscada', 'n/a')
                    alm_p = item.get('bodega', 'No detectado')
                    html_body += f"<tr {clase_fila}><td>{ord_p}</td><td>{alm_p}</td></tr>"
                html_body += "</table>"
            else:
                html_body += """<br><div style="padding: 15px; background-color: #e6fffa; border: 1px solid #38b2ac; color: #234e52; border-radius: 5px;">
                    <b>¡Éxito total!</b> No se encontraron transferencias pendientes para el día de hoy.
                </div>"""

            html_body += f"""
        <div class="footer">
            <p>Este es un correo generado automaticamente por el Sistema de Match de Transferencias.<br>
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
            # CONTROL FERREO Generar marca de exito para evitar re envios y permitir recuperacion
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
    # Orquestador integral parsea insumo aplica cache valida por fases y cierra proceso
    # resumen de flujo
    # detectar archivo activo
    # cargar mapa maestro
    # separar cache y pendientes
    # validar en lotes
    # revalidar y reprocesar si aplica
    # enviar correo final
    # AUTORREPARACION Buscar archivo original o archivado de hoy
    archivo_actual = archivo_entrada
    if not os.path.exists(archivo_actual):
        arch_hist = f"{archivo_entrada.replace('.txt', '')}_{fecha_hoy}.txt"
        if os.path.exists(arch_hist):
            archivo_actual = arch_hist
        else:
            return

    mapa_ordenes_completas = {}
    duplicados_detectados = []
    
    # Parseo de insumo de transferencias y captura de duplicados para colapso por llave
    conteo_local = 0
    with open(archivo_actual, "r") as f:
        for linea in f:
            linea_limpia = linea.strip()
            if linea_limpia:
                conteo_local += 1
                partes = linea_limpia.split('|')
                if len(partes) >= 1:
                    # Columna 0 Numero de Transferencia 52405366 0 limpiar 0
                    orden = partes[0].strip()
                    if orden.endswith(".0"): orden = orden[:-2]
                    
                    # Columna 1 Bodega solo para el reporte
                    bodega = "n/a"
                    if len(partes) >= 2:
                        bodega = partes[1].strip()
                        if bodega.endswith(".0"): bodega = bodega[:-2]
                    
                    if orden:
                        if orden in mapa_ordenes_completas:
                            duplicados_detectados.append(orden)
                        # Guardamos orden como llave y bodega
                        mapa_ordenes_completas[orden] = {'sales_check': orden, 'bodega': bodega}
    
    print(f" [i] Registros en {archivo_actual}: {conteo_local}")

    # CONTROL FERREO Renombrado LOCAL inmediato del archivo de entrada tras carga exitosa en memoria
    if archivo_actual == archivo_entrada:
        try:
            arch_rename = f"{archivo_entrada.replace('.txt', '')}_{fecha_hoy}.txt"
            if os.path.exists(arch_rename): os.remove(arch_rename)
            os.rename(archivo_entrada, arch_rename)
            print(f" [✓] Control Ferreo: {archivo_entrada} archivado de inmediato como {arch_rename}.")
        except Exception as e:
            print(f" [!] Aviso: No se pudo archivar localmente el archivo de entrada: {e}")

    if duplicados_detectados:
        set_dups = list(set(duplicados_detectados))
        print(f" [i] ALERTA: Se detectaron {len(duplicados_detectados)} registros duplicados en origen (seran colapsados).")
        print(f" [i] Detalle de algunas transferencias duplicadas: {set_dups[:20]}{'...' if len(set_dups)>20 else ''}")

    print("\n ══════════════════════════════════════════════════════════════════════")
    print("  [SISTEMA] INICIALIZANDO EJECUCION DE VERIFICACION TRANSFERE")
    print(" ══════════════════════════════════════════════════════════════════════\n")

    todas_keys = list(mapa_ordenes_completas.keys())
    total_gnx = len(todas_keys)
    res_previos = cargar_estado()
    hechos_hoy = {i['orden_buscada'] for i in res_previos}
    
    # MEMORIA FERREA Filtrar pendientes comparando contra lo ya procesado hoy
    pendientes = [o for o in todas_keys if o not in hechos_hoy]

    cache = cargar_historial_ayer()
    para_api = []
    recuperadas_cache = []

    escritor = threading.Thread(target=hilo_guardado_continuo, daemon=True)
    escritor.start()

    # Se decide por transferencia si se toma cache historica o ruta de validacion API
    for orden in pendientes:
        # cada transferencia toma ruta cache o ruta api
        # Asegurar limpieza de orden para comparacion con cache historico
        o_limpia = str(orden)
        if len(o_limpia) == 11: o_limpia = "0" + o_limpia
        
        hit = cache.get(orden) or cache.get(o_limpia)
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
        print(f"  [FASE 1] VALIDACION BATCH API TRANSFERE (Lotes de 100)")
        print(f"  Pendientes: {len(para_api)} | Total hoy: {total_gnx}")
        print(f" ════════════════════════════════════════════════════════════")
        with requests.Session() as sesion:
            estrategia_reintento = Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
            adaptador = HTTPAdapter(pool_connections=5, pool_maxsize=5, max_retries=estrategia_reintento)
            sesion.mount("https://", adaptador)
            sesion.auth = (USUARIO, CONTRASENA)

            verificar_reentrada(res_previos, sesion)

            # Segmentacion en lotes de 100 para consultas paralelas estables
            lotes_finales = []
            for i in range(0, len(para_api), 100):
                seg = para_api[i:i + 100]
                lotes_finales.append([(o, mapa_ordenes_completas[o]['sales_check'], mapa_ordenes_completas[o].get('bodega', 'n/a')) for o in seg])

            procesadas = len(hechos_hoy) + len(pendientes) - len(para_api)

            with ThreadPoolExecutor(max_workers=5) as ex:
                # cinco workers para equilibrio entre velocidad y estabilidad
                futuros = {ex.submit(consultar_lotes_ordenes, l, sesion): i for i, l in enumerate(lotes_finales)}
                idx = 1
                for f in as_completed(futuros):
                    res = f.result()
                    for r in res: cola_resultados.put(r)
                    procesadas += len(res)
                    print(f"lotes: [{idx}/{len(lotes_finales)}] | transferencias: {procesadas}/{total_gnx}...")
                    
                    # INTELIGENCIA Pausa de seguridad cada 200 lotes 20 000 transferencias
                    if idx % 200 == 0 and idx < len(lotes_finales):
                        print(f"\n [i] ALCANZADO LIMITE DE 200 LOTES. Pausa de desconexion programada (60 seg)...")
                        time.sleep(60)
                        print(f" [✓] Reanudando consultas batch...\n")
                    
                    idx += 1
            print("finalizando escritura en disco...")

    cola_resultados.put(None)
    escritor.join()

    total_hoy = len(para_api)

    # Bloque comun de cierre para homogenizar calculos en flujo normal y recuperacion
    def despachar_correo():
        d_fin = cargar_estado()
        remanentes_hoy = [i for i in d_fin if i.get('etiqueta_final') == "NO MAL ERROR"]
        # En recuperacion para api puede ser 0 usamos el estado consolidado del dia
        total_hoy_correo = len([i for i in d_fin if "YA VALIDADA ANTES" not in i.get('etiqueta_final', '')])
        validados = total_hoy_correo - len(remanentes_hoy)
        if validados < 0: validados = 0
        enviar_correo(total_gnx, total_hoy_correo, validados, remanentes_hoy)

    d_f1 = cargar_estado()
    generar_reporte_final(d_f1, fase=1)

    errores_f2 = []
    with requests.Session() as s:
        s.auth = (USUARIO, CONTRASENA)
        errores_f2 = procesar_lote_errores(s, fase_actual=2)

    if errores_f2:
        print(f"\n [!] ATENCION: Detectados {len(errores_f2)} errores tras Fase 2.")
        
        # Archivo de control con fecha para asegurar control ferreo diario
        archivo_chk_ssh = f".chk_ssh_transfere_{fecha_hoy}"
        
        if not os.path.exists(archivo_chk_ssh):
            print(f" [!] Iniciando protocolo de Auto-Reparacion GNX via SSH (Archivo control: {archivo_chk_ssh} no encontrado)...")
            sys.stdout.flush()
            exito_gnx = subir_ftp_y_ejecutar_ssh(errores_f2)
            if exito_gnx:
                open(archivo_chk_ssh, "w").close()
                print(f" [✓] Protocolo SSH completado. Archivo de control {archivo_chk_ssh} generado.")
            else:
                print(f" [X] El protocolo SSH fallo. Se intentara continuar con la validacion normal.")
        else:
            print(f" [i] AVISO: El protocolo de Reparacion SSH ya se ejecuto exitosamente el dia de hoy ({archivo_chk_ssh}).")
            print(f" [i] Saltando ejecucion SSH para evitar saturacion. Procediendo a Fase 3.")
            exito_gnx = True

        if exito_gnx:
            errores_f3 = []
            with requests.Session() as s3:
                s3.auth = (USUARIO, CONTRASENA)
                errores_f3 = esperar_y_revalidar(s3, 3, errores_f2, 20)

            if errores_f3:
                print(f"\n ============================================================")
                print(f"  [ESTADO] --- INICIANDO FASE 4: REVISION FINAL DEFINITIVA TRANSFERE ---")
                print(f" ============================================================")
                with requests.Session() as s4:
                    s4.auth = (USUARIO, CONTRASENA)
                    errores_f4 = esperar_y_revalidar(s4, 4, errores_f3, 30)
                despachar_correo()
            else:
                despachar_correo()
        else:
            despachar_correo()
    else:
        despachar_correo()

if __name__ == "__main__":
    # Entrada con recuperacion avanzada prioriza completar procesos interrumpidos del dia
    # main con tres rutas
    # proceso ya cerrado
    # recuperacion de avance previo
    # inicio normal del dia
    servicio_log = duplicador_salida(archivo_log)
    sys.stdout = servicio_log
    print(f"--- inicializando sistema automatizado TRANSFERE [{fecha_hoy}] ---\n")
    
    if obtener_credenciales():
        # CONTROL DE ESTADO Revisar progreso previo antes de cualquier accion
        d_prev = cargar_estado()
        archivo_chk_ssh = f".chk_ssh_transfere_{fecha_hoy}"
        archivo_chk_correo = f".correo_enviado_transfere_{fecha_hoy}"
        
        # Caso A El proceso ya termino completamente hoy
        if os.path.exists(archivo_chk_correo):
            print(f" [i] El proceso del dia {fecha_hoy} ya finalizo exitosamente.")
            print(" [i] Nada pendiente por ejecutar.")
        
        # Caso B Hay datos procesados pero falta terminar etapas o el correo RECUPERACION PRIORITARIA
        elif len(d_prev) > 0:
            # Caso B existe progreso previo por lo que se prioriza reanudar antes de iniciar de cero
            print(f" [!] MODO RECUPERACION: Detectado progreso previo del dia de hoy ({len(d_prev)} registros).")
            
            # INTELIGENCIA Antes de cualquier protocolo de reparacion debemos terminar la Fase 1 si falta algo
            ejecutar_verificacion()
            
            # Recargar estado tras completar o intentar completar Fase 1
            d_prev = cargar_estado()
            remanentes = [i for i in d_prev if i.get('etiqueta_final') == "NO MAL ERROR"]
            total_gnx = len(d_prev)
            total_hoy = len([i for i in d_prev if "YA VALIDADA ANTES" not in i.get('etiqueta_final', '')])

            def despachar_correo_rec():
                d_fin = cargar_estado()
                remanentes_hoy = [i for i in d_fin if i.get('etiqueta_final') == "NO MAL ERROR"]
                validados = total_hoy - len(remanentes_hoy)
                if validados < 0: validados = 0
                print(f" [!] Ejecutando despacho de notificacion pendiente...")
                enviar_correo(total_gnx, total_hoy, validados, remanentes_hoy)

            # Sub caso 1 Ya se ejecuto SSH o no hubo errores solo falta correo o fases finales
            if os.path.exists(archivo_chk_ssh) or not remanentes:
                if remanentes:
                    print(f" [!] Reanudando validaciones post-SSH pendientes ({len(remanentes)} errores).")
                    with requests.Session() as s_rec:
                        s_rec.auth = (USUARIO, CONTRASENA)
                        # esperar y revalidar ya tiene checkpoints internos para saltar el sleep si es necesario
                        errores_rec3 = esperar_y_revalidar(s_rec, 3, remanentes, 20)
                        if errores_rec3:
                            esperar_y_revalidar(s_rec, 4, errores_rec3, 30)
                despachar_correo_rec()
            
            # Sub caso 2 Hubo errores pero NO se llego a ejecutar el protocolo SSH
            else:
                print(f" [!] El proceso previo detecto {len(remanentes)} errores pero el protocolo SSH no se inicio.")
                print(f" [!] Reiniciando flujo desde protocolo SSH para asegurar integridad.")
                with requests.Session() as s_full:
                    s_full.auth = (USUARIO, CONTRASENA)
                    # Re entramos a la logica principal desde donde se quedo
                    exito_gnx = subir_ftp_y_ejecutar_ssh(remanentes)
                    if exito_gnx:
                        open(archivo_chk_ssh, "w").close()
                        e3 = esperar_y_revalidar(s_full, 3, remanentes, 20)
                        if e3: esperar_y_revalidar(s_full, 4, e3, 30)
                    despachar_correo_rec()
            print(" [✓] Recuperacion completada.")

        # Caso C No hay progreso previo iniciar ejecucion normal
        else:
            # Caso C primer intento del dia se busca materia prima y se ejecuta flujo completo
            hay_archivo = os.path.exists(archivo_entrada)
            if not hay_archivo:
                if not fase_cero_ftp():
                    print(" [!] ALERTA: No se pudo descargar archivo de hoy por FTP.")
            
            if os.path.exists(archivo_entrada):
                ejecutar_verificacion()
            else:
                print(" [X] Error: No se encontro materia prima para procesar hoy.")
    else:
        print("Error: Faltan credenciales validas en json (.acceso_wms o .acceso_gnx)")
    sys.stdout = servicio_log.consola
    servicio_log.close()
# ®Carlos Alfonso Ortega Molina®
