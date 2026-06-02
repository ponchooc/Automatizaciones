#!/usr/bin/bash
# ============================================================================
# script: match_general.sh
# descripcion: Ejecuta el proceso completo de validación de SCNs y cruce con ASN
#              combinando la lógica de revisadorotro.sh y revisadorotro_2.sh
# autor: CARLOS ALFONSO ORTEGA MOLINA
# fecha: 2025-12-19
# ============================================================================

set -euo pipefail

# ============================================================================
# variables globales
# ============================================================================
readonly SCN_TXT="scn_mundo_wms.txt"
readonly STATUS_GNX="status_scn_gnx.txt"
readonly STATUS_LGA="status_scn_lga.txt"
readonly STATUS_WMS="status_scn_wms.txt"
readonly MATCH_FINAL="match_gnx_lgs_wms.txt"
readonly MATCH_TEMP="match_gnx_lgs_wms_temp.txt"
readonly LOG_FILE="proceso_scn.log"
readonly MATCH_FINAL_COMPLETO="match_gnx_lgs_wms_completo.txt"
readonly LOG_FILE_MATCH="proceso_match_asn.log"
readonly DIRECTORIO_TRABAJO="/respaldo_migracion/reportes_gnx/reportes/WMS"
readonly PREFIJO_DETORD="GPOSAN_VAL_OCS_DETORD"
readonly PREFIJO_ASNCLIE="GPOSAN_VAL_OCS_ASNCLIE"
readonly MAX_REINTENTOS=2
readonly ESPERA_SEGUNDOS=60

# Destinatarios de correo
readonly REMITENTE="Sistemas_Mercaderias@sears.com.mx"
readonly DESTINATARIOS="ortegac@sanborns.com.mx,juarezfg@hitss.com"

CSV=""
ASN=""
CSV_ORIG=""
ASN_ORIG=""

# ============================================================================
# funcion: obtener_patron_fecha
# descripcion: Genera el patron de fecha del dia actual (MES_DIA_ANIO)
# ============================================================================
obtener_patron_fecha() {
    local mes=$(date +%m | sed 's/^0//')
    local dia=$(date +%d)
    local anio=$(date +%Y)
    echo "${mes}_${dia}_${anio}"
}

# ============================================================================
# funcion: buscar_archivos_dia
# descripcion: Busca los archivos DETORD y ASNCLIE del dia actual
# retorna: 0 si encuentra ambos, 1 si no
# ============================================================================
buscar_archivos_dia() {
    local patron_fecha=$(obtener_patron_fecha)
    log_mensaje "Buscando archivos del dia con patron: ${patron_fecha}"
    
    # Buscar archivo DETORD
    CSV_ORIG=$(ls -1 "${DIRECTORIO_TRABAJO}/${PREFIJO_DETORD}_${patron_fecha}_"*.csv 2>/dev/null | head -1)
    # Buscar archivo ASNCLIE
    ASN_ORIG=$(ls -1 "${DIRECTORIO_TRABAJO}/${PREFIJO_ASNCLIE}_${patron_fecha}_"*.csv 2>/dev/null | head -1)
    
    if [[ -n "${CSV_ORIG}" && -f "${CSV_ORIG}" && -n "${ASN_ORIG}" && -f "${ASN_ORIG}" ]]; then
        log_mensaje "Archivo DETORD encontrado: ${CSV_ORIG}"
        log_mensaje "Archivo ASNCLIE encontrado: ${ASN_ORIG}"
        echo "Archivos encontrados:"
        echo "  DETORD: ${CSV_ORIG}"
        echo "  ASNCLIE: ${ASN_ORIG}"
        return 0
    else
        log_mensaje "No se encontraron archivos del dia"
        [[ -z "${CSV_ORIG}" ]] && log_mensaje "  - Falta archivo DETORD"
        [[ -z "${ASN_ORIG}" ]] && log_mensaje "  - Falta archivo ASNCLIE"
        return 1
    fi
}

# ============================================================================
# funcion: enviar_correo_sin_archivos
# descripcion: Envia correo notificando que no hay archivos disponibles
# ============================================================================
enviar_correo_sin_archivos() {
    local patron_fecha=$(obtener_patron_fecha)
    local asunto="AUN NO TENEMOS ARCHIVOS"
    local cuerpo_file="cuerpo_sin_archivos.txt"
    
    cat > "${cuerpo_file}" <<EOF
Aun no recibimos los archivos se re-intentara en una hora.

Patron de busqueda: ${patron_fecha}
Directorio: ${DIRECTORIO_TRABAJO}
Archivos esperados:
  - ${PREFIJO_DETORD}_${patron_fecha}_*.csv
  - ${PREFIJO_ASNCLIE}_${patron_fecha}_*.csv

Fecha/Hora: $(date '+%Y-%m-%d %H:%M:%S')
EOF
    
    cat "${cuerpo_file}" | mailx -s "${asunto}" -r "${REMITENTE}" ${DESTINATARIOS}
    local mail_exit=$?
    
    if [[ ${mail_exit} -eq 0 ]]; then
        log_mensaje "Correo de aviso enviado a: ${DESTINATARIOS}"
    else
        log_mensaje "Error al enviar correo de aviso"
    fi
    
    rm -f "${cuerpo_file}"
    return ${mail_exit}
}

# ============================================================================
# funcion: esperar_y_reintentar
# descripcion: Logica de busqueda con reintento
# ============================================================================
esperar_y_reintentar() {
    local intento=1
    
    while [[ ${intento} -le ${MAX_REINTENTOS} ]]; do
        log_mensaje "Intento ${intento} de ${MAX_REINTENTOS}"
        echo "Intento ${intento} de ${MAX_REINTENTOS}: Buscando archivos del dia..."
        
        if buscar_archivos_dia; then
            log_mensaje "Archivos encontrados en intento ${intento}"
            return 0
        fi
        
        if [[ ${intento} -lt ${MAX_REINTENTOS} ]]; then
            echo "Archivos no encontrados. Enviando notificacion y esperando 1 hora..."
            enviar_correo_sin_archivos
            log_mensaje "Esperando ${ESPERA_SEGUNDOS} segundos antes del siguiente intento"
            sleep ${ESPERA_SEGUNDOS}
        fi
        
        intento=$((intento + 1))
    done
    
    log_mensaje "Se agotaron los reintentos. Archivos no encontrados."
    echo "ERROR: No se encontraron los archivos despues de ${MAX_REINTENTOS} intentos."
    return 1
}

# ============================================================================
# funcion: log_mensaje
# ============================================================================
log_mensaje() {
    local mensaje="$1"
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] ${mensaje}" >> "${LOG_FILE}"
}

log_mensaje_match() {
    local mensaje="$1"
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] ${mensaje}" >> "${LOG_FILE_MATCH}"
}

# ============================================================================
# funcion: limpiar_archivos_temporales
# ============================================================================
limpiar_archivos_temporales() {
    log_mensaje "limpiando archivos temporales"
    rm -f "${SCN_TXT}" "${STATUS_GNX}" "${STATUS_LGA}" "${STATUS_WMS}" "${MATCH_TEMP}"
    log_mensaje "archivos temporales eliminados"
}

# ============================================================================
# funcion: validar_csv
# ============================================================================
validar_csv() {
    log_mensaje "validando existencia de archivo csv: ${CSV}"
    if [[ ! -f "${CSV}" ]]; then
        echo ""
        echo "error: archivo csv no encontrado: ${CSV}"
        echo ""
        echo "verifique que el archivo existe en el directorio actual"
        echo "directorio actual: $(pwd)"
        echo ""
        log_mensaje "error: archivo csv no encontrado"
        return 1
    fi
    if [[ ! -r "${CSV}" ]]; then
        echo ""
        echo "error: archivo csv no es legible: ${CSV}"
        echo ""
        echo "verifique los permisos del archivo"
        echo ""
        log_mensaje "error: archivo csv no es legible"
        return 1
    fi
    log_mensaje "archivo csv validado correctamente"
    return 0
}

validar_asn() {
    log_mensaje_match "validando existencia de archivo ASN: ${ASN}"
    if [[ ! -f "${ASN}" ]]; then
        echo ""
        echo "error: archivo ASN no encontrado: ${ASN}"
        echo ""
        echo "verifique que el archivo existe en el directorio actual"
        echo "directorio actual: $(pwd)"
        echo ""
        log_mensaje_match "error: archivo ASN no encontrado"
        return 1
    fi
    if [[ ! -r "${ASN}" ]]; then
        echo ""
        echo "error: archivo ASN no es legible: ${ASN}"
        echo ""
        echo "verifique los permisos del archivo"
        echo ""
        log_mensaje_match "error: archivo ASN no es legible"
        return 1
    fi
    log_mensaje_match "archivo ASN validado correctamente"
    return 0
}

# ============================================================================
# funcion: obtener_indices_csv (para CSV de órdenes)
# ============================================================================
obtener_indices_csv() {
    local header
    header=$(head -1 "${CSV}" | tr -d '\r\n')
    header=$(echo "${header}" | sed 's/"//g')
    IFS=',' read -ra cols <<< "${header}"
    for i in "${!cols[@]}"; do
        case "${cols[$i]}" in
            SCN) IDX_SCN=$((i+1)) ;;
            STATUS) IDX_STATUS=$((i+1)) ;;
            "ORDER RELAEASE") IDX_ORDER_RELEASE=$((i+1)) ;;
            "FECHA CREACION") IDX_FECHA_CREACION=$((i+1)) ;;
            "FECHA MODIFICACION") IDX_FECHA_MODIFICACION=$((i+1)) ;;
        esac
    done
    if [[ -z "${IDX_SCN:-}" || -z "${IDX_STATUS:-}" || -z "${IDX_ORDER_RELEASE:-}" || -z "${IDX_FECHA_CREACION:-}" || -z "${IDX_FECHA_MODIFICACION:-}" ]]; then
        echo "error: no se pudieron identificar todas las columnas requeridas en la cabecera del CSV"
        exit 1
    fi
}

# ============================================================================
# funcion: extraer_scns
# ============================================================================
extraer_scns() {
    log_mensaje "extrayendo scns de columna SCN del csv"
    awk -F',' -v idx_scn="${IDX_SCN}" 'NR>1 && $idx_scn != "" { gsub(/"/,"",$idx_scn); print "0" $idx_scn }' "${CSV}" | awk NF > "${SCN_TXT}"
    local total_scns=$(wc -l < "${SCN_TXT}")
    log_mensaje "scns extraidos: ${total_scns}"
    if [[ ${total_scns} -eq 0 ]]; then
        log_mensaje "error: no se extrajeron scns del csv"
        return 1
    fi
    return 0
}

# ============================================================================
# funcion: extraer_estados_wms
# ============================================================================
extraer_estados_wms() {
    log_mensaje "extrayendo datos wms del csv"
    awk -F',' -v idx_scn="${IDX_SCN}" -v idx_status="${IDX_STATUS}" -v idx_order="${IDX_ORDER_RELEASE}" -v idx_fcrea="${IDX_FECHA_CREACION}" -v idx_fmod="${IDX_FECHA_MODIFICACION}" '
        NR>1 && $idx_scn != "" && $idx_status != "" && $idx_order != "" && $idx_fcrea != "" && $idx_fmod != "" {
            gsub(/"/,"",$idx_scn);
            gsub(/"/,"",$idx_status);
            gsub(/"/,"",$idx_order);
            gsub(/"/,"",$idx_fcrea);
            gsub(/"/,"",$idx_fmod);
            print "0" $idx_scn "|" $idx_status "|" $idx_order "|" $idx_fcrea "|" $idx_fmod
        }
    ' "${CSV}" > "${STATUS_WMS}"
    local total_wms=$(wc -l < "${STATUS_WMS}")
    log_mensaje "registros wms extraidos: ${total_wms}"
    if [[ ${total_wms} -eq 0 ]]; then
        log_mensaje "error: no se extrajeron estados wms del csv"
        return 1
    fi
    return 0
}

# ============================================================================
# funcion: consultar_informix
# ============================================================================
consultar_informix() {
    log_mensaje "iniciando consultas en base de datos informix"
    dbaccess gen - > /dev/null 2>&1 <<'EOSQL'
CREATE TEMP TABLE scn_mundo_wms (scn_num CHAR(16)) WITH NO LOG;
LOAD FROM 'scn_mundo_wms.txt' INSERT INTO scn_mundo_wms;
UNLOAD TO 'status_scn_gnx.txt' DELIMITER '|'
SELECT
    a.num_scn,
    a.estado,
    a.cod_pto||''||a.num_edc
FROM edc_cab a, ordedc_cab b
WHERE a.cod_emp = 1
  AND a.num_scn IN (SELECT scn_num FROM scn_mundo_wms)
  AND b.cod_emp = a.cod_emp
  AND b.cod_pto = a.cod_pto
  AND b.num_edc = a.num_edc;
DROP TABLE IF EXISTS dts_scn_gnx;
CREATE TEMP TABLE dts_scn_gnx (
    scn_no CHAR(16),
    edo_1 CHAR(2),
    ord_release CHAR(20)
) WITH NO LOG;
LOAD FROM 'status_scn_gnx.txt' DELIMITER '|' INSERT INTO dts_scn_gnx;
UNLOAD TO 'status_scn_lga.txt' DELIMITER '|'
SELECT P.sales_check, D.st_etiqueta
FROM dblga@lga_prod_tcp:lgahventa P,
     dblga@lga_prod_tcp:lgadventa E,
     dblga@lga_prod_tcp:lgaetiqeta D
WHERE P.cod_empresa = E.cod_empresa
  AND P.cd_id = E.cd_id
  AND P.sales_check = E.sales_check
  AND P.cod_empresa = D.cod_empresa
  AND P.cd_id = D.cd_id
  AND E.no_etiqueta = D.no_etiqueta
  AND P.cod_empresa = 1
  AND P.cd_id = 870
  AND P.sales_check IN (SELECT scn_num FROM scn_mundo_wms);
DROP TABLE IF EXISTS dts_scn_lga;
CREATE TEMP TABLE dts_scn_lga (
    scn_no CHAR(16),
    edo_eti CHAR(2)
) WITH NO LOG;
LOAD FROM 'status_scn_lga.txt' DELIMITER '|' INSERT INTO dts_scn_lga;
DROP TABLE IF EXISTS dts_scn_wms;
CREATE TEMP TABLE dts_scn_wms (
    scn_no CHAR(16),
    estado CHAR(20),
    order_release CHAR(20),
    fecha_creacion CHAR(20),
    fecha_modificacion CHAR(20)
) WITH NO LOG;
LOAD FROM 'status_scn_wms.txt' DELIMITER '|' INSERT INTO dts_scn_wms;
UNLOAD TO 'match_gnx_lgs_wms_temp.txt' DELIMITER '|'
SELECT DISTINCT
    a.ord_release,
    a.scn_no,
    a.edo_1,
    b.edo_eti,
    c.estado,
    c.order_release,
    c.fecha_creacion,
    c.fecha_modificacion
FROM dts_scn_gnx a, dts_scn_lga b, dts_scn_wms c
WHERE b.scn_no = a.scn_no
  AND c.scn_no = b.scn_no
ORDER BY 1;
EOSQL
    local db_exit=$?
    if [[ ${db_exit} -ne 0 ]]; then
        log_mensaje "error: dbaccess fallo con codigo ${db_exit}"
        return 1
    fi
    log_mensaje "consultas informix completadas exitosamente"
    return 0
}

# ============================================================================
# funcion: convertir_estado_gnx
# ============================================================================
convertir_estado_gnx() {
    local codigo="$1"
    case "${codigo}" in
        I)  echo "I - Impreso" ;;
        P)  echo "P - Transito" ;;
        X)  echo "X - Cancelado" ;;
        C)  echo "C - CANCELACION EN PROCESO" ;;
        D)  echo "D - DEVOLUCION" ;;
        E)  echo "E - Entregado" ;;
        G)  echo "G - PREPAPRADO EN BACK ORDER)" ;;
        N)  echo "N - FICTICIO" ;;
        *)  echo "${codigo}" ;;
    esac
}

# ============================================================================
# funcion: convertir_estado_lga
# ============================================================================
convertir_estado_lga() {
    local codigo="$1"
    case "${codigo}" in
        0)  echo "0 - Disponible" ;;
        5)  echo "5 - Impreso" ;;
        6)  echo "6 - Encamion" ;;
        7)  echo "7 - Embarcada" ;;
        8)  echo "8 - Retenida" ;;
        9)  echo "9 - Reembarcado" ;;
        10) echo "10 - Cancelado" ;;
        *)  echo "${codigo}" ;;
    esac
}

# ============================================================================
# funcion: procesar_estados
# ============================================================================
procesar_estados() {
    log_mensaje "procesando y convirtiendo estados gnx y lga"
    if [[ ! -f "${MATCH_TEMP}" ]]; then
        log_mensaje "error: archivo temporal no existe"
        return 1
    fi
    while IFS='|' read -r ord_release scn edo_gnx edo_lga edo_wms order_release_wms fecha_creacion fecha_modificacion; do
        local edo_gnx_desc=$(convertir_estado_gnx "${edo_gnx}")
        local edo_lga_desc=$(convertir_estado_lga "${edo_lga}")
        echo "${ord_release}|${scn}|${edo_gnx_desc}|${edo_lga_desc}|${edo_wms}|${fecha_creacion}|${fecha_modificacion}"
    done < "${MATCH_TEMP}" > "${MATCH_TEMP}.processed"
    mv "${MATCH_TEMP}.processed" "${MATCH_TEMP}"
    log_mensaje "estados gnx y lga convertidos correctamente"
    return 0
}

# ============================================================================
# funcion: agregar_cabecera
# ============================================================================
agregar_cabecera() {
    log_mensaje "agregando cabecera al archivo final"
    if [[ ! -f "${MATCH_TEMP}" ]]; then
        log_mensaje "error: archivo temporal no existe"
        return 1
    fi
    echo "ORDEN RELEASE|SALES CHECK|ESTADO GNX|ESTADO LGA|ESTADO WMS|FECHA CREACION|FECHA MODIFICACION" > "${MATCH_FINAL}"
    cat "${MATCH_TEMP}" >> "${MATCH_FINAL}"
    log_mensaje "cabecera agregada correctamente"
    return 0
}

# ============================================================================
# funcion: validar_resultados
# ============================================================================
validar_resultados() {
    log_mensaje "validando archivo de resultados final"
    if [[ ! -f "${MATCH_FINAL}" ]]; then
        log_mensaje "error: archivo final no fue generado"
        return 1
    fi
    if [[ ! -s "${MATCH_FINAL}" ]]; then
        log_mensaje "error: archivo final esta vacio"
        registrar_debug
        return 1
    fi
    local total_registros=$(($(wc -l < "${MATCH_FINAL}") - 1))
    log_mensaje "archivo final generado: ${total_registros} registros"
    return 0
}

# ============================================================================
# funcion: registrar_debug
# ============================================================================
registrar_debug() {
    log_mensaje "registrando informacion de debug"
    local count_gnx=$(wc -l < "${STATUS_GNX}" 2>/dev/null || echo 0)
    local count_lga=$(wc -l < "${STATUS_LGA}" 2>/dev/null || echo 0)
    local count_wms=$(wc -l < "${STATUS_WMS}" 2>/dev/null || echo 0)
    log_mensaje "debug: status_scn_gnx.txt = ${count_gnx} registros"
    log_mensaje "debug: status_scn_lga.txt = ${count_lga} registros"
    log_mensaje "debug: status_scn_wms.txt = ${count_wms} registros"
}

# ============================================================================
# funcion: obtener_indices_csv_asn (para archivo ASN)
# ============================================================================
obtener_indices_csv_asn() {
    local archivo="$1"
    local header
    header=$(head -1 "${archivo}" | tr -d '\r\n')
    header=$(echo "${header}" | sed 's/"//g')
    IFS=',' read -ra cols <<< "${header}"
    for i in "${!cols[@]}"; do
        case "${cols[$i]}" in
            "FACTURA / SCN") IDX_SCN_ASN=$((i+1)) ;;
            "ORDER RELEASE / ASN") IDX_ORDER_RELEASE_ASN=$((i+1)) ;;
            STATUS) IDX_STATUS_ASN=$((i+1)) ;;
            "FECHA CREACION") IDX_FECHA_CREACION_ASN=$((i+1)) ;;
            "FECHA MODIFICACION") IDX_FECHA_MODIFICACION_ASN=$((i+1)) ;;
        esac
    done
    if [[ -z "${IDX_SCN_ASN:-}" || -z "${IDX_ORDER_RELEASE_ASN:-}" || -z "${IDX_STATUS_ASN:-}" || -z "${IDX_FECHA_CREACION_ASN:-}" || -z "${IDX_FECHA_MODIFICACION_ASN:-}" ]]; then
        echo "error: no se pudieron identificar todas las columnas requeridas en la cabecera del CSV ASN"
        exit 1
    fi
}

# ============================================================================
# funcion: cruce_asn
# ============================================================================
cruce_asn() {
    log_mensaje_match "Iniciando cruce entre ${MATCH_FINAL} y ${ASN}"
    obtener_indices_csv_asn "${ASN}"
    TMP_ASN="tmp_asn.$$"
    TMP_GNX="tmp_gnx.$$"
    awk -F',' -v idx_scn="${IDX_SCN_ASN}" -v idx_order="${IDX_ORDER_RELEASE_ASN}" -v idx_status="${IDX_STATUS_ASN}" -v idx_fcrea="${IDX_FECHA_CREACION_ASN}" -v idx_fmod="${IDX_FECHA_MODIFICACION_ASN}" '
        NR>1 {
            scn=$idx_scn; gsub(/"/,"",scn);
            order=$idx_order; gsub(/"/,"",order);
            status=$idx_status; gsub(/"/,"",status);
            fcrea=$idx_fcrea; gsub(/"/,"",fcrea);
            fmod=$idx_fmod; gsub(/"/,"",fmod);
            while(length(scn)<16) scn="0"scn;
            print scn "|" order "|" status "|" fcrea "|" fmod
        }
    ' "${ASN}" > "${TMP_ASN}"
    # El archivo MATCH_FINAL tiene: ORDEN RELEASE|SALES CHECK|ESTADO GNX|ESTADO LGA|ESTADO WMS|FECHA CREACION|FECHA MODIFICACION
    # El SCN está en la segunda columna
    awk -F'|' 'NR>1 {print $2 "|" $0}' "${MATCH_FINAL}" > "${TMP_GNX}"
    awk -F'|' '
        NR==FNR {
            scn_asn=$1
            asn_data[scn_asn]=$0
            next
        }
        {
            scn_gnx=$1
            # Extraer todos los campos de la orden (del archivo original)
            n = split($0, campos_gnx, "|")
            gnx_line = ""
            for(i=2; i<=n; i++) {
                gnx_line = gnx_line campos_gnx[i]
                if(i < n) gnx_line = gnx_line "|"
            }
            if(scn_gnx in asn_data) {
                split(asn_data[scn_gnx], campos_asn, "|")
                # campos_asn: scn|order|status|fcrea|fmod
                print gnx_line "|ASN|" campos_asn[1] "|" campos_asn[2] "|" campos_asn[3] "|" campos_asn[4] "|" campos_asn[5]
            } else {
                print gnx_line "|ASN|NO SE ENCONTRO ASN|NO SE ENCONTRO ASN|NO SE ENCONTRO ASN|NO SE ENCONTRO ASN|NO SE ENCONTRO ASN"
            }
        }
    ' "${TMP_ASN}" "${TMP_GNX}" > "${MATCH_FINAL_COMPLETO}.tmp"
    echo "ORDEN RELEASE|SALES CHECK|ESTADO GNX|ESTADO LGA|ESTADO WMS|FECHA CREACION|FECHA MODIFICACION|TIPO|SCN ASN|ORDER RELEASE ASN|STATUS ASN|FECHA CREACION ASN|FECHA MODIFICACION ASN" > "${MATCH_FINAL_COMPLETO}"
    cat "${MATCH_FINAL_COMPLETO}.tmp" >> "${MATCH_FINAL_COMPLETO}"
    rm -f "${TMP_ASN}" "${TMP_GNX}" "${MATCH_FINAL_COMPLETO}.tmp"
    log_mensaje_match "Cruce completado. Archivo generado: ${MATCH_FINAL_COMPLETO}"
    echo "Cruce completado. Archivo generado: ${MATCH_FINAL_COMPLETO}"
}

# ============================================================================
# funcion: main
# ============================================================================
main() {
    # Cambiar al directorio de trabajo
    cd "${DIRECTORIO_TRABAJO}" || {
        echo "ERROR: No se puede acceder al directorio ${DIRECTORIO_TRABAJO}"
        exit 1
    }
    
    log_mensaje "=========================================="
    log_mensaje "Iniciando proceso automatico de match"
    log_mensaje "Directorio de trabajo: ${DIRECTORIO_TRABAJO}"
    log_mensaje "=========================================="
    
    # Buscar archivos del dia con logica de reintento
    esperar_y_reintentar || exit 1
    
    # Pretratamiento de archivos de entrada
    # IMPORTANTE: Omitir las 2 primeras lineas (encabezado vacio)
    CSV="pretratado_ordenes.csv"
    ASN="pretratado_asn.csv"
    
    # Limpiar archivo de ordenes: omitir 2 primeras lineas, quitar CR, lineas en blanco, espacios extra
    awk 'NR>2 {gsub(/\r/,""); gsub(/[[:space:]]+$/,""); gsub(/^ +| +$/,"",$0); if(NF && $0!="") print}' "$CSV_ORIG" | tr -d '\r' > "$CSV"
    
    # Limpiar archivo ASN: omitir 2 primeras lineas, quitar CR, lineas en blanco, espacios extra
    awk 'NR>2 {gsub(/\r/,""); gsub(/[[:space:]]+$/,""); gsub(/^ +| +$/,"",$0); if(NF && $0!="") print}' "$ASN_ORIG" | tr -d '\r' > "$ASN"
    
    log_mensaje "Archivo DETORD preprocesado: $(wc -l < "$CSV") lineas"
    log_mensaje "Archivo ASNCLIE preprocesado: $(wc -l < "$ASN") lineas"
    log_mensaje "=========================================="
    limpiar_archivos_temporales
    validar_csv || exit 1
    validar_asn || exit 1
    obtener_indices_csv
    extraer_scns || exit 1
    extraer_estados_wms || exit 1
    consultar_informix || exit 1
    procesar_estados || exit 1
    agregar_cabecera || exit 1
    validar_resultados || exit 1
    limpiar_archivos_temporales
    cruce_asn || exit 1

    # Limpieza final del archivo de salida para formato Windows (CRLF) y sin líneas en blanco
    if [[ -f "${MATCH_FINAL_COMPLETO}" ]]; then
        # Eliminar líneas en blanco
        grep -v '^$' "${MATCH_FINAL_COMPLETO}" > "${MATCH_FINAL_COMPLETO}.tmp" && mv "${MATCH_FINAL_COMPLETO}.tmp" "${MATCH_FINAL_COMPLETO}"
        # Convertir a CRLF para Windows
        awk '{ sub(/\r$/,""); print $0 "\r" }' "${MATCH_FINAL_COMPLETO}" > "${MATCH_FINAL_COMPLETO}.tmp" && mv "${MATCH_FINAL_COMPLETO}.tmp" "${MATCH_FINAL_COMPLETO}"
        log_mensaje "archivo convertido a formato CRLF"
    fi

    log_mensaje "=========================================="
    log_mensaje "proceso completado exitosamente"
    log_mensaje "=========================================="

    # =================== ENVÍO DE CORREO CON ADJUNTO ZIP ===================
    # Los destinatarios estan definidos en las variables globales al inicio del script


    # Comprimir el archivo generado
    ZIP_FILE="${MATCH_FINAL_COMPLETO}.zip"
    if [ -f "$MATCH_FINAL_COMPLETO" ] && [ -s "$MATCH_FINAL_COMPLETO" ]; then
        zip -j "$ZIP_FILE" "$MATCH_FINAL_COMPLETO"
        ZIP_EXIT_CODE=$?
        if [ $ZIP_EXIT_CODE -ne 0 ]; then
            log_mensaje "ERROR: Fallo al comprimir el archivo $MATCH_FINAL_COMPLETO (zip exit code: $ZIP_EXIT_CODE)"
            return 1
        fi
    else
        log_mensaje "ERROR: El archivo $MATCH_FINAL_COMPLETO no existe o está vacío, no se puede comprimir ni enviar."
        return 1
    fi

    # Asunto y cuerpo del correo (mejorado)
    ASUNTO='Aqui se envia el match correspondiente'
    CUERPO_FILE="cuerpo_mail.txt"
    cat > "${CUERPO_FILE}" <<EOF
Buenos dias,

Se anexa el archivo que contiene el match de Genesix, LGA y WMS.

Saludos.
EOF
    if [ -f match_gnx_lgs_wms_completo.txt.zip ] && [ -s match_gnx_lgs_wms_completo.txt.zip ]; then
        ( cat "${CUERPO_FILE}"; uuencode match_gnx_lgs_wms_completo.txt.zip match_gnx_lgs_wms_completo.txt.zip ) | mailx -s "${ASUNTO}" -r "${REMITENTE}" ${DESTINATARIOS}
        MAILX_EXIT_CODE=$?
        if [ $MAILX_EXIT_CODE -eq 0 ]; then
            log_mensaje "Correo enviado correctamente a: ${DESTINATARIOS} con adjunto: match_gnx_lgs_wms_completo.txt.zip"
            echo "[\$(date '+%Y-%m-%d %H:%M:%S')] Correo enviado correctamente a: ${DESTINATARIOS} con adjunto: match_gnx_lgs_wms_completo.txt.zip"
        else
            log_mensaje "Error al enviar correo a: ${DESTINATARIOS}"
        fi
        rm -f match_gnx_lgs_wms_completo.txt.zip
    fi
    rm -f "${CUERPO_FILE}"

    return 0
}

main
exit $?
