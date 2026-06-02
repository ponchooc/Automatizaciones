#!/usr/bin/bash
################################################################################
# Script: procesar_rdd.sh
# Descripción: Procesamiento automatizado de Registros de Devolución (RDD)
# Sistema: AIX 7.2 | Bash 5.2.15 | Informix 12.10.FC13
# Base de datos: gen
# Autor: Generado para procesamiento RDD Sears
# Fecha: 2025-11-26
################################################################################

set -o pipefail

# Constantes
readonly CODEMP=1
readonly RUTA_4GL="/gnx_prod/manto/genesix/trabajo/sears/.ace/i68"
readonly PROG_4GL="/gnx_prod/sears/sql/ora_integra/ora_vtas_fixed_49_23_51_definitivo_forza.4go"
readonly REPORTE_DIR="${RUTA_4GL}/reportes/gnx"
readonly SERVIDOR_REMOTO="10.128.50.3"
readonly USUARIO_REMOTO="usr-ssh"
readonly TMP_DIR="/tmp/rdd_process_$$"

# Variables globales
AMBIENTE=""
RUTA_REMOTA=""

################################################################################
# Funciones de utilidad
################################################################################

limpiar_tmp() {
    [[ -d "${TMP_DIR}" ]] && rm -rf "${TMP_DIR}"
}

trap limpiar_tmp EXIT

inicializar() {
    mkdir -p "${TMP_DIR}" || {
        echo "ERROR: No se pudo crear directorio temporal ${TMP_DIR}"
        exit 1
    }
}

mostrar_encabezado() {
    clear
    echo "================================================================================"
    echo "        PROCESAMIENTO AUTOMATIZADO DE REGISTROS DE DEVOLUCIÓN (RDD)"
    echo "================================================================================"
    echo ""
}

pausar() {
    echo ""
    read -p "Presione ENTER para continuar..." dummy
}

confirmar() {
    local mensaje="$1"
    local respuesta
    while true; do
        read -p "${mensaje} (S/N): " respuesta
        case "${respuesta}" in
            [Ss]|[Ss][Ii]) return 0 ;;
            [Nn]|[Nn][Oo]) return 1 ;;
            *) echo "Respuesta inválida. Use S o N." ;;
        esac
    done
}

validar_decimal() {
    local valor="$1"
    local enteros=$2
    local decimales=$3
    
    # Validar que sea numérico
    if ! [[ "${valor}" =~ ^[0-9]+$ ]]; then
	 valor=$(awk -F. '{print $1}' <<< "$valor")
	 return 1
    fi
    
    # Validar longitud máxima
    local longitud=${#valor}
    if [[ ${longitud} -gt ${enteros} ]]; then
        return 1
    fi
    
    return 0
}

validar_char() {
    local valor="$1"
    local longitud_max=$2
    
    local longitud=${#valor}
    if [[ ${longitud} -gt ${longitud_max} ]]; then
        return 1
    fi
    
    return 0
}

################################################################################
# Funciones de entrada de datos
################################################################################

solicitar_ambiente() {
    local opcion
    echo "Seleccione el ambiente de destino:"
    echo "  1) UAT"
    echo "  2) PRODUCCIÓN"
    echo ""
    while true; do
        read -p "Opción (1-2): " opcion
        case "${opcion}" in
            1)
                AMBIENTE="UAT"
                RUTA_REMOTA="/E/software/desarrollo/Integraciones/Lectura/int49"
                return 0
                ;;
            2)
                AMBIENTE="PRODUCCION"
                RUTA_REMOTA="/E/software/integraciones/Lectura/int49"
                return 0
                ;;
            *)
                echo "ERROR: Opción inválida. Seleccione 1 o 2."
                ;;
        esac
    done
}

solicitar_lote() {
    confirmar "¿Desea procesar múltiples RDD en lote?"
}

capturar_rdds() {
    local es_lote=$1
    local cod_pto num_rdd
    
    if [[ ${es_lote} -eq 0 ]]; then
        echo ""
        echo "Ingrese los RDD a procesar (formato: COD_PTO NUM_RDD)"
        echo "Ejemplo: 114 3013"
        echo "COD_PTO: decimal(4,0) - máximo 4 dígitos"
        echo "NUM_RDD: decimal(8,0) - máximo 8 dígitos"
        echo "Escriba 'FIN' cuando termine"
        echo ""
        
        local contador=0
        while true; do
            read -p "RDD #$((contador+1)): " cod_pto num_rdd
            
            [[ "${cod_pto}" == "FIN" || "${cod_pto}" == "fin" ]] && break
            
            # Validar cod_pto: decimal(4,0)
            if ! validar_decimal "${cod_pto}" 4 0; then
                echo "ERROR: COD_PTO debe ser numérico con máximo 4 dígitos (decimal(4,0))"
                continue
            fi
            
            # Validar num_rdd: decimal(8,0)
            if ! validar_decimal "${num_rdd}" 8 0; then
                echo "ERROR: NUM_RDD debe ser numérico con máximo 8 dígitos (decimal(8,0))"
                continue
            fi
            
            echo "${cod_pto}|${num_rdd}" >> "${TMP_DIR}/rdds_lista.txt"
            ((contador++))
        done
        
        if [[ ${contador} -eq 0 ]]; then
            echo "ERROR: No se capturó ningún RDD"
            return 1
        fi
        
        echo "Total de RDD capturados: ${contador}"
    else
        echo ""
        echo "COD_PTO: decimal(4,0) - máximo 4 dígitos"
        echo "NUM_RDD: decimal(8,0) - máximo 8 dígitos"
        echo ""
        
        while true; do
            read -p "Ingrese COD_PTO (almacén): " cod_pto
            
            if ! validar_decimal "${cod_pto}" 4 0; then
                echo "ERROR: COD_PTO debe ser numérico con máximo 4 dígitos (decimal(4,0))"
                continue
            fi
            break
        done
        
        while true; do
            read -p "Ingrese NUM_RDD (número de devolución): " num_rdd
            
            if ! validar_decimal "${num_rdd}" 8 0; then
                echo "ERROR: NUM_RDD debe ser numérico con máximo 8 dígitos (decimal(8,0))"
                continue
            fi
            break
        done
        
        echo "${cod_pto}|${num_rdd}" > "${TMP_DIR}/rdds_lista.txt"
    fi
    
    return 0
}

################################################################################
# Funciones de procesamiento de base de datos
################################################################################

validar_rdd_cab() {
    local cod_pto=$1
    local num_rdd=$2
    local archivo_salida="${TMP_DIR}/rdd_cab_${cod_pto}_${num_rdd}.unl"
    
    echo "  → Validando existencia en tabla rdd_cab..."
    
    dbaccess gen <<EOF > /dev/null 2>&1
UNLOAD TO '${archivo_salida}' DELIMITER '|'
SELECT cod_emp, cod_pto, num_rdd, num_f33, pto_alm, num_scn, scn_nvo 
FROM rdd_cab
WHERE cod_emp = ${CODEMP} AND cod_pto = ${cod_pto} AND num_rdd = ${num_rdd};
EOF
    
    if [[ $? -ne 0 ]]; then
        echo "  ✗ ERROR: Fallo al consultar base de datos"
        return 1
    fi
    
    if [[ ! -s "${archivo_salida}" ]]; then
        echo "  ✗ ERROR: No existe registro en rdd_cab para cod_pto=${cod_pto}, num_rdd=${num_rdd}"
        return 1
    fi
    
    # Validar que el num_scn extraído sea char(16)
    local num_scn=$(awk -F'|' '{print $6}' "${archivo_salida}" | head -1)
    local scn_nvo=$(awk -F'|' '{print $7}' "${archivo_salida}" | head -1)
    
    if ! validar_char "${num_scn}" 16; then
        echo "  ✗ ERROR: num_scn excede longitud máxima de 16 caracteres"
        return 1
    fi
    
    if [[ -n "${scn_nvo}" ]] && ! validar_char "${scn_nvo}" 16; then
        echo "  ✗ ERROR: scn_nvo excede longitud máxima de 16 caracteres"
        return 1
    fi
    
    echo "  ✓ Registro encontrado en rdd_cab"
    echo "    num_scn: ${num_scn}"
    [[ -n "${scn_nvo}" ]] && echo "    scn_nvo: ${scn_nvo}"
    
    return 0
}

verificar_ora_integra_envio() {
    local cod_pto=$1
    local num_rdd=$2
    local archivo_salida="${TMP_DIR}/ora_integra_${cod_pto}_${num_rdd}.unl"
    
    echo "  → Verificando relación con ora_integra_envio..."
    
    dbaccess gen <<EOF > /dev/null 2>&1
UNLOAD TO '${archivo_salida}' DELIMITER '|'
SELECT o.num_scn 
FROM rdd_cab r, ora_integra_envio o
WHERE r.cod_pto = ${cod_pto} 
  AND r.num_rdd = ${num_rdd}
  AND (o.num_scn = r.num_scn OR o.num_scn = r.scn_nvo)
  AND o.tipo = 'RD';
EOF
    
    if [[ $? -ne 0 ]]; then
        echo "  ✗ ERROR: Fallo al consultar ora_integra_envio"
        return 2
    fi
    
    if [[ ! -s "${archivo_salida}" ]]; then
        echo "  ! CASO 2: No existe relación en ora_integra_envio"
        return 1
    fi
    
    local num_scn=$(cat "${archivo_salida}" | head -1)
    echo "  ✓ Relación encontrada en ora_integra_envio"
    echo "    num_scn: ${num_scn}"
    
    return 0
}

insertar_ora_integra_envio() {
    local cod_pto=$1
    local num_rdd=$2
    local archivo_entrada="${TMP_DIR}/rdd_cab_${cod_pto}_${num_rdd}.unl"
    
    echo "  → Insertando registro en ora_integra_envio (CASO 2)..."
    
    # Extraer datos del archivo de rdd_cab
    local pto_emi=$(awk -F'|' '{print $2}' "${archivo_entrada}" | head -1)
    local pto_alm=$(awk -F'|' '{print $5}' "${archivo_entrada}" | head -1)
    local num_scn=$(awk -F'|' '{print $6}' "${archivo_entrada}" | head -1)
    
    # Validar pto_emi: decimal(4,0)
    if ! validar_decimal "${pto_emi}" 3 0; then
        echo "  ✗ ERROR: pto_emi inválido (debe ser decimal(4,0))"
        return 1
    fi
    
    # Validar pto_alm: decimal(4,0)
    if ! validar_decimal "${pto_alm}" 3 0; then
        echo "  ✗ ERROR: pto_alm inválido (debe ser decimal(4,0))"
        return 1
    fi
    
    # Validar num_scn: char(16)
    if ! validar_char "${num_scn}" 16; then
        echo "  ✗ ERROR: num_scn excede longitud máxima de 16 caracteres"
        return 1
    fi
    
    echo "    Datos a insertar:"
    echo "      pto_emi: ${pto_emi}"
    echo "      cod_pto: ${pto_alm}"
    echo "      num_scn: ${num_scn}"
    
    dbaccess gen <<EOF > /dev/null 2>&1
INSERT INTO ora_integra_envio 
(pto_emi, num_ped, cod_pto, pto_tra, tipo, num_scn, cod_fam2, cod_pro, status, accion, fec_alt, pro_nam)
VALUES (${pto_emi}, 0, ${pto_alm}, 0, 'RD', '${num_scn}', NULL, NULL, NULL, 'CREATE', TODAY, 'genesix');
EOF
    
    if [[ $? -ne 0 ]]; then
        echo "  ✗ ERROR: Fallo al insertar en ora_integra_envio"
        return 1
    fi
    
    echo "  ✓ Registro insertado exitosamente en ora_integra_envio"
    
    # Crear archivo de salida para continuar con el flujo
    echo "${num_scn}" > "${TMP_DIR}/ora_integra_${cod_pto}_${num_rdd}.unl"
    return 0
}

insertar_i15() {
    local cod_pto=$1
    local num_rdd=$2
    
    echo "  → Insertando datos en tabla i15..."
    
    dbaccess gen <<EOF > /dev/null 2>&1
DELETE FROM i15 WHERE 1=1;

INSERT INTO i15
SELECT o.num_scn 
FROM rdd_cab r, ora_integra_envio o
WHERE r.cod_pto = ${cod_pto} 
  AND r.num_rdd = ${num_rdd}
  AND (o.num_scn = r.num_scn OR o.num_scn = r.scn_nvo)
  AND o.tipo = 'RD';
EOF
    
    if [[ $? -ne 0 ]]; then
        echo "  ✗ ERROR: Fallo al insertar en tabla i15"
        return 1
    fi
    
    # Validar que se insertó el registro
    local archivo_validacion="${TMP_DIR}/i15_validacion_${cod_pto}_${num_rdd}.unl"
    
    dbaccess gen <<EOF > /dev/null 2>&1
UNLOAD TO '${archivo_validacion}' DELIMITER '|'
SELECT num_scn FROM i15;
EOF
    
    if [[ ! -s "${archivo_validacion}" ]]; then
        echo "  ✗ ERROR: No se insertaron datos en i15"
        return 1
    fi
    
    local num_scn_i15=$(cat "${archivo_validacion}" | head -1)
    echo "  ✓ Datos insertados en i15"
    echo "    num_scn: ${num_scn_i15}"
    
    return 0
}

################################################################################
# Funciones de procesamiento 4GL y archivos
################################################################################

ejecutar_4gl() {
    echo ""
    echo "───────────────────────────────────────────────────────────────────────────────"
    echo "PASO: Ejecución del proceso 4GL"
    echo "───────────────────────────────────────────────────────────────────────────────"
    echo ""
    echo "Se ejecutará el programa Informix 4GL para generar los archivos CSV"
    echo "Ruta: ${RUTA_4GL}"
    echo "Programa: ${PROG_4GL}"
    echo ""
    
    if ! confirmar "¿Desea proceder con la ejecución del proceso 4GL?"; then
        echo "Proceso cancelado por el usuario"
        return 1
    fi
    
    echo ""
    echo "Ejecutando proceso 4GL..."
    
    cd "${RUTA_4GL}" || {
        echo "✗ ERROR: No se pudo acceder a ${RUTA_4GL}"
        return 1
    }
    
    fglgo "${PROG_4GL}" 2>&1
    local resultado=$?
    
    if [[ ${resultado} -ne 0 ]]; then
        echo "✗ ERROR: Fallo la ejecución del proceso 4GL (codigo: ${resultado})"
        return 1
    fi
    
    echo "✓ Proceso 4GL ejecutado correctamente"
    sleep 2
    return 0
}

buscar_archivos_generados() {
    local cod_pto=$1
    local num_rdd=$2
    local archivo_scn="${TMP_DIR}/ora_integra_${cod_pto}_${num_rdd}.unl"
    
    if [[ ! -f "${archivo_scn}" ]]; then
        echo "  ✗ ERROR: No se encontró archivo con num_scn"
        return 1
    fi
    
    local num_scn=$(cat "${archivo_scn}" | head -1 | tr -d ' \n\r')
    
    # Validar num_scn: char(16)
    if ! validar_char "${num_scn}" 16; then
        echo "  ✗ ERROR: num_scn excede longitud máxima de 16 caracteres"
        return 1
    fi
    
    echo ""
    echo "  → Buscando archivos generados para num_scn: ${num_scn}..."
    
    cd "${REPORTE_DIR}" || {
        echo "  ✗ ERROR: No se pudo acceder a ${REPORTE_DIR}"
        return 1
    }
    
    local archivos_encontrados=$(grep -l "${num_scn}" *.csv 2>/dev/null)
    
    if [[ -z "${archivos_encontrados}" ]]; then
        echo "  ✗ ERROR: No se encontraron archivos CSV para el num_scn ${num_scn}"
        return 1
    fi
    
    echo "  ✓ Archivos encontrados:"
    echo "${archivos_encontrados}" | while read archivo; do
        echo "    - ${archivo}"
        echo "${archivo}" >> "${TMP_DIR}/archivos_${cod_pto}_${num_rdd}.lst"
    done
    
    return 0
}

################################################################################
# Funciones de transferencia SFTP
################################################################################

transferir_archivos() {
    local cod_pto=$1
    local num_rdd=$2
    local lista_archivos="${TMP_DIR}/archivos_${cod_pto}_${num_rdd}.lst"
    
    if [[ ! -f "${lista_archivos}" ]]; then
        echo "  ✗ ERROR: No hay archivos para transferir"
        return 1
    fi
    
    echo ""
    echo "  → Preparando transferencia SFTP..."
    echo "    Servidor: ${SERVIDOR_REMOTO}"
    echo "    Usuario: ${USUARIO_REMOTO}"
    echo "    Destino: ${RUTA_REMOTA}"
    echo ""
    
    local batch_file="${TMP_DIR}/sftp_batch_${cod_pto}_${num_rdd}.txt"
    
    # Generar comandos SFTP
    echo "cd ${RUTA_REMOTA}" > "${batch_file}"
    
    while IFS= read -r archivo; do
        echo "put ${REPORTE_DIR}/${archivo}" >> "${batch_file}"
    done < "${lista_archivos}"
    
    echo "quit" >> "${batch_file}"
    
    echo "  Archivos a transferir:"
    cat "${lista_archivos}" | while read arch; do
        echo "    - ${arch}"
    done
    
    echo ""
    echo "  → Ejecutando transferencia SFTP..."
    echo "    Se solicitará el password del servidor remoto"
    echo ""
    
    # Ejecutar SFTP con batch
    sftp -b "${batch_file}" "${USUARIO_REMOTO}@${SERVIDOR_REMOTO}"
    local resultado=$?
    
    if [[ ${resultado} -ne 0 ]]; then
        echo ""
        echo "  ✗ ERROR: Fallo la transferencia SFTP (codigo: ${resultado})"
        return 1
    fi
    
    echo ""
    echo "  ✓ Archivos transferidos exitosamente"
    return 0
}

################################################################################
# Función principal de procesamiento
################################################################################

procesar_rdd() {
    local cod_pto=$1
    local num_rdd=$2
    
    echo ""
    echo "================================================================================"
    echo "Procesando RDD: ${cod_pto}/${num_rdd}"
    echo "================================================================================"
    echo ""
    
    # Paso 1: Validar rdd_cab
    if ! validar_rdd_cab "${cod_pto}" "${num_rdd}"; then
        return 1
    fi
    
    # Paso 2: Verificar ora_integra_envio
    verificar_ora_integra_envio "${cod_pto}" "${num_rdd}"
    local resultado=$?
    
    if [[ ${resultado} -eq 1 ]]; then
        # CASO 2: Insertar en ora_integra_envio
        if ! insertar_ora_integra_envio "${cod_pto}" "${num_rdd}"; then
            return 1
        fi
    elif [[ ${resultado} -eq 2 ]]; then
        return 1
    fi
    
    # Paso 3: Insertar en i15
    if ! insertar_i15 "${cod_pto}" "${num_rdd}"; then
        return 1
    fi
    
    echo ""
    echo "✓ Validaciones completadas para RDD ${cod_pto}/${num_rdd}"
    return 0
}

################################################################################
# Menú principal
################################################################################

main() {
    inicializar
    
    while true; do
        mostrar_encabezado
        
        # Solicitar ambiente
        solicitar_ambiente
        echo ""
        echo "Ambiente seleccionado: ${AMBIENTE}"
        pausar
        
        mostrar_encabezado
        
        # Solicitar si es lote o individual
        local es_lote=1
        if solicitar_lote; then
            es_lote=0
        fi
        
        # Capturar RDD(s)
        mostrar_encabezado
        if ! capturar_rdds ${es_lote}; then
            pausar
            continue
        fi
        
        pausar
        
        # Procesar cada RDD
        local errores=0
        local procesados=0
        
        while IFS='|' read -r cod_pto num_rdd; do
            if procesar_rdd "${cod_pto}" "${num_rdd}"; then
                ((procesados++))
            else
                ((errores++))
            fi
        done < "${TMP_DIR}/rdds_lista.txt"
        
        echo ""
        echo "───────────────────────────────────────────────────────────────────────────────"
        echo "Resumen de validaciones:"
        echo "  Procesados exitosamente: ${procesados}"
        echo "  Con errores: ${errores}"
        echo "───────────────────────────────────────────────────────────────────────────────"
        pausar
        
        if [[ ${procesados} -eq 0 ]]; then
            echo "No hay RDD válidos para continuar"
            pausar
            continue
        fi
        
        # Ejecutar proceso 4GL
        if ! ejecutar_4gl; then
            if ! confirmar "¿Desea reiniciar el proceso?"; then
                exit 0
            fi
            continue
        fi
        
        # Buscar y transferir archivos para cada RDD procesado
        mostrar_encabezado
        echo "TRANSFERENCIA DE ARCHIVOS AL SERVIDOR REMOTO"
        echo "================================================================================"
        echo ""
        
        local transferidos=0
        local fallos_transferencia=0
        
        while IFS='|' read -r cod_pto num_rdd; do
            if buscar_archivos_generados "${cod_pto}" "${num_rdd}"; then
                if transferir_archivos "${cod_pto}" "${num_rdd}"; then
                    ((transferidos++))
                else
                    ((fallos_transferencia++))
                fi
            else
                ((fallos_transferencia++))
            fi
        done < "${TMP_DIR}/rdds_lista.txt"
        
        echo ""
        echo "================================================================================"
        echo "PROCESO COMPLETADO"
        echo "================================================================================"
        echo "  RDD procesados: ${procesados}"
        echo "  Archivos transferidos exitosamente: ${transferidos}"
        echo "  Fallos en transferencia: ${fallos_transferencia}"
        echo "================================================================================"
        echo ""
        
        if ! confirmar "¿Desea procesar más RDD?"; then
            echo ""
            echo "Finalizando proceso..."
            exit 0
        fi
        
        # Limpiar archivos temporales para nueva iteración
        rm -f "${TMP_DIR}"/*.unl "${TMP_DIR}"/*.lst "${TMP_DIR}"/*.txt
    done
}

################################################################################
# Punto de entrada
################################################################################

main "$@"
FINAL_DEL_SCRIPT

