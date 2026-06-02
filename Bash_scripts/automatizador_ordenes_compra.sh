#!/bin/bash
# Script automatizado para reenvio de ordenes de compra
# Basado en la logica de opcion 1 de Revis_ord_tra_REVI.sh
#####################################################################################################

# Configuracion de colores
VERDE="\033[1;32m"
CYAN="\033[1;36m"
MAGENTA="\033[1;35m"
AMARILLO="\033[1;33m"
RESET="\033[0m"
NEGRITA="\033[1m"

# Archivos y rutas
ARCHIVO_I04="/respaldo_migracion/reportes_gnx/i04.txt"
RUTA_REPORTES="/respaldo_migracion/reportes_gnx"
RUTA_SQL="/gnx_prod/sears/sql/ora_integra/"

# Funcion para centrar texto
centrar_texto() {
    texto="$1"
    ancho=$(tput cols 2>/dev/null || echo 80)
    padding=$(( (ancho - ${#texto}) / 2 ))
    printf "%${padding}s%s\n" "" "$texto"
}

# Funcion para mostrar mensajes de exito
mostrar_exito() {
    echo -e "${VERDE}[OK] $1${RESET}"
    sleep 1
}

# Funcion para mostrar mensajes de error
mostrar_error() {
    echo -e "${MAGENTA}[ERROR] $1${RESET}"
    sleep 2
}

# Funcion para mostrar mensajes de informacion
mostrar_info() {
    echo -e "${CYAN}[i] $1${RESET}"
}

# Funcion para verificar la existencia de un archivo
verificar_archivo() {
    if [ ! -f "$1" ]; then
        mostrar_error "El archivo $1 no existe"
        return 1
    fi
    return 0
}

# Funcion para ejecutar consulta SQL
ejecutar_consulta_sql() {
    consulta="$1"

    echo -e "${AMARILLO}Ejecutando dbaccess gen $consulta...${RESET}"
    dbaccess gen $consulta

    if [ $? -ne 0 ]; then
        mostrar_error "Error al ejecutar la consulta SQL"
        return 1
    fi
    mostrar_exito "Consulta SQL ejecutada correctamente"
    return 0
}

# Funcion para ejecutar programa fglgo
ejecutar_programa_fglgo() {
    programa="$1"
    comando="$2"

    echo -e "${AMARILLO}Ejecutando $comando $programa...${RESET}"
    $comando $programa

    if [ $? -ne 0 ]; then
        mostrar_error "Error al ejecutar $comando $programa"
        return 1
    fi
    mostrar_exito "Programa ejecutado correctamente"
    return 0
}

# Funcion para mostrar resultados de ordenes de compra
mostrar_resultados_ordenes_compra() {
    clear
    echo -e "${NEGRITA}"
    centrar_texto "RESULTADO DEL PROCESAMIENTO"
    echo -e "${RESET}"
    echo "----------------------------------------"

    echo -e "${CYAN}Fecha y hora de finalizacion: $(date)${RESET}"
    echo

    echo -e "${AMARILLO}Archivos procesados (usuario desa):${RESET}"

    archivos_oc=$(ls -lt $RUTA_REPORTES/SE_OC*.csv 2>/dev/null | grep " desa " | sed -n '1p' | awk '{print $NF}')
    archivos_asn=$(ls -lt $RUTA_REPORTES/SE_ASN_PO*.csv 2>/dev/null | grep " desa " | sed -n '1p' | awk '{print $NF}')

    if [ -n "$archivos_oc" ] && [ -f "$archivos_oc" ]; then
        ls -l "$archivos_oc"
    fi

    if [ -n "$archivos_asn" ] && [ -f "$archivos_asn" ]; then
        ls -l "$archivos_asn"
    fi

    echo
    echo "Lineas de los archivos:"

    total_lineas=0
    archivos_encontrados=0

    if [ -n "$archivos_oc" ] && [ -f "$archivos_oc" ]; then
        nombre_archivo=$(basename "$archivos_oc")
        lineas=$(wc -l < "$archivos_oc" 2>/dev/null || echo 0)
        printf "%8d %s\n" "$lineas" "$nombre_archivo"
        total_lineas=$((total_lineas + lineas))
        archivos_encontrados=$((archivos_encontrados + 1))
    fi

    if [ -n "$archivos_asn" ] && [ -f "$archivos_asn" ]; then
        nombre_archivo=$(basename "$archivos_asn")
        lineas=$(wc -l < "$archivos_asn" 2>/dev/null || echo 0)
        printf "%8d %s\n" "$lineas" "$nombre_archivo"
        total_lineas=$((total_lineas + lineas))
        archivos_encontrados=$((archivos_encontrados + 1))
    fi

    printf "%8d total\n" "$total_lineas"

    if [ $archivos_encontrados -eq 0 ]; then
        echo -e "${MAGENTA}No se encontraron archivos generados del usuario desa${RESET}"
    fi
}

# Funcion automatica para validar y depurar el archivo de entrada sin interaccion
preparar_archivo_entrada_automatico() {
    archivo_entrada="$1"

    verificar_archivo "$archivo_entrada"
    if [ $? -ne 0 ]; then
        return 1
    fi

    tmp_archivo="${archivo_entrada}.tmp.$$"
    cat /dev/null > "$tmp_archivo" || { mostrar_error "No se pudo crear archivo temporal"; return 1; }

    ordenes_validas=0
    lineas_invalidas=0

    while IFS= read -r linea || [ -n "$linea" ]; do
        if echo "$linea" | grep -q "^[0-9][0-9]*$"; then
            echo "$linea" >> "$tmp_archivo"
            ordenes_validas=$((ordenes_validas + 1))
        elif [ -n "$linea" ]; then
            lineas_invalidas=$((lineas_invalidas + 1))
        fi
    done < "$archivo_entrada"

    cat "$tmp_archivo" > "$archivo_entrada"
    rm -f "$tmp_archivo"

    if [ "$lineas_invalidas" -gt 0 ]; then
        mostrar_info "Se ignoraron $lineas_invalidas lineas invalidas del archivo de entrada"
    fi

    mostrar_info "Ordenes de compra validas a procesar: ${NEGRITA}$ordenes_validas${RESET}"

    if [ "$ordenes_validas" -eq 0 ]; then
        mostrar_info "No hay ordenes de compra para procesar"
        return 2
    fi

    return 0
}

# Flujo principal automatizado para reenvio de ordenes de compra
procesar_reenvio_ordenes_automatico() {
    clear
    echo -e "${NEGRITA}"
    centrar_texto "REENVIO AUTOMATIZADO DE ORDENES DE COMPRA"
    echo -e "${RESET}"
    echo "----------------------------------------"

    preparar_archivo_entrada_automatico "$ARCHIVO_I04"
    resultado_preparacion=$?

    if [ $resultado_preparacion -eq 1 ]; then
        mostrar_error "No fue posible preparar el archivo de entrada"
        return 1
    fi

    if [ $resultado_preparacion -eq 2 ]; then
        mostrar_info "Fin de programa sin procesamiento"
        return 0
    fi

    echo
    echo -e "${AMARILLO}Iniciando procesamiento de ordenes de compra...${RESET}"
    echo "============================================"

    ejecutar_consulta_sql "llenado_i04.sql"
    if [ $? -ne 0 ]; then
        return 1
    fi

    ejecutar_programa_fglgo "ora_compras_10" "fglgo"
    if [ $? -ne 0 ]; then
        return 1
    fi

    sleep 3
    mostrar_resultados_ordenes_compra

    echo
    mostrar_info "Depurando archivo de entrada $ARCHIVO_I04"
    cat /dev/null > "$ARCHIVO_I04"
    if [ $? -ne 0 ]; then
        mostrar_error "No se pudo depurar el archivo $ARCHIVO_I04"
        return 1
    fi
    mostrar_exito "El archivo $ARCHIVO_I04 fue depurado y quedo vacio"

    echo
    echo -e "${VERDE}${NEGRITA}Fin de programa${RESET}"
    return 0
}

# Funcion principal automatizada
main() {
    cd "$RUTA_SQL" || {
        mostrar_error "No se pudo acceder a $RUTA_SQL"
        exit 1
    }

    procesar_reenvio_ordenes_automatico
    exit $?
}

# Iniciar ejecucion del script
main

