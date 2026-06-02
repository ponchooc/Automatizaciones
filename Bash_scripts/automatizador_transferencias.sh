#!/bin/bash
# Automatizador de transferencias - Ejecución sin interacción con el usuario
# Fecha: 10/03/2026
#####################################################################################################

# Configuración de colores
VERDE="\033[1;32m"
CYAN="\033[1;36m"
MAGENTA="\033[1;35m"
AMARILLO="\033[1;33m"
RESET="\033[0m"
NEGRITA="\033[1m"

# Archivos y rutas
ARCHIVO_I04="/respaldo_migracion/reportes_gnx/i04.txt"
RUTA_REPORTES="/respaldo_migracion/reportes_gnx"

# Función para centrar texto
centrar_texto() {
    texto="$1"
    ancho=$(tput cols 2>/dev/null || echo 80)
    padding=$(( (ancho - ${#texto}) / 2 ))
    printf "%${padding}s%s\n" "" "$texto"
}

# Función para mostrar mensajes de éxito
mostrar_exito() {
    echo -e "${VERDE}[OK] $1${RESET}"
    sleep 1
}

# Función para mostrar mensajes de error
mostrar_error() {
    echo -e "${MAGENTA}[ERROR] $1${RESET}"
    sleep 2
}

# Función para mostrar mensajes de información
mostrar_info() {
    echo -e "${CYAN}[i] $1${RESET}"
}

# Función para ejecutar consulta SQL
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

# Función para renombrar archivo reciente
renombrar_archivo_reciente() {
    patron=$1
    nuevo_prefijo=$2

    archivo_reciente=$(ls -lt $RUTA_REPORTES/${patron}*.csv 2>/dev/null | grep " desa " | sed -n '1p' | awk '{print $NF}')

    if [ -n "$archivo_reciente" ]; then
        nuevo_nombre=$(echo "$archivo_reciente" | sed "s/${patron}/${nuevo_prefijo}/")
        echo -e "${VERDE}Renombrando: $archivo_reciente -> $nuevo_nombre${RESET}"
        mv "$archivo_reciente" "$nuevo_nombre"
        if [ $? -ne 0 ]; then
            mostrar_error "Error al renombrar $archivo_reciente"
            return 1
        fi
        mostrar_exito "Archivo renombrado correctamente"
        return 0
    else
        mostrar_info "No se encontraron archivos ${patron}*.csv del usuario desa"
        return 0
    fi
}

# Función para renombrar todos los archivos transferencia
renombrar_archivos_transferencia() {
    renombrar_archivo_reciente "TR_ORDENTIE" "SE_ORDENTIE"
    renombrar_archivo_reciente "TR_OT" "SE_OT"
    mostrar_exito "Proceso de renombrado completado"
}

# Función para mostrar resultados de transferencias
mostrar_resultados_transferencias() {
    clear
    echo -e "${NEGRITA}"
    centrar_texto "RESULTADO DEL PROCESAMIENTO"
    echo -e "${RESET}"
    echo "----------------------------------------"

    echo -e "${CYAN}Fecha y hora de finalización: $(date)${RESET}"
    echo

    echo -e "${AMARILLO}Archivos procesados:${RESET}"

    archivo_ordentie=$(ls -lt $RUTA_REPORTES/SE_ORDENTIE*.csv 2>/dev/null | grep " desa " | sed -n '1p' | awk '{print $NF}')
    archivo_ot=$(ls -lt $RUTA_REPORTES/SE_OT*.csv 2>/dev/null | grep " desa " | sed -n '1p' | awk '{print $NF}')

    if [ -n "$archivo_ordentie" ] && [ -f "$archivo_ordentie" ]; then
        ls -l "$archivo_ordentie"
    fi

    if [ -n "$archivo_ot" ] && [ -f "$archivo_ot" ]; then
        ls -l "$archivo_ot"
    fi

    echo
    echo "Líneas de los archivos:"

    total_lineas=0
    archivos_procesados=0

    if [ -n "$archivo_ordentie" ] && [ -f "$archivo_ordentie" ]; then
        nombre_archivo=$(basename "$archivo_ordentie")
        lineas=$(wc -l < "$archivo_ordentie")
        printf "%8d %s\n" "$lineas" "$nombre_archivo"
        total_lineas=$((total_lineas + lineas))
        archivos_procesados=$((archivos_procesados + 1))
    fi

    if [ -n "$archivo_ot" ] && [ -f "$archivo_ot" ]; then
        nombre_archivo=$(basename "$archivo_ot")
        lineas=$(wc -l < "$archivo_ot")
        printf "%8d %s\n" "$lineas" "$nombre_archivo"
        total_lineas=$((total_lineas + lineas))
        archivos_procesados=$((archivos_procesados + 1))
    fi

    printf "%8d total\n" "$total_lineas"

    if [ $archivos_procesados -eq 0 ]; then
        echo -e "${MAGENTA}No se encontraron archivos SE_* del usuario desa${RESET}"
    fi
}

# ─── INICIO ────────────────────────────────────────────────────────────────────

clear
echo -e "${NEGRITA}"
centrar_texto "AUTOMATIZADOR DE TRANSFERENCIAS"
echo -e "${RESET}"
echo "----------------------------------------"
echo -e "${AMARILLO}Fecha: $(date +"%d/%m/%Y %H:%M:%S")${RESET}"
echo -e "${AMARILLO}Sistema: G E N E S I X${RESET}"
echo "----------------------------------------"
echo

# Validar existencia del archivo i04.txt
if [ ! -f "$ARCHIVO_I04" ]; then
    mostrar_error "El archivo $ARCHIVO_I04 no existe. Hoy no hay nada para procesar."
    exit 1
fi

# Validar que el archivo tenga contenido
num_lineas=$(wc -l < "$ARCHIVO_I04" 2>/dev/null || echo 0)

if [ "$num_lineas" -eq 0 ]; then
    mostrar_info "El archivo $ARCHIVO_I04 está vacío. Hoy no hay nada para procesar."
    exit 0
fi

mostrar_info "Archivo encontrado con ${NEGRITA}$num_lineas${RESET}${CYAN} líneas. Iniciando procesamiento...${RESET}"
echo

# Cambiar al directorio de trabajo
cd /gnx_prod/sears/sql/ora_integra/ || {
    mostrar_error "No se pudo acceder al directorio /gnx_prod/sears/sql/ora_integra/"
    exit 1
}

echo -e "${AMARILLO}Iniciando procesamiento de transferencias...${RESET}"
echo "============================================"

# Ejecutar llenado_i04.sql
ejecutar_consulta_sql "llenado_i04.sql"
if [ $? -ne 0 ]; then
    exit 1
fi

# Ejecutar llenado_i13.sql
ejecutar_consulta_sql "llenado_i13.sql"
if [ $? -ne 0 ]; then
    exit 1
fi

# Ejecutar programa principal de transferencias
echo -e "${AMARILLO}Ejecutando fglgo ora_tran_10...${RESET}"
fglgo ora_tran_10

if [ $? -ne 0 ]; then
    mostrar_error "Error al ejecutar fglgo ora_tran_10"
    exit 1
fi
mostrar_exito "Procesamiento terminado"

# Renombrar archivos generados
echo -e "${CYAN}Renombrando archivos...${RESET}"
renombrar_archivos_transferencia

# Mostrar resultados
mostrar_resultados_transferencias

echo
echo -e "${VERDE}${NEGRITA}El proceso concluyó con éxito${RESET}"
echo

