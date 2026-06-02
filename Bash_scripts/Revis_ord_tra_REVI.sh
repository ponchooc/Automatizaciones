#!/bin/bash
# Script de administración para reenvío de órdenes de compra y transferencias
# Fecha: 11/08/2025
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

# Función para verificar la existencia de un archivo
verificar_archivo() {
    if [ ! -f "$1" ]; then
        mostrar_error "El archivo $1 no existe"
        return 1
    fi
    return 0
}

# Función para gestionar el archivo de entrada
gestionar_archivo_entrada() {
    tipo_proceso="$1"
    archivo_entrada="$2"

    verificar_archivo "$archivo_entrada"
    if [ $? -ne 0 ]; then
        echo -e "${AMARILLO}Creando archivo...${RESET}"
        touch "$archivo_entrada" || { mostrar_error "No se pudo crear el archivo"; return 1; }
    fi

    num_lineas=$(wc -l < "$archivo_entrada" 2>/dev/null || echo 0)

    echo -e "${CYAN}[i] El archivo actual contiene ${NEGRITA}$num_lineas${RESET}${CYAN} líneas${RESET}"
    echo
    echo -e "${NEGRITA}¿Qué desea hacer?${RESET}"
    echo -e "  ${VERDE}1${RESET} Borrar el contenido y crear nuevo archivo"
    echo -e "  ${VERDE}2${RESET} Trabajar con el archivo actual"
    echo

    read -p "Seleccione una opción [1/2]: " opcion_archivo

    case $opcion_archivo in
        1)
            cat /dev/null > "$archivo_entrada"
            mostrar_exito "Contenido del archivo borrado correctamente"
            sleep 2

            clear
            echo -e "${NEGRITA}"
            centrar_texto "INGRESO DE $tipo_proceso"
            echo -e "${RESET}"
            echo "----------------------------------------"
            mostrar_info "Ingrese los números de $tipo_proceso uno por línea"
            mostrar_info "Debe ingresar únicamente números enteros sin decimales ni letras"
            mostrar_info "Para finalizar ingrese x o X"
            echo

            while true; do
                read -p "> " numero

                if [ "$numero" = "x" ] || [ "$numero" = "X" ]; then
                    break
                fi

                if echo "$numero" | grep -q "^[0-9][0-9]*$"; then
                    echo "$numero" >> "$archivo_entrada"
                else
                    mostrar_error "Entrada inválida ingrese solo números enteros"
                fi
            done

            num_lineas=$(wc -l < "$archivo_entrada" 2>/dev/null || echo 0)
            echo
            mostrar_info "Se han registrado ${NEGRITA}$num_lineas${RESET}${CYAN} $tipo_proceso${RESET}"
            ;;

        2)
            echo
            mostrar_info "Se utilizará el archivo actual con ${NEGRITA}$num_lineas${RESET}${CYAN} líneas${RESET}"
            ;;

        *)
            mostrar_error "Opción inválida se utilizará el archivo actual"
            ;;
    esac

    return 0
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

# Función para ejecutar programa fglgo
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

# Función para mostrar resultados de órdenes de compra
mostrar_resultados_ordenes_compra() {
    clear
    echo -e "${NEGRITA}"
    centrar_texto "RESULTADO DEL PROCESAMIENTO"
    echo -e "${RESET}"
    echo "----------------------------------------"

    echo -e "${CYAN}Fecha y hora de finalización: $(date)${RESET}"
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
    echo "Líneas de los archivos:"

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

# Función para procesar reenvío de órdenes de compra
procesar_reenvio_ordenes() {
    clear
    echo -e "${NEGRITA}"
    centrar_texto "REENVÍO DE ÓRDENES DE COMPRA"
    echo -e "${RESET}"
    echo "----------------------------------------"

    gestionar_archivo_entrada "ÓRDENES DE COMPRA" "$ARCHIVO_I04"

    echo
    read -p "¿Desea proceder con el procesamiento? s/n: " confirmar
    if [ "$confirmar" != "s" ] && [ "$confirmar" != "S" ]; then
        mostrar_info "Operación cancelada por el usuario"
        read -p "Presione ENTER para volver al menú principal..."
        return
    fi

    echo
    echo -e "${AMARILLO}Iniciando procesamiento de órdenes de compra...${RESET}"
    echo "============================================"

    ejecutar_consulta_sql "llenado_i04.sql"
    if [ $? -ne 0 ]; then
        read -p "Presione ENTER para continuar..."
        return
    fi

    ejecutar_programa_fglgo "ora_compras_10" "fglgo"
    if [ $? -ne 0 ]; then
        read -p "Presione ENTER para continuar..."
        return
    fi

    sleep 3
    mostrar_resultados_ordenes_compra

    echo
    echo -e "${VERDE}${NEGRITA}Fin de programa${RESET}"
    echo
    read -p "Presione ENTER para volver al menú principal..."
}

# Función para procesar transferencias
procesar_transferencias() {
    clear
    echo -e "${NEGRITA}"
    centrar_texto "REENVÍO DE TRANSFERENCIAS"
    echo -e "${RESET}"
    echo "----------------------------------------"

    gestionar_archivo_entrada "TRANSFERENCIAS" "$ARCHIVO_I04"

    echo
    read -p "¿Desea proceder con el procesamiento? s/n: " confirmar
    if [ "$confirmar" != "s" ] && [ "$confirmar" != "S" ]; then
        mostrar_info "Operación cancelada por el usuario"
        read -p "Presione ENTER para volver al menú principal..."
        return
    fi

    echo
    echo -e "${AMARILLO}Iniciando procesamiento de transferencias...${RESET}"
    echo "============================================"

    ejecutar_consulta_sql "llenado_i04.sql"
    if [ $? -ne 0 ]; then
        read -p "Presione ENTER para continuar..."
        return
    fi

    ejecutar_consulta_sql "llenado_i13.sql"
    if [ $? -ne 0 ]; then
        read -p "Presione ENTER para continuar..."
        return
    fi

    echo
    read -p "¿Desea ejecutar gen9354x? s/n: " ejecutar_gen
    if [ "$ejecutar_gen" = "s" ] || [ "$ejecutar_gen" = "S" ]; then
        ejecutar_programa_fglgo "gen9354x" "zrun"
        if [ $? -ne 0 ]; then
            mostrar_error "Error en gen9354x, pero continuando..."
        fi
    fi

    echo -e "${AMARILLO}Ejecutando fglgo ora_tran_10...${RESET}"
    fglgo ora_tran_10

    if [ $? -ne 0 ]; then
        mostrar_error "Error al ejecutar fglgo ora_tran_10"
        read -p "Presione ENTER para continuar..."
        return
    fi
    mostrar_exito "Procesamiento terminado"

    echo -e "${CYAN}Renombrando archivos...${RESET}"
    renombrar_archivos_transferencia

    mostrar_resultados_transferencias

    echo
    echo -e "${VERDE}${NEGRITA}El proceso concluyó con éxito${RESET}"
    echo
    read -p "Presione ENTER para volver al menú principal..."
}

# Función para mostrar el menú principal
mostrar_menu() {
    clear
    echo -e "${NEGRITA}"
    centrar_texto "SISTEMA ÚNICO DE REENVÍOS"
    echo -e "${RESET}"
    echo "----------------------------------------"
    echo -e "${AMARILLO}Fecha: $(date +"%d/%m/%Y %H:%M:%S")${RESET}"
    echo -e "${AMARILLO}Sistema: G E N E S I X${RESET}"
    echo "----------------------------------------"
    echo -e "${NEGRITA}Menú principal${RESET}"
    echo
    echo -e "  ${VERDE}1${RESET} Reenvío de órdenes de compra"
    echo -e "  ${VERDE}2${RESET} Transferencias"
    echo -e "  ${MAGENTA}0${RESET} Salir"
    echo
    echo "----------------------------------------"
    echo
}

# Función principal
main() {
    opcion=""

    while true; do
        mostrar_menu
        read -p "Ingrese una opción: " opcion

        case $opcion in
            1)
                cd /gnx_prod/sears/sql/ora_integra/
                procesar_reenvio_ordenes
                ;;
            2)
                cd /gnx_prod/sears/sql/ora_integra/
                procesar_transferencias
                ;;
            0)
                clear
                echo -e "${VERDE}Gracias por utilizar el sistema, hasta pronto${RESET}"
                exit 0
                ;;
            *)
                mostrar_error "Opción inválida, intente nuevamente"
                ;;
        esac
    done
}

# Iniciar ejecución del script
main
