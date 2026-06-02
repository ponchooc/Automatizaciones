#!/bin/bash
# script de administracion para devoluciones al proveedor
# fecha 16/06/2025
# elaborado por Carlos Alfonso Ortega Molina
#####################################################################################################

# configuracion de colores
VERDE="\033[0;32m"
ROJO="\033[0;31m"
AZUL="\033[0;34m"
AMARILLO="\033[1;33m"
RESET="\033[0m"
NEGRITA="\033[1m"

# archivos y rutas
ARCHIVO_DEVOLUCIONES="/tmp/devs_$(date +%M%S)_$$.txt"
ARCHIVO_BULTOS="/tmp/bultos_$(date +%M%S)_$$.txt"
ARCHIVO_LOAD="/tmp/load_devs_$(date +%M%S)_$$.txt"
RUTA_REPORTES="/respaldo_migracion/reportes_gnx"
RUTA_FELIPE="/respaldo_migracion/reportes_gnx/felipe"
cd /respaldo_migracion/reportes_gnx/felipe
gunzip ora_devol*

# funcion para limpiar archivos temporales
limpiar_archivos_temporales() {
    # elimina todos los archivos temporales creados durante la ejecucion
    [ -f "$ARCHIVO_DEVOLUCIONES" ] && rm -f "$ARCHIVO_DEVOLUCIONES" 2>/dev/null
    [ -f "$ARCHIVO_BULTOS" ] && rm -f "$ARCHIVO_BULTOS" 2>/dev/null
    [ -f "$ARCHIVO_LOAD" ] && rm -f "$ARCHIVO_LOAD" 2>/dev/null
}

# funcion para centrar texto
centrar_texto() {
    # recibe el texto a centrar y lo muestra en pantalla centrado
    texto="$1"
    ancho=$(tput cols)
    padding=$(( (ancho - ${#texto}) / 2 ))
    printf "%${padding}s%s\n" "" "$texto"
}

# funcion para mostrar mensajes de exito
mostrar_exito() {
    # recibe un mensaje y lo muestra con formato de exito
    echo -e "${VERDE}[OK] $1${RESET}"
    sleep 1
}

# funcion para mostrar mensajes de error
mostrar_error() {
    # recibe un mensaje y lo muestra con formato de error
    echo -e "${ROJO}[ERROR] $1${RESET}"
    sleep 2
}

# funcion para mostrar mensajes de informacion
mostrar_info() {
    # recibe un mensaje y lo muestra con formato informativo
    echo -e "${AZUL}[i] $1${RESET}"
}

# funcion para capturar numeros de devoluciones
capturar_devoluciones() {
    # captura los numeros de devolucion del usuario
    clear
    echo -e "${NEGRITA}"
    centrar_texto "INGRESO DE DEVOLUCIONES"
    echo -e "${RESET}"
    echo "----------------------------------------"
    mostrar_info "ingrese los numeros de devoluciones uno por linea"
    mostrar_info "debe ingresar unicamente numeros enteros sin decimales ni letras"
    mostrar_info "para finalizar ingrese x o X"
    echo

    # crear archivo temporal para devoluciones
    touch "$ARCHIVO_DEVOLUCIONES" || { mostrar_error "no se pudo crear archivo temporal"; return 1; }

    while true; do
        read -p "> " numero

        # verificar si es la señal para terminar
        if [ "$numero" = "x" -o "$numero" = "X" ]; then
            break
        fi

        # validar que sea un numero entero metodo compatible con aix
        if echo "$numero" | grep -q "^[0-9][0-9]*$"; then
            echo "$numero" >> "$ARCHIVO_DEVOLUCIONES"
        else
            mostrar_error "entrada invalida ingrese solo numeros enteros"
        fi
    done

    # verificar que se capturaron devoluciones
    num_devoluciones=$(wc -l < "$ARCHIVO_DEVOLUCIONES")
    if [ "$num_devoluciones" -eq 0 ]; then
        mostrar_error "no se capturaron devoluciones"
        return 1
    fi

    echo
    mostrar_info "se han registrado ${NEGRITA}$num_devoluciones${RESET}${AZUL} devoluciones${RESET}"
    return 0
}

# funcion para mostrar resultados de consulta formateados
mostrar_consulta_formateada() {
    # muestra los resultados de la consulta en formato profesional
    lista_devoluciones="$1"
    
    echo -e "${AMARILLO}ejecutando consulta de devoluciones...${RESET}"
    
    # crear archivo temporal para la consulta
    archivo_consulta="/tmp/consulta_$$.sql"
    echo "select * from ora_int_corp where num_orp in ($lista_devoluciones);" > "$archivo_consulta"
    
    echo -e "${NEGRITA}"
    centrar_texto "RESULTADOS DE LA CONSULTA"
    echo -e "${RESET}"
    echo "=========================================="
    
    # ejecutar consulta y mostrar resultados
    dbaccess gen "$archivo_consulta" 2>/dev/null | grep -v "^$" | grep -v "Database selected"
    
    rm -f "$archivo_consulta" 2>/dev/null
    echo "=========================================="
    echo
    read -p "presione ENTER para continuar..."
}

# funcion para capturar bultos por devolucion
capturar_bultos() {
    # captura los bultos para cada devolucion
    clear
    echo -e "${NEGRITA}"
    centrar_texto "CAPTURA DE BULTOS POR DEVOLUCION"
    echo -e "${RESET}"
    echo "----------------------------------------"

    # crear archivo temporal para bultos
    touch "$ARCHIVO_BULTOS" || { mostrar_error "no se pudo crear archivo temporal de bultos"; return 1; }

    # leer devoluciones en un array primero
    devoluciones_array=""
    while read -r devolucion; do
        if [ -z "$devoluciones_array" ]; then
            devoluciones_array="$devolucion"
        else
            devoluciones_array="$devoluciones_array $devolucion"
        fi
    done < "$ARCHIVO_DEVOLUCIONES"

    # ahora pedir bultos para cada devolucion sin interferir con la lectura de archivo
    for devolucion in $devoluciones_array; do
        while true; do
            echo
            read -p "ingrese bultos para devolucion $devolucion: " bultos
            
            # validar que sea un numero entero positivo mayor que cero
            if echo "$bultos" | grep -q "^[0-9][0-9]*$" && [ "$bultos" -gt 0 ]; then
                echo "$devolucion,$bultos" >> "$ARCHIVO_BULTOS"
                mostrar_exito "bultos registrados para devolucion $devolucion"
                break
            else
                mostrar_error "debe ingresar un numero entero positivo mayor que cero"
            fi
        done
    done

    return 0
}

# funcion para aplicar updates de bultos
aplicar_updates_bultos() {
    # aplica los updates de bultos a la tabla devcab
    echo -e "${AMARILLO}aplicando actualizaciones de bultos...${RESET}"
    
    total_updates=0
    while IFS=',' read -r devolucion bultos; do
        # crear archivo temporal para el update
        archivo_update="/tmp/update_$$.sql"
        cat > "$archivo_update" << EOF
update devcab set ban_cdt = 'S', bultos = $bultos
where cod_emp = 1
and pto_emi = 999
and num_dev = $devolucion;
EOF
        
        # ejecutar update
        resultado=$(dbaccess gen "$archivo_update" 2>&1)
        
        if echo "$resultado" | grep -q "1 row(s) updated"; then
            mostrar_exito "devolucion $devolucion actualizada con $bultos bultos"
            total_updates=$((total_updates + 1))
        else
            mostrar_error "error al actualizar devolucion $devolucion"
        fi
        
        rm -f "$archivo_update" 2>/dev/null
    done < "$ARCHIVO_BULTOS"
    
    echo
    mostrar_info "total de registros actualizados: ${NEGRITA}$total_updates${RESET}"
    echo
    read -p "presione ENTER para continuar..."
    return 0
}

# funcion para vaciar tabla ora_dev_del
vaciar_tabla_ora_dev_del() {
    # vacia la tabla ora_dev_del
    echo -e "${AMARILLO}vaciando tabla ora_dev_del...${RESET}"
    
    archivo_delete="/tmp/delete_ora_dev_del_$$.sql"
    echo "delete from ora_dev_del;" > "$archivo_delete"
    
    resultado=$(dbaccess gen "$archivo_delete" 2>&1)
    rm -f "$archivo_delete" 2>/dev/null
    
    if [ $? -eq 0 ]; then
        mostrar_exito "tabla ora_dev_del vaciada correctamente"
    else
        mostrar_error "error al vaciar tabla ora_dev_del"
        return 1
    fi
    
    return 0
}

# funcion para cargar devoluciones en tabla ora_dev_del
cargar_devoluciones_tabla() {
    # carga las devoluciones en la tabla ora_dev_del usando load
    echo -e "${AMARILLO}cargando devoluciones en tabla ora_dev_del...${RESET}"
    
    # verificar que el archivo de devoluciones existe
    if [ ! -f "$ARCHIVO_DEVOLUCIONES" ]; then
        mostrar_error "archivo de devoluciones no existe"
        echo
        read -p "desea continuar con el proceso s/n: " continuar
        if [ "$continuar" != "s" -a "$continuar" != "S" ]; then
            mostrar_info "proceso cancelado por el usuario"
            return 1
        fi
        return 0
    fi
        # copiar devoluciones al archivo de load con ruta absoluta
    cp "$ARCHIVO_DEVOLUCIONES" "$ARCHIVO_LOAD" || { 
        mostrar_error "error al preparar archivo de carga"
        echo
        read -p "desea continuar con el proceso s/n: " continuar
        if [ "$continuar" != "s" -a "$continuar" != "S" ]; then
            mostrar_info "proceso cancelado por el usuario"
            return 1
        fi
        return 0
    }
    
    # crear archivo de control para load con ruta absoluta
    archivo_load_sql="/tmp/load_devs_$$.sql"
    cat > "$archivo_load_sql" << EOF
load from "$ARCHIVO_LOAD" insert into ora_dev_del;
EOF
    
    resultado=$(dbaccess gen "$archivo_load_sql" 2>&1)
    rm -f "$archivo_load_sql" 2>/dev/null
    
    # verificar si hay algun numero seguido de "row(s) loaded"
    if echo "$resultado" | grep -q "[0-9][0-9]* row(s) loaded"; then
        num_cargadas=$(echo "$resultado" | grep "row(s) loaded" | awk '{print $1}')
        mostrar_exito "$num_cargadas devoluciones cargadas en tabla ora_dev_del"
    else
        mostrar_error "error al cargar devoluciones en tabla ora_dev_del"
        echo "detalle del error: $resultado"
        echo
        read -p "desea continuar con el proceso s/n: " continuar
        if [ "$continuar" != "s" -a "$continuar" != "S" ]; then
            mostrar_info "proceso cancelado por el usuario"
            return 1
        fi
    fi
    
    return 0
}

# funcion para ejecutar deletes y reportar resultados
ejecutar_deletes() {
    # ejecuta los deletes en las tablas ora_int_corp y ora_int_dorp
    lista_devoluciones="$1"
    
    echo
    read -p "confirma que desea proceder con la eliminacion de registros s/n: " confirmar
    if [ "$confirmar" != "s" -a "$confirmar" != "S" ]; then
        mostrar_info "operacion de eliminacion cancelada por el usuario"
        return 0
    fi
    
    echo -e "${AMARILLO}ejecutando eliminaciones...${RESET}"
    
    # delete en ora_int_corp
    archivo_delete1="/tmp/delete_corp_$$.sql"
    echo "delete from ora_int_corp where num_orp in ($lista_devoluciones);" > "$archivo_delete1"
    
    resultado1=$(dbaccess gen "$archivo_delete1" 2>&1)
    rm -f "$archivo_delete1" 2>/dev/null
    
    if echo "$resultado1" | grep -q "row(s) deleted"; then
        num_deleted1=$(echo "$resultado1" | grep "row(s) deleted" | awk '{print $1}')
        mostrar_exito "$num_deleted1 registros eliminados de ora_int_corp"
    else
        mostrar_info "0 registros eliminados de ora_int_corp (posiblemente no existian)"
    fi
    
    # delete en ora_int_dorp
    archivo_delete2="/tmp/delete_dorp_$$.sql"
    echo "delete from ora_int_dorp where num_orp in ($lista_devoluciones);" > "$archivo_delete2"
    
    resultado2=$(dbaccess gen "$archivo_delete2" 2>&1)
    rm -f "$archivo_delete2" 2>/dev/null
    
    if echo "$resultado2" | grep -q "row(s) deleted"; then
        num_deleted2=$(echo "$resultado2" | grep "row(s) deleted" | awk '{print $1}')
        mostrar_exito "$num_deleted2 registros eliminados de ora_int_dorp"
    else
        mostrar_info "0 registros eliminados de ora_int_dorp (posiblemente no existian)"
    fi
    
    echo
    read -p "presione ENTER para continuar..."
    return 0
}

# funcion para ejecutar fglgo ora_devol
ejecutar_ora_devol() {
    # ejecuta el programa ora_devol en la carpeta felipe
    echo
    read -p "confirma que desea ejecutar fglgo ora_devol s/n: " confirmar
    if [ "$confirmar" != "s" -a "$confirmar" != "S" ]; then
        mostrar_info "ejecucion de ora_devol cancelada por el usuario"
        return 0
    fi
    
    echo -e "${AMARILLO}cambiando a directorio $RUTA_FELIPE...${RESET}"
    cd "$RUTA_FELIPE" || { mostrar_error "no se pudo acceder al directorio $RUTA_FELIPE"; return 1; }
    
    echo -e "${AMARILLO}ejecutando fglgo ora_devol...${RESET}"
    fglgo ora_devol
    
    if [ $? -eq 0 ]; then
        mostrar_exito "fglgo ora_devol ejecutado correctamente"
        return 0
    else
        mostrar_error "error al ejecutar fglgo ora_devol"
        echo
        read -p "desea continuar con el proceso s/n: " continuar
        if [ "$continuar" != "s" -a "$continuar" != "S" ]; then
            mostrar_info "proceso cancelado por el usuario"
            return 1
        fi
        return 0
    fi
}

# funcion para mostrar archivos generados
mostrar_archivos_generados() {
    # muestra los archivos generados mas recientes
    clear
    echo -e "${NEGRITA}"
    centrar_texto "ARCHIVOS GENERADOS"
    echo -e "${RESET}"
    echo "----------------------------------------"
    
    echo -e "${AZUL}fecha y hora de finalizacion: $(date)${RESET}"
    echo
    
    echo -e "${AMARILLO}archivos SE_ASN_DEV* mas recientes:${RESET}"
    ls -lt $RUTA_REPORTES/SE_ASN_DEV* 2>/dev/null | grep " desa " | head -2
    
    echo
    echo -e "${AMARILLO}archivos SE_ORDENTIE* mas recientes:${RESET}"
    ls -lt $RUTA_REPORTES/SE_ORDENTIE* 2>/dev/null | grep " desa " | head -2
    
    echo
    echo "lineas de los archivos:"
    
    # contar lineas de archivos SE_ASN_DEV*
    total_lineas=0
    for archivo in $(ls -t $RUTA_REPORTES/SE_ASN_DEV* 2>/dev/null | head -2); do
        if ls -l "$archivo" | grep -q " desa "; then
            nombre_archivo=$(basename "$archivo")
            lineas=$(wc -l < "$archivo")
            printf "%8d %s\n" "$lineas" "$nombre_archivo"
            total_lineas=$((total_lineas + lineas))
        fi
    done
    
    # contar lineas de archivos SE_ORDENTIE*
    for archivo in $(ls -t $RUTA_REPORTES/SE_ORDENTIE* 2>/dev/null | head -2); do
        if ls -l "$archivo" | grep -q " desa "; then
            nombre_archivo=$(basename "$archivo")
            lineas=$(wc -l < "$archivo")
            printf "%8d %s\n" "$lineas" "$nombre_archivo"
            total_lineas=$((total_lineas + lineas))
        fi
    done
    
    printf "%8d total\n" "$total_lineas"
}

# funcion principal para procesar devoluciones
procesar_devoluciones() {
    # gestiona el proceso completo de devoluciones al proveedor
    clear
    echo -e "${NEGRITA}"
    centrar_texto "DEVOLUCIONES AL PROVEEDOR"
    echo -e "${RESET}"
    echo "----------------------------------------"

    # capturar devoluciones
    capturar_devoluciones
    if [ $? -ne 0 ]; then
        return 1
    fi

    # crear lista para queries (formato: 'num1','num2','num3')
    lista_devoluciones=""
    while read -r devolucion; do
        if [ -z "$lista_devoluciones" ]; then
            lista_devoluciones="'$devolucion'"
        else
            lista_devoluciones="$lista_devoluciones,'$devolucion'"
        fi
    done < "$ARCHIVO_DEVOLUCIONES"

    # mostrar consulta inicial
    mostrar_consulta_formateada "$lista_devoluciones"

    # preguntar si habra correcciones
    echo
    read -p "habra correcciones de bultos s/n: " correcciones
    
    if [ "$correcciones" = "s" -o "$correcciones" = "S" ]; then
        # capturar bultos
        capturar_bultos
        if [ $? -ne 0 ]; then
            return 1
        fi
        
        # aplicar updates de bultos
        aplicar_updates_bultos
        if [ $? -ne 0 ]; then
            return 1
        fi
    fi

    # vaciar tabla ora_dev_del
    vaciar_tabla_ora_dev_del
    if [ $? -ne 0 ]; then
        return 1
    fi

    # cargar devoluciones en tabla
    cargar_devoluciones_tabla
    if [ $? -ne 0 ]; then
        return 1
    fi

    # ejecutar deletes
    ejecutar_deletes "$lista_devoluciones"
    if [ $? -ne 0 ]; then
        return 1
    fi

    # ejecutar ora_devol
    ejecutar_ora_devol
    if [ $? -ne 0 ]; then
        return 1
    fi

    # mostrar archivos generados
    mostrar_archivos_generados

    echo
    echo -e "${VERDE}${NEGRITA}el proceso de devoluciones concluyo con exito${RESET}"
    echo
    read -p "presione ENTER para volver al menu principal..."
    
    # limpiar archivos temporales al final
    limpiar_archivos_temporales
}

# funcion para mostrar el menu principal
mostrar_menu() {
    # muestra el menu principal del sistema
    clear
    echo -e "${NEGRITA}"
    centrar_texto "SISTEMA DE DEVOLUCIONES AL PROVEEDOR"
    echo -e "${RESET}"
    echo "----------------------------------------"
    echo -e "${AMARILLO}fecha: $(date +"%d/%m/%Y %H:%M:%S")${RESET}"
    echo -e "${AMARILLO}sistema: g e n e s i x ${RESET}"
    echo "----------------------------------------"
    echo -e "${NEGRITA}menu principal${RESET}"
    echo
    echo -e "  ${VERDE}1${RESET} procesar devoluciones al proveedor"
    echo -e "  ${ROJO}0${RESET} salir"
    echo
    echo "----------------------------------------"
    echo
}

# funcion principal
main() {
    # funcion principal que inicia el programa
    opcion=""

    while true; do
        mostrar_menu
        read -p "ingrese una opcion: " opcion

        case $opcion in
            1)
                procesar_devoluciones
                ;;
            0)
                clear
                echo -e "${VERDE}gracias por utilizar el sistema de devoluciones hasta pronto${RESET}"
                limpiar_archivos_temporales
                exit 0
                ;;
            *)
                mostrar_error "opcion invalida intente nuevamente"
                ;;
        esac
    done
}

# iniciar ejecucion del script
main
