#!/bin/bash
# script de procesamiento diario para archivos y bases de datos
# creado para la mesa de control de oracle
# Autor:Carlos Alfonso Ortega Molina

# definicion de variables
rutaPrincipal="/gnx_prod/manto/desa/trabajo/sears/ORACLE"
rutaDestino="/gnx_prod/manto/desa/trabajo/sears/ORACLE/archi_diarios"
archivoDestinatarios="${rutaPrincipal}/NO_BORRAR_destinatarios.txt"
baseDatos="gen"
remitente="desa@sears33.sanborns.net"
fechaActual=$(date +"%d%m")
logFile="/tmp/proceso_diario_${fechaActual}.log"
tempDir="/tmp/mail_temp_$$"

# funcion para registrar mensajes en el log
registrar_mensaje() {
    local mensaje="$1"
    local tipo="${2:-INFO}"
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] [${tipo}] ${mensaje}" >> "$logFile"
    echo "[${tipo}] ${mensaje}"
}

# funcion para limpiar al salir
limpiar() {
    if [ -d "$tempDir" ]; then
        rm -rf "$tempDir"
    fi
    registrar_mensaje "limpieza realizada"
}

# configurar limpieza al salir
trap limpiar EXIT

# funcion principal
main() {
    registrar_mensaje "iniciando script de procesamiento diario"

    # crear directorio temporal para archivos de correo
    mkdir -p "$tempDir"
    if [ $? -ne 0 ]; then
        registrar_mensaje "no se pudo crear directorio temporal" "ERROR"
        exit 1
    fi

    # ir a la carpeta principal
    cd "$rutaPrincipal"
    if [ $? -ne 0 ]; then
        registrar_mensaje "no se puede acceder a $rutaPrincipal" "ERROR"
        exit 1
    fi

    # verificar que existe la carpeta destino
    if [ ! -d "$rutaDestino" ]; then
        mkdir -p "$rutaDestino"
        if [ $? -ne 0 ]; then
            registrar_mensaje "no se pudo crear $rutaDestino" "ERROR"
            exit 1
        fi
    fi

    # 1. verificar archivos existentes, renombrarlos y moverlos
    registrar_mensaje "verificando archivos existentes"
    for archivo in match_13_14.txt match_resto.txt OC_val_I.txt OC_val_N.txt cliyclitie.txt; do
        if [ -f "$archivo" ]; then
            registrar_mensaje "archivo $archivo encontrado"
            mv "$archivo" "${rutaDestino}/${archivo%.*}_${fechaActual}.${archivo##*.}"
            if [ $? -eq 0 ]; then
                registrar_mensaje "archivo $archivo renombrado y movido correctamente"
            else
                registrar_mensaje "error al renombrar y mover $archivo" "ERROR"
            fi
        else
            registrar_mensaje "archivo $archivo no encontrado" "ADVERTENCIA"
        fi
    done

    # 2. ejecutar scripts sql que generaran los nuevos archivos
    registrar_mensaje "ejecutando scripts sql"
    for script in Obt_OC_Importa.sql Obt_OC_Nacional.sql TRA_Vic.sql cli_tie.sql; do
        if [ -f "$script" ]; then
            registrar_mensaje "ejecutando script $script"
            dbaccess "$baseDatos" "$script"
            if [ $? -eq 0 ]; then
                registrar_mensaje "script $script ejecutado correctamente"
            else
                registrar_mensaje "error al ejecutar script $script" "ERROR"
            fi
        else
            registrar_mensaje "script $script no encontrado" "ERROR"
        fi
    done

    # 3. enviar un solo correo con los 5 archivos recien generados comprimidos en un zip
    registrar_mensaje "preparando envio de correo"

    if [ ! -f "$archivoDestinatarios" ]; then
        registrar_mensaje "archivo de destinatarios no encontrado" "ERROR"
    else
        destinatarios=$(cat "$archivoDestinatarios" | tr '\n' ' ')
        if [ -z "$destinatarios" ]; then
            registrar_mensaje "no se encontraron destinatarios" "ERROR"
        else
            archivosExisten=true
            for archivo in match_13_14.txt match_resto.txt OC_val_I.txt OC_val_N.txt cliyclitie.txt; do
                if [ ! -f "$archivo" ]; then
                    registrar_mensaje "archivo $archivo no encontrado, podria faltar un adjunto" "ADVERTENCIA"
                    archivosExisten=false
                fi
            done

            registrar_mensaje "enviando correo con los archivos comprimidos"

            # Crear el archivo zip con los 5 archivos
            zipFile="${tempDir}/archivos_conciliacion_${fechaActual}.zip"
            zip -j "$zipFile" match_13_14.txt match_resto.txt OC_val_I.txt OC_val_N.txt cliyclitie.txt

            if [ $? -ne 0 ]; then
                registrar_mensaje "error al crear el archivo zip" "ERROR"
            else
                # Crear el cuerpo del mensaje
                cat > "${tempDir}/body.txt" << EOF
Buen día les envío los archivos para que puedan realizar su respectivo match, adjuntos en un solo archivo comprimido (zip).

PRIMER ARCHIVO (match_13_14.txt)
Número de transferencia
Unidad que envía la mercancía (870 u 840)
Unidad que recibe la mercancía (870, 840 o tienda que recibe)
Esto cubre la i13, i14 y la parte del recibirse en Vallejo o Cigatam desde tienda y parte de la i40, de tiendas a 870 u 840.

SEGUNDO ARCHIVO  (match_40.txt)
Número de transferencia
Tienda que envía la mercancía (diferente de 870 u 840)
Tienda que recibe la mercancía ( diferente de 870 u 840)
Esto cubre la i40 entre tiendas

TERCER ARCHIVO (OC_val_I.txt)
Ordenes de Compra de Importación

CUARTO ARCHIVO (OC_val_N.txt)
Ordenes de Compra Nacionales

QUINTO ARCHIVO (cliyclitie.txt)
Ordenes CLI & CLI_TIE

Saludos.
EOF

                # Enviar el correo con el zip adjunto
                (
                    cat "${tempDir}/body.txt"
                    uuencode "$zipFile" "$(basename "$zipFile")"
                ) | mailx -s "Base para c conciliacion Genesix vs Oracle" $destinatarios

                if [ $? -eq 0 ]; then
                    registrar_mensaje "correo con archivo zip enviado correctamente"
                else
                    registrar_mensaje "error al enviar correo" "ERROR"
                fi
            fi
        fi
    fi

    registrar_mensaje "procesamiento diario finalizado"
    return 0
}

# ejecutar funcion principal
main
exit $?
