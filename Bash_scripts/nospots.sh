#!/bin/bash
# Autor: Carlos Alfonso Ortega Molina
# Descripcion: script para procesar reportes de no spots validando precios y skus
# Sistema: AIX 7.2 | Informix 12.10 | Bash 5.2

# configuracion de variables
base_datos="gen"
ruta_trabajo="/gnx_prod/manto/desa/trabajo/sears/carlos_ortega"
ruta_bts="/gnx_prod/manto/desa/trabajo/sears/mangel/integridad"
archivo_entrada="$ruta_trabajo/GPOSAN_VAL_NO_SPOTS.csv"
archivo_procesado="$ruta_trabajo/GPOSAN_VAL_NO_SPOTS._PROCESADO.csv"
archivo_sales_checks="$ruta_trabajo/sales_checks_temp.txt"
archivo_bts_salida="$ruta_trabajo/bts_salida_temp.txt"
archivo_inconsistencias="$ruta_trabajo/inconsistencias_temp.txt"

# validacion de archivo de entrada
if [[ ! -f "$archivo_entrada" ]]; then
    echo "error: no se encuentra el archivo $archivo_entrada en la carpeta actual"
    exit 1
fi

# limpiar archivos temporales si existen
rm -f "$archivo_sales_checks" "$archivo_bts_salida" "$archivo_inconsistencias"

# extraer cabecera y titulo
linea_titulo=$(head -n 1 "$archivo_entrada")
linea_cabecera=$(head -n 2 "$archivo_entrada" | tail -n 1)

# extraer sales checks unicos y validar longitud
# el sales check esta en la columna 15
tail -n +3 "$archivo_entrada" | awk -F'","' '{print $15}' | sed 's/"//g' | sort -u | while read sales_check; do
    longitud=${#sales_check}
    if [[ $longitud -le 16 ]]; then
        # padding con ceros a la izquierda hasta 16 posiciones
        printf "%016d\n" "$sales_check" >> "$archivo_sales_checks"
    else
        echo "sales check $sales_check tiene mas de 16 digitos" >> "$archivo_inconsistencias"
    fi
done

# validacion si hay sales checks para procesar
if [[ -f "$archivo_sales_checks" ]]; then
    # ejecutar herramienta bts_bat usando rutas absolutas
    cd "$ruta_bts" || exit 1
    fglgo bts_bat "$archivo_sales_checks" "$archivo_bts_salida"
    cd "$ruta_trabajo" || exit 1
fi

# generar archivo de salida con cabeceras
echo "$linea_titulo" > "$archivo_procesado"
echo "$linea_cabecera" >> "$archivo_procesado"

# procesar cada sales check del reporte original para reconstruir con datos reales
if [[ -f "$archivo_bts_salida" ]]; then
    # crear un mapeo de datos de bts para busqueda rapida
    # formato bts: sales_check|sku|descripcion|precio|...
    while IFS="|" read sc_bts sku_bts desc_bts _; do
        # validar que la descripcion empiece con NS
        if [[ "$desc_bts" != NS* ]]; then
            continue
        fi

        # limpiar espacios del sku de bts para el query
        sku_clean=$(echo "$sku_bts" | sed 's/ //g')
        # padding de espacios a la izquierda para completar 15 posiciones en int_art
        sku_query=$(printf "%15s" "$sku_clean")
        
        # obtener la tienda (primeros 3 digitos despues de los ceros)
        # asumiendo que el sales check de 16 posiciones tiene el formato 0TTT...
        tienda=$(echo "$sc_bts" | cut -c 2-4)
        
        # consulta a tabla puntos para cod_tar
        cod_tar=$(echo "set isolation to dirty read; select first 1 cod_tar from puntos where cod_emp=1 and cod_pto=$tienda" | dbaccess "$base_datos" 2>/dev/null | grep -E "^[ ]*[0-9]+$" | tr -d ' ')
        
        if [[ -n "$cod_tar" ]]; then
            # consulta a tabla tarifas para pvp_tar con el int_art formateado
            pvp_tar=$(echo "set isolation to dirty read; select first 1 pvp_tar from tarifas where cod_tar=$cod_tar and int_art='$sku_query'" | dbaccess "$base_datos" 2>/dev/null | grep -E "^[ ]*[0-9.]+$" | tr -d ' ' | cut -d'.' -f1)
        else
            pvp_tar="0"
        fi
        
        # buscar la primera linea coincidente en el csv para usarla como base
        sc_original=$(echo "$sc_bts" | sed 's/^0*//')
        linea_orig=$(grep "\"$sc_original\"" "$archivo_entrada" | head -n 1)
        
        if [[ -n "$linea_orig" ]]; then
            # reconstruir la linea manteniendo el formato
            # se actualiza el sku (col 7) y el precio (col 16) segun bts y querys
            # campo 7: sku_bts, campo 16: pvp_tar
            echo "$linea_orig" | awk -v sku="SRS$sku_clean" -v precio="$pvp_tar" -F'","' 'BEGIN{OFS="\",\""} {$7=sku; $16=precio; print $0}' >> "$archivo_procesado"
        fi
    done < "$archivo_bts_salida"
fi

# agregar inconsistencias si existen
if [[ -f "$archivo_inconsistencias" ]]; then
    echo "\"\"" >> "$archivo_procesado"
    echo "\"Inconsistencias Detectadas\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\"" >> "$archivo_procesado"
    while read error; do
        echo "\"$error\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\"" >> "$archivo_procesado"
    done < "$archivo_inconsistencias"
fi

# limpieza final
rm -f "$archivo_sales_checks" "$archivo_bts_salida" "$archivo_inconsistencias"

echo "proceso completado. archivo generado: $archivo_procesado"

# Autor: Carlos Alfonso Ortega Molina

