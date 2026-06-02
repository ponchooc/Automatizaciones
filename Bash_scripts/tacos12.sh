#!/usr/bin/bash

###############################################################################
# script: tacos5_con_validacion_bts.sh
# procesa un archivo csv de no spots
# aplica filtros, elimina duplicados, obtiene precios desde tarifas
# valida con BTS y ajusta lineas NS automaticamente
###############################################################################

###############################################################################
# validacion de parametros
###############################################################################

if [ $# -ne 1 ]; then
    echo "error: debe indicar el nombre del archivo csv a procesar"
    echo "uso: $0 <archivo_csv>"
    exit 1
fi

archivo="$1"

###############################################################################
# definicion de nombres de archivos de trabajo
###############################################################################

archivo_filtrado="${archivo%.*}_FILTRADO.csv"
archivo_salida="${archivo%.*}_PROCESADO.csv"
articulos_tmp="${archivo%.*}_ARTICULOS.tmp"
tarifas_tmp="${archivo%.*}_TARIFAS.tmp"
sql_tmp="${archivo%.*}_CONSULTA_TARIFAS.sql"
sales_checks_tmp="${archivo%.*}_SALES_CHECKS.txt"
bts_salida="${archivo%.*}_BTS.txt"
archivo_final="${archivo%.*}_FINAL.csv"

###############################################################################
# validaciones iniciales
###############################################################################

if [ ! -f "$archivo" ]; then
    echo "error: el archivo '$archivo' no existe"
    exit 2
fi

for f in "$archivo_filtrado" "$archivo_salida" "$articulos_tmp" "$tarifas_tmp" "$sql_tmp" "$sales_checks_tmp" "$bts_salida" "$archivo_final"; do
    if [ -f "$f" ]; then
        echo "error: el archivo '$f' ya existe"
        exit 3
    fi
done

###############################################################################
# paso 1: filtro de registros y eliminacion de duplicados
###############################################################################

echo "=== PASO 1: Filtrado y eliminacion de duplicados ==="

awk -F',' -v OFS=',' '
NR == 1 || NR == 2 {
    print
    next
}
{
    detalle_ns = $8
    precio_no_spots = $16
    gsub(/^"|"$/, "", detalle_ns)
    gsub(/^"|"$/, "", precio_no_spots)
    if (detalle_ns ~ /^NS/ || (detalle_ns == "" && precio_no_spots != "0" && precio_no_spots != "")) {
        lineas[++n] = $0
    }
}
END {
    for (i = 1; i <= n; i++) {
        ncampos = split(lineas[i], campos, FS)
        clave = ""
        for (j = 1; j <= ncampos; j++) {
            if (j == 2) continue
            clave = clave campos[j] "||"
        }
        if (!(clave in visto)) {
            print lineas[i]
            visto[clave] = 1
        }
    }
}
' "$archivo" > "$archivo_filtrado"

echo "registros despues de filtro: $(tail -n +3 "$archivo_filtrado" | wc -l)"

###############################################################################
# paso 2: extraccion de articulos unicos
###############################################################################

echo "=== PASO 2: Extraccion de articulos unicos ==="

tail -n +3 "$archivo_filtrado" | awk -F',' '
{
    campo_art = $4
    gsub(/^"|"$/, "", campo_art)
    if (campo_art ~ /^SRS[0-9]+$/) {
        gsub(/^SRS/, "", campo_art)
    }
    if (campo_art != "") print campo_art
}
' | sort -u > "$articulos_tmp"

total_articulos=$(wc -l < "$articulos_tmp")
echo "articulos unicos extraidos: $total_articulos"

if [ ! -s "$articulos_tmp" ]; then
    cp "$archivo_filtrado" "$archivo_salida"
    rm -f "$archivo_filtrado" "$articulos_tmp"
    echo "sin articulos para consultar tarifas"
    mv "$archivo_salida" "$archivo_final"
    echo "archivo final generado: $archivo_final"
    exit 0
fi

###############################################################################
# paso 3: generacion del script sql para tarifas
###############################################################################

echo "=== PASO 3: Consulta de tarifas en Informix ==="

{
    echo "database gen;"
    echo "create temp table tmp_int_art (int_art integer) with no log;"
    while read -r art; do
        echo "insert into tmp_int_art values ($art);"
    done < "$articulos_tmp"
    echo "select t.int_art, t.pvp_tar"
    echo "  from tarifas t, tmp_int_art a"
    echo " where t.cod_emp = 1"
    echo "   and t.cod_tar = 1"
    echo "   and t.int_art = a.int_art;"
} > "$sql_tmp"

dbaccess gen "$sql_tmp" 2>/dev/null | \
awk '
    /^[[:space:]]*[0-9]+[[:space:]]+[0-9]+\.[0-9][0-9][[:space:]]*$/ {
        gsub(/^[[:space:]]+|[[:space:]]+$/, "", $0)
        split($0, cols, /[[:space:]]+/)
        print cols[1] "," cols[2]
    }
' > "$tarifas_tmp"

total_tarifas=$(wc -l < "$tarifas_tmp")
echo "tarifas obtenidas: $total_tarifas"

if [ ! -s "$tarifas_tmp" ]; then
    cp "$archivo_filtrado" "$archivo_salida"
    rm -f "$archivo_filtrado" "$articulos_tmp" "$tarifas_tmp" "$sql_tmp"
    echo "sin tarifas devueltas por dbaccess"
    mv "$archivo_salida" "$archivo_final"
    echo "archivo final generado: $archivo_final"
    exit 0
fi

###############################################################################
# paso 4: aplicacion de tarifas al archivo filtrado
###############################################################################

echo "=== PASO 4: Aplicacion de tarifas ==="

head -n 2 "$archivo_filtrado" > "$archivo_salida"

awk -F',' -v OFS=',' '
BEGIN {
    while ((getline linea < ARGV[2]) > 0) {
        split(linea, t, ",")
        int_art = t[1]
        pvp = t[2]
        gsub(/^[[:space:]]+|[[:space:]]+$/, "", int_art)
        gsub(/^[[:space:]]+|[[:space:]]+$/, "", pvp)
        tarifas[int_art] = pvp
    }
    close(ARGV[2])
    ARGV[2] = ""
}
NR <= 2 { next }
{
    campo_art = $4
    gsub(/^"|"$/, "", campo_art)
    if (campo_art ~ /^SRS[0-9]+$/) {
        gsub(/^SRS/, "", campo_art)
    }
    int_art = campo_art
    if (int_art in tarifas) {
        pvp = tarifas[int_art]
        gsub(/^[[:space:]]+|[[:space:]]+$/, "", pvp)
        sub(/\..*$/, "", pvp)
        $17 = pvp
    }
    print
}
' "$archivo_filtrado" "$tarifas_tmp" >> "$archivo_salida"

echo "archivo procesado con tarifas: $archivo_salida"

###############################################################################
# paso 5: extraccion de sales checks unicos
###############################################################################

echo "=== PASO 5: Extraccion de SALES CHECKs unicos ==="

awk -F',' 'NR > 2 && $15 != "" {
    gsub(/^"|"$/, "", $15)
    sales = $15
    # agregar ceros a la izquierda hasta completar 16 caracteres
    while (length(sales) < 16) {
        sales = "0" sales
    }
    print sales
}' "$archivo_salida" | sort -u > "$sales_checks_tmp"

total_sales=$(wc -l < "$sales_checks_tmp")
echo "sales checks unicos: $total_sales"

###############################################################################
# paso 6: validacion bts
###############################################################################

echo "=== PASO 6: Validacion BTS ==="

# guardar directorio actual
dir_actual=$(pwd)
echo "[DEBUG] Directorio actual: $dir_actual"

# obtener rutas absolutas de los archivos
sales_checks_abs="$dir_actual/$sales_checks_tmp"
bts_salida_abs="$dir_actual/$bts_salida"

echo "[DEBUG] Archivo sales checks (absoluto): $sales_checks_abs"
echo "[DEBUG] Archivo salida BTS (absoluto): $bts_salida_abs"
echo "[DEBUG] Primeros 3 SALES CHECKs a consultar:"
head -3 "$sales_checks_abs"

# cambiar a directorio de bts
cd /gnx_prod/manto/desa/trabajo/sears/mangel/integridad || exit 4
echo "[DEBUG] Cambiado a directorio BTS: $(pwd)"

# ejecutar bts_batch con rutas absolutas
echo "[DEBUG] Ejecutando: fglgo bts_bat $sales_checks_abs $bts_salida_abs"
fglgo bts_bat "$sales_checks_abs" "$bts_salida_abs"

if [ ! -f "$bts_salida_abs" ]; then
    echo "[ERROR] No se genero el archivo BTS en: $bts_salida_abs"
    cd "$dir_actual"
    exit 5
fi

total_ns_bts=$(wc -l < "$bts_salida_abs")
echo "lineas NS encontradas en BTS: $total_ns_bts"

if [ "$total_ns_bts" -gt 0 ]; then
    echo "[DEBUG] Primeras 5 lineas del archivo BTS:"
    head -5 "$bts_salida_abs"
else
    echo "[DEBUG] El archivo BTS esta vacio"
    echo "[DEBUG] Verificando que el archivo exista:"
    ls -l "$bts_salida_abs"
fi

# volver al directorio original
cd "$dir_actual"
echo "[DEBUG] Volviendo a directorio: $(pwd)"

# si BTS no encontro lineas NS, copiar procesado a final sin cambios
if [ "$total_ns_bts" -eq 0 ]; then
    echo ""
    echo "WARNING: BTS no encontro lineas NS"
    echo "Esto puede deberse a:"
    echo "  1. Los SALES CHECKs no existen en la base de datos"
    echo "  2. No hay articulos NS asociados a estos SALES CHECKs en BTS"
    echo ""
    echo "Verificando primeros 5 SALES CHECKs enviados a BTS:"
    head -5 "$sales_checks_tmp"
    echo ""
    echo "Copiando archivo procesado sin cambios a archivo final..."
    cp "$archivo_salida" "$archivo_final"
    rm -f "$archivo_filtrado" "$articulos_tmp" "$tarifas_tmp" "$sql_tmp" "$sales_checks_tmp" "$bts_salida"
    echo "Archivo final generado: $archivo_final"
    exit 0
fi

###############################################################################
# paso 7: ajuste de lineas NS segun BTS
###############################################################################

echo "=== PASO 7: Ajuste de lineas NS ==="
echo "[DEBUG] Generando mapa de BTS..."
# crear mapa de BTS: cuantas veces aparece cada combinacion sales+int_art
awk -F'|' '{
    sales = $1
    gsub(/^[[:space:]]+|[[:space:]]+$/, "", sales)
    int_art = $2
    gsub(/^[[:space:]]+|[[:space:]]+$/, "", int_art)
    bts_count[sales,int_art]++
}
END {
    for (key in bts_count) {
        split(key, parts, SUBSEP)
        print parts[1] "," parts[2], bts_count[key]
    }
}' "$bts_salida" > "${archivo_salida}.bts_map"

echo "[DEBUG] Primeras 5 lineas del mapa BTS:"
head -5 "${archivo_salida}.bts_map"

# crear mapa del CSV procesado: cuantas veces aparece cada combinacion
awk -F',' 'NR > 2 {
    sales = $15
    gsub(/^"|"$/, "", sales)
    # agregar padding de ceros para que coincida con BTS
    while (length(sales) < 16) {
        sales = "0" sales
    }
    detalle_ns = $8
    gsub(/^"|"$/, "", detalle_ns)

    if (detalle_ns ~ /^NS/) {
        # es linea NS
        int_art = $7
        gsub(/^"|"$/, "", int_art)
        if (int_art ~ /^SRS[0-9]+$/) {
            gsub(/^SRS/, "", int_art)
        } else {
            int_art = $4
            gsub(/^"|"$/, "", int_art)
            if (int_art ~ /^SRS[0-9]+$/) {
                gsub(/^SRS/, "", int_art)
            }
        }

        key = sales "," int_art
        if (key in bts_map) {
            csv_count[key]++
            if (csv_count[key] <= bts_map[key]) {
                print
            }
        }
    } else {
        # no es linea NS: imprimir siempre
        print
    }
}
' "$archivo_salida" > "${archivo_salida}.csv_map"

echo "[DEBUG] Primeras 5 lineas del mapa CSV:"
head -5 "${archivo_salida}.csv_map"

# comparar mapas y generar lista de faltantes
echo "[DEBUG] Comparando mapas para encontrar faltantes..."

# comparar mapas y generar lista de faltantes
awk -v csv_map_file="${archivo_salida}.csv_map" '
BEGIN {
    # cargar mapa csv - usar split porque el formato es "clave contador" separado por espacio
    while ((getline linea < csv_map_file) > 0) {
        split(linea, partes, / /)
        key = partes[1]
        count = partes[2]
        csv_map[key] = count
    }
    close(csv_map_file)
}
{
    # leer mapa bts - mismo formato
    split($0, partes, / /)
    key = partes[1]
    bts_count = partes[2]
    csv_count = csv_map[key] + 0

    if (bts_count > csv_count) {
        # faltan lineas: necesitamos agregar (bts_count - csv_count) lineas
        split(key, parts, ",")
        sales = parts[1]
        int_art = parts[2]
        needed = bts_count - csv_count
        print sales, int_art, needed
    }
}' "${archivo_salida}.bts_map" > "${archivo_salida}.faltantes"

if [ -s "${archivo_salida}.faltantes" ]; then
    echo "[DEBUG] Se encontraron lineas faltantes:"
    cat "${archivo_salida}.faltantes"
else
    echo "[DEBUG] No hay lineas faltantes"
fi

echo "[DEBUG] Procesando CSV para eliminar sobrantes..."

# copiar cabeceras
head -n 2 "$archivo_salida" > "$archivo_final"

# procesar CSV: eliminar sobrantes y mantener las correctas
awk -F',' -v OFS=',' -v bts_map_file="${archivo_salida}.bts_map" '
BEGIN {
    # cargar mapa bts - usar split porque el formato es "clave contador" separado por espacio
    while ((getline linea < bts_map_file) > 0) {
        split(linea, partes, / /)
        key = partes[1]
        count = partes[2]
        bts_map[key] = count
    }
    close(bts_map_file)
}
NR <= 2 { next } # Saltar cabeceras
{
    sales = $15
    gsub(/^"|"$/, "", sales)
    # agregar padding de ceros para que coincida con BTS
    while (length(sales) < 16) {
        sales = "0" sales
    }
    detalle_ns = $8
    gsub(/^"|"$/, "", detalle_ns)

    if (detalle_ns ~ /^NS/) {
        # es linea NS
        int_art = $7
        gsub(/^"|"$/, "", int_art)
        if (int_art ~ /^SRS[0-9]+$/) {
            gsub(/^SRS/, "", int_art)
        } else {
            int_art = $4
            gsub(/^"|"$/, "", int_art)
            if (int_art ~ /^SRS[0-9]+$/) {
                gsub(/^SRS/, "", int_art)
            }
        }

        key = sales "," int_art
        if (key in bts_map) {
            csv_count[key]++
            if (csv_count[key] <= bts_map[key]) {
                print
            }
        }
    }
}
END {
    # Eliminar líneas sobrantes
    for (key in csv_count) {
        if (csv_count[key] > bts_map[key]) {
            csv_count[key] = bts_map[key]
        }
    }
}' "$archivo_salida" > "$archivo_final"

# agregar lineas faltantes desde el archivo original
if [ -s "${archivo_salida}.faltantes" ]; then
    echo "agregando lineas faltantes desde archivo original..."

    while read -r sales_faltante int_art_faltante cantidad_faltante; do
        echo "  buscando $cantidad_faltante lineas de sales=$sales_faltante int_art=$int_art_faltante"

        # quitar ceros iniciales para buscar en CSV original (que no tiene padding)
        sales_sin_padding=$(echo "$sales_faltante" | sed 's/^0*//')

        # buscar en archivo original y guardar en temporal para evitar problema con stdin
        grep "$sales_sin_padding" "$archivo" > "${archivo_salida}.grep_tmp"

        # procesar archivo temporal y agregar las necesarias
        awk -F',' -v sales="$sales_faltante" -v int_art="$int_art_faltante" -v needed="$cantidad_faltante" '
        BEGIN { count = 0 }
        {
            detalle_ns = $8
            gsub(/^"|"$/, "", detalle_ns)

            if (detalle_ns ~ /^NS/ && count < needed) {
                # verificar si el int_art coincide
                check_int_art = $7
                gsub(/^"|"$/, "", check_int_art)
                if (check_int_art ~ /^SRS[0-9]+$/) {
                    gsub(/^SRS/, "", check_int_art)
                }

                if (check_int_art == int_art) {
                    print
                    count++
                }
            }
        }
        ' "${archivo_salida}.grep_tmp" >> "$archivo_final"

    done < "${archivo_salida}.faltantes"

    # limpiar temporal
    rm -f "${archivo_salida}.grep_tmp"
fi

# limpiar archivos temporales de mapas
# COMENTADO TEMPORALMENTE PARA DEBUG
# rm -f "${archivo_salida}.bts_map" "${archivo_salida}.csv_map" "${archivo_salida}.faltantes"

total_lineas_final=$(tail -n +3 "$archivo_final" | wc -l)
total_lineas_procesado=$(tail -n +3 "$archivo_salida" | wc -l)

echo "lineas en procesado: $total_lineas_procesado"
echo "lineas en final (ajustado): $total_lineas_final"

if [ $total_lineas_final -lt $total_lineas_procesado ]; then
    eliminadas=$((total_lineas_procesado - total_lineas_final))
    echo "lineas NS eliminadas (sobrantes): $eliminadas"
elif [ $total_lineas_final -gt $total_lineas_procesado ]; then
    agregadas=$((total_lineas_final - total_lineas_procesado))
    echo "lineas NS agregadas (faltantes): $agregadas"
else
    echo "sin cambios: todas las lineas NS coinciden con BTS"
fi

###############################################################################
# paso 8: limpieza de archivos temporales
###############################################################################

echo "=== PASO 8: Limpieza de archivos temporales ==="

rm -f "$archivo_filtrado" "$articulos_tmp" "$tarifas_tmp" "$sql_tmp" "$sales_checks_tmp" "$bts_salida"
rm -f "${archivo_salida}.bts_map" "${archivo_salida}.csv_map" "${archivo_salida}.faltantes"
rm -f "$archivo_salida"

echo ""
echo "==================================================="
echo "PROCESO COMPLETADO"
echo "==================================================="
echo "Archivo final generado: $archivo_final"
echo ""