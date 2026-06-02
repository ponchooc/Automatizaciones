#!/usr/bin/env bash

# --- Configuración ---
ENTRADA="alimenta.txt"
DATABASE="gen"
TEMP_SQL="temp_query.sql"
TEMP_OUT="temp_salida.txt"

# 1. LIMPIEZA INICIAL
# Borramos archivos de salida previos para que si un número de repeticiones 
# ya no existe en el archivo nuevo, no se queden datos viejos.
rm -f de*.txt 2>/dev/null

# Verificar si existe el archivo de entrada
if [ ! -f "$ENTRADA" ]; then
    echo "Error: El archivo $ENTRADA no existe en esta ruta."
    exit 1
fi

echo "Procesando registros..."

# 2. LEER alimenta.txt (Columna 1: tienda, 2: documento, 3: veces)
while read -r tienda documento veces; do
    # Saltar líneas vacías
    [ -z "$tienda" ] && continue

    # Crear el SQL para la línea actual
    # Usamos la tienda del registro ($tienda) para el filtro y la última columna del SELECT
    cat <<EOF > "$TEMP_SQL"
UNLOAD TO "$TEMP_OUT"
SELECT B.cod_emp, TODAY, B.num_orp, B.cod_pto, B.int_art, ABS(B.uni_mov), $tienda
FROM orp_cab A
INNER JOIN orp_det B 
        ON B.cod_emp = A.cod_emp 
       AND B.cod_pto = A.cod_pto 
       AND B.num_orp = A.num_orp
WHERE A.cod_emp = 1 
  AND A.cod_pto = $tienda 
  AND A.num_orp = $documento;
EOF

    # Ejecutar el query en Informix (silencioso)
    dbaccess "$DATABASE" "$TEMP_SQL" > /dev/null 2>&1

    # 3. PROCESAR RESULTADO
    if [ -f "$TEMP_OUT" ] && [ -s "$TEMP_OUT" ]; then
        
        # Leemos el archivo temporal generado por Informix línea por línea
        while read -r linea_raw; do
            # Transformación: 
            # - s/\.0//g  -> quita todos los ".0"
            # - s/\///g   -> quita todas las "/" de la fecha
            linea_final=$(echo "$linea_raw" | sed 's/\.0//g; s/\///g')

            # Nombre del archivo basado en la cantidad (de1.txt, de2.txt, etc.)
            ARCHIVO_DESTINO="de${veces}.txt"

            # 4. REPETIR la línea N veces según la columna 'veces'
            for ((i=1; i<=veces; i++)); do
                echo "$linea_final" >> "$ARCHIVO_DESTINO"
            done
        done < "$TEMP_OUT"
        
        # Borrar el temporal de salida para que no se use en la siguiente vuelta
        rm -f "$TEMP_OUT"
    fi

done < "$ENTRADA"

# 5. LIMPIEZA FINAL DE ARCHIVOS DE TRABAJO
rm -f "$TEMP_SQL" "$TEMP_OUT"

echo "Proceso terminado exitosamente."
echo "Archivos generados:"
ls de*.txt 2>/dev/null
