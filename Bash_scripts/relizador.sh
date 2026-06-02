#!/usr/bin/bash

set -euo pipefail

RUTA_BASE="/gnx_prod/manto/desa/trabajo/sears/carlos_ortega"
ARCHIVO_ENTRADA="${RUTA_BASE}/relices.txt"
ARCHIVO_SALIDA="${RUTA_BASE}/relices_procesado.txt"
BASE_DATOS="gen"

if [ ! -f "$ARCHIVO_ENTRADA" ]; then
  echo "Error: no existe el archivo de entrada: $ARCHIVO_ENTRADA" >&2
  exit 1
fi

cantidad=$(wc -l < "$ARCHIVO_ENTRADA" | tr -d '[:space:]')
if [ -z "$cantidad" ] || [ "$cantidad" -eq 0 ]; then
  echo "Error: el archivo de entrada esta vacio." >&2
  exit 1
fi

echo "Leyendo el archivo relices.txt con ${cantidad} lineas"
echo "Procesando"

: > "$ARCHIVO_SALIDA"

SQL_TEMP="${RUTA_BASE}/relices_sql.$$"
limpiar() {
  rm -f "$SQL_TEMP"
}
trap limpiar EXIT

{
  echo "CREATE TEMP TABLE tmp_relices (num_scn CHAR(16)) WITH NO LOG;"

  while IFS= read -r linea || [ -n "$linea" ]; do
    linea=${linea%$'\r'}
    if [ -z "$linea" ]; then
      continue
    fi
    if [[ ! "$linea" =~ ^[0-9]{16}$ ]]; then
      echo "Error: linea invalida en relices.txt: '$linea'" >&2
      exit 1
    fi
    echo "INSERT INTO tmp_relices (num_scn) VALUES ('$linea');"
  done < "$ARCHIVO_ENTRADA"

  echo "UNLOAD TO '${ARCHIVO_SALIDA}'"
  echo "SELECT A.num_scn,"
  echo "       CASE WHEN (B.intentos IS NULL)"
  echo "            THEN A.cod_pto||A.num_edc||'000'"
  echo "            ELSE A.cod_pto||A.num_edc||LPAD(TO_CHAR(B.intentos), 3, '0')"
  echo "       END AS ord_rel"
  echo "  FROM edc_cab A, OUTER ora_rt_envio B"
  echo " WHERE A.cod_emp = 1"
  echo "   AND A.num_scn IN (SELECT num_scn FROM tmp_relices)"
  echo "   AND B.num_scn = A.num_scn"
  echo "   AND B.cod_pto = A.cod_pto"
  echo "   AND B.num_edc = A.num_edc;"
} > "$SQL_TEMP"

SALIDA_DBACCESS="${RUTA_BASE}/relices_dbaccess.log.$$"
if dbaccess "$BASE_DATOS" - < "$SQL_TEMP" > "$SALIDA_DBACCESS" 2>&1; then
  :
else
  cat "$SALIDA_DBACCESS" >&2
  rm -f "$SALIDA_DBACCESS"
  exit 1
fi
rm -f "$SALIDA_DBACCESS"

echo "Proceso SQL termina cuando este termine"

cantidad_salida=$(wc -l < "$ARCHIVO_SALIDA" | tr -d '[:space:]')

: > "$ARCHIVO_ENTRADA"

echo "Programa finalizado correctamente con ${cantidad_salida} lineas en el archivo relices_procesado.txt"
