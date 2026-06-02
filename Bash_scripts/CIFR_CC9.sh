#!/usr/bin/bash

###############################################################################
# CIFR_CC2.sh
# Reporte de cifras Genesix vs WMS por correo electrónico en HTML
# + Generación y adjunto de archivos detallados por transferencia
# Autor: Carlos Alfonso Ortega Molina
###############################################################################

# Mostrar SQL generado si debug_sql=1
debug_sql=0

###############################################################################
# FUNCIONES DE UTILIDAD
###############################################################################

# Valida que la fecha tenga formato DDMMAAAA
validar_fecha() {
    [[ $1 =~ ^[0-9]{8}$ ]]
}

# Convierte DDMMAAAA a YYYYMMDD para comparar fechas
to_yyyymmdd() {
    echo "$1" | awk '{print substr($0,5,4) substr($0,3,2) substr($0,1,2)}'
}

# Devuelve el nombre del mes en español y mayúsculas a partir de MM
nombre_mes() {
    case "$1" in
        "01") echo "ENERO" ;;
        "02") echo "FEBRERO" ;;
        "03") echo "MARZO" ;;
        "04") echo "ABRIL" ;;
        "05") echo "MAYO" ;;
        "06") echo "JUNIO" ;;
        "07") echo "JULIO" ;;
        "08") echo "AGOSTO" ;;
        "09") echo "SEPTIEMBRE" ;;
        "10") echo "OCTUBRE" ;;
        "11") echo "NOVIEMBRE" ;;
        "12") echo "DICIEMBRE" ;;
        *) echo "" ;;
    esac
}

# Extrae el valor numérico después de una etiqueta en la salida de dbaccess
extraer_valor() {
    # $1 es la etiqueta a buscar (wms_total o genesix_total)
    awk -v tag="$1" '
        found && /^[[:space:]]*[0-9]/ {print $1; exit}
        $0 ~ tag {found=1}
    '
}

# Obtiene la fecha de ayer en formato DDMMAAAA compatible con AIX
ayer_ddmmaaaa() {
    perl -MPOSIX -le 'print strftime("%d%m%Y", localtime(time()-86400))'
}

###############################################################################
# CONSTRUCCIÓN DE RANGO DE FECHAS Y MESES
###############################################################################

construir_label_fecha() {
    if [ $# -eq 0 ]; then
        # Solo cifras de AYER si no hay parámetros
        fecha_inicial=$(ayer_ddmmaaaa)
        fecha_final=$fecha_inicial
        fecha_label="AYER"
        meses=( "AYER" )
    elif [ $# -eq 2 ]; then
        fecha_inicial="$1"
        fecha_final="$2"
        if ! validar_fecha "$fecha_inicial" || ! validar_fecha "$fecha_final"; then
            echo "error: las fechas deben tener formato ddmmaaaa y solo contener numeros."
            exit 1
        fi
        fecha_ini_num=$(to_yyyymmdd "$fecha_inicial")
        fecha_fin_num=$(to_yyyymmdd "$fecha_final")
        if [ "$fecha_ini_num" -gt "$fecha_fin_num" ]; then
            echo "error: la fecha inicial debe ser menor o igual a la fecha final."
            exit 1
        fi
        mes_ini="${fecha_inicial:2:2}"
        mes_fin="${fecha_final:2:2}"
        anio_ini="${fecha_inicial:4:4}"
        anio_fin="${fecha_final:4:4}"
        if [ "$mes_ini" = "$mes_fin" ] && [ "$anio_ini" = "$anio_fin" ]; then
            fecha_label="$(nombre_mes "$mes_ini")"
            meses=( "$fecha_label" )
        else
            # Generar lista de meses entre las fechas dadas
            meses=()
            ini="$anio_ini$mes_ini"
            fin="$anio_fin$mes_fin"
            y=${anio_ini}
            m=${mes_ini}
            while : ; do
                meses+=( "$(nombre_mes $(printf "%02d" $m))" )
                if [ "$y$(printf "%02d" $m)" = "$fin" ]; then break; fi
                m=$((10#$m + 1))
                if [ $m -gt 12 ]; then m=1; y=$((y+1)); fi
            done
            fecha_label="${fecha_inicial}-${fecha_final}"
        fi
    else
        echo "uso: $0 [fecha_inicial(ddmmaaaa) fecha_final(ddmmaaaa)]"
        exit 1
    fi
}

###############################################################################
# BLOQUE PRINCIPAL: EJECUCIÓN Y ARMADO DE RESULTADOS
###############################################################################

construir_label_fecha "$@"

# Determinar rango de fechas para los queries detallados y totales
if [ $# -eq 0 ]; then
    FECHA1=$(ayer_ddmmaaaa)
    FECHA2=$FECHA1
else
    FECHA1="$fecha_inicial"
    FECHA2="$fecha_final"
fi

# Usar el mismo rango para totales y detalle
fecha_oc="'$FECHA1' and '$FECHA2'"
fecha_bt="'$FECHA1' and '$FECHA2'"
fecha_tr="'$FECHA1' and '$FECHA2'"

# Arreglos para guardar resultados por mes
declare -A tras_genesix tras_wms tras_dif
declare -A predis_genesix predis_wms predis_dif
declare -A bt_genesix bt_wms bt_dif

###############################################################################
# EJECUCIÓN DE QUERIES Y EXTRACCIÓN DE CIFRAS
###############################################################################

ejecutar_query() {
    # $1 es el tipo: tras, predis, bt
    # $2 es el mes en formato mm
    local tipo="$1"
    local sqltmp="tmp_${tipo}_$$.sql"
    local paso=""
    local filtro_mes=""
    # Filtro de mes corregido para fechas DD/MM/YYYY
    if [ "$2" != "" ] && [ "$2" != "ANTIER-AYER" ] && [ "$2" != "AYER" ]; then
        filtro_mes="and substr(fec_emb,4,2) = '$2'"
    fi
    case "$tipo" in
        predis)
            paso="paso_oc"
            cat > "$sqltmp" <<EOF
select a.fec_emb,a.cod_pto,a.pto_tra,a.con_emb,a.ped_com,emi_com,sum(num_pzas) num_pzas
  from ora_oc2_emb a
where cod_pto = 870
  and fec_emb between $fecha_oc
  $filtro_mes
group by 1,2,3,4,5,6
  into temp $paso with no log;

select sum(num_pzas) as wms_total from $paso;

select sum(uni_rep) as genesix_total
from $paso a, tra_cab b, tra_det c, tra_env d
where b.cod_emp=1
  and b.pto_emi=999
  and b.cod_pto=a.cod_pto
  and b.pto_tra=a.pto_tra
  and b.emi_com=a.emi_com
  and b.ped_com=a.ped_com
  and c.cod_emp=b.cod_emp
  and c.pto_emi=b.pto_emi
  and c.num_rep=b.num_rep
  and c.cod_pto=b.cod_pto
  and c.pto_tra=b.pto_tra
  and d.cod_emp=c.cod_emp
  and d.pto_emi=c.pto_emi
  and d.num_rep=c.num_rep
  and d.cod_pto=c.cod_pto
  and d.pto_tra=c.pto_tra
  and d.con_emb=c.con_emb;

drop table $paso;
EOF
            ;;
        bt)
            paso="paso_bt"
            cat > "$sqltmp" <<EOF
select a.fec_emb,a.cod_pto,a.con_emb,a.num_scn,sum(num_pzas) num_pzas
  from ora_bt2_emb a
where a.cod_pto = 870
  and a.fec_emb between $fecha_bt
  $filtro_mes
group by 1,2,3,4
  into temp $paso;

select sum(num_pzas) as wms_total from $paso;

select sum(uni_rep) as genesix_total
  from $paso a, outer (tra_cab b, tra_det c)
where b.cod_emp = 1
   and b.num_scn = a.num_scn
   and b.bt_tra  = 'S'
   and c.cod_emp = b.cod_emp
   and c.pto_emi = b.pto_emi
   and c.num_rep = b.num_rep
   and c.cod_pto = b.cod_pto
   and c.pto_tra = b.pto_tra
   and c.con_emb[4,10] = a.con_emb;

drop table $paso;
EOF
            ;;
        tras)
            paso="paso_tr"
            cat > "$sqltmp" <<EOF
select a.fec_emb,a.cod_pto,a.pto_tra,a.con_emb,a.num_rep,pto_emi, sum(num_pzas) num_pzas
  from ora_tr2_emb a
where cod_pto = 870
  and fec_emb between $fecha_tr
  $filtro_mes
 group by 1,2,3,4,5,6
  into temp $paso;

select sum(num_pzas) as wms_total from $paso;

select sum(uni_rep) as genesix_total
  from $paso a, outer (tra_cab b, tra_det c)
where b.cod_emp = 1
   and b.pto_emi = a.pto_emi
   and b.num_rep = a.num_rep
   and b.cod_pto = a.cod_pto
   and b.pto_tra = a.pto_tra
   and c.cod_emp = b.cod_emp
   and c.pto_emi = b.pto_emi
   and c.num_rep = b.num_rep
   and c.cod_pto = b.cod_pto
   and c.pto_tra = b.pto_tra
   and c.con_emb[4,10] = a.con_emb;

drop table $paso;
EOF
            ;;
    esac

    if [ "$debug_sql" = "1" ]; then
        echo "----- sql para $tipo -----"
        cat "$sqltmp"
        echo "-------------------------"
    fi

    local output=$(dbaccess gen "$sqltmp" 2>/dev/null)
    rm -f "$sqltmp"

    local wms=$(echo "$output" | extraer_valor wms_total)
    local genesix=$(echo "$output" | extraer_valor genesix_total)

    wms=${wms:-0}
    genesix=${genesix:-0}

    echo "$genesix $wms"
}

for mes in "${meses[@]}"; do
    # Obtener MM para el filtro
    if [ "$mes" = "ANTIER-AYER" ] || [ "$mes" = "AYER" ]; then
        mm=""
    else
        case "$mes" in
            ENERO) mm="01" ;;
            FEBRERO) mm="02" ;;
            MARZO) mm="03" ;;
            ABRIL) mm="04" ;;
            MAYO) mm="05" ;;
            JUNIO) mm="06" ;;
            JULIO) mm="07" ;;
            AGOSTO) mm="08" ;;
            SEPTIEMBRE) mm="09" ;;
            OCTUBRE) mm="10" ;;
            NOVIEMBRE) mm="11" ;;
            DICIEMBRE) mm="12" ;;
        esac
    fi

    # Ejecutar queries para cada tipo y mes
    read tg tw <<< $(ejecutar_query "tras" "$mm")
    tras_genesix[$mes]=$tg
    tras_wms[$mes]=$tw
    tras_dif[$mes]=$(echo "${tg:-0} - ${tw:-0}" | bc)

    read pg pw <<< $(ejecutar_query "predis" "$mm")
    predis_genesix[$mes]=$pg
    predis_wms[$mes]=$pw
    predis_dif[$mes]=$(echo "${pg:-0} - ${pw:-0}" | bc)

    read bg bw <<< $(ejecutar_query "bt" "$mm")
    bt_genesix[$mes]=$bg
    bt_wms[$mes]=$bw
    bt_dif[$mes]=$(echo "${bg:-0} - ${bw:-0}" | bc)
done

# Leer destinatarios del archivo
destinatarios=$(paste -sd, NO_BORRAR_destinatarios2.txt)

###############################################################################
# GENERACIÓN DE ARCHIVOS DETALLADOS Y ADJUNTOS
###############################################################################

# Obtener fecha y hora para los nombres de archivo (AIX compatible)
DD=$(date +%d)
MM=$(date +%m)
HH=$(date +%H)
mm_=$(date +%M)

# ------------------ DETALLE TRANSFERENCIAS STOCK ------------------
cat > det_trans_stock_${DD}${MM}_${HH}${mm_}.sql <<EOF
select con_emb, num_rep, sum(num_pzas) as total_genesix
  from ora_tr2_emb
 where cod_pto = 870
   and fec_emb between '$FECHA1' and '$FECHA2'
 group by 1,2
  into temp genesix_simple with no log;

select a.con_emb, a.num_rep, sum(uni_rep) as total_wms
  from ora_tr2_emb a, outer (tra_cab b, tra_det c)
 where a.cod_pto = 870
   and a.fec_emb between '$FECHA1' and '$FECHA2'
   and b.cod_emp = 1
   and b.pto_emi = a.pto_emi
   and b.num_rep = a.num_rep
   and b.cod_pto = a.cod_pto
   and b.pto_tra = a.pto_tra
   and c.cod_emp = b.cod_emp
   and c.pto_emi = b.pto_emi
   and c.num_rep = b.num_rep
   and c.cod_pto = b.cod_pto
   and c.pto_tra = b.pto_tra
   and c.con_emb[4,10] = a.con_emb
 group by 1,2
  into temp wms_simple with no log;

select
    coalesce(g.con_emb, w.con_emb) as con_emb,
    coalesce(g.num_rep, w.num_rep) as num_rep,
    g.total_genesix,
    w.total_wms,
    (nvl(g.total_genesix,0) - nvl(w.total_wms,0)) as diferencia,
    case
        when g.con_emb is null then 'Solo en WMS'
        when w.con_emb is null then 'Solo en Genesix'
        when g.total_genesix != w.total_wms then 'Cantidad diferente'
    end as tipo_diferencia
  from genesix_simple g
  full outer join wms_simple w
    on g.con_emb = w.con_emb
   and g.num_rep = w.num_rep
 where not (
        g.total_genesix = w.total_wms
    and g.con_emb is not null
    and w.con_emb is not null
);
EOF

dbaccess gen det_trans_stock_${DD}${MM}_${HH}${mm_}.sql 2>/dev/null | awk 'BEGIN{OFS="|"} /^[^ ]/ && !/^$/ {gsub(/[[:space:]]+/, "|"); print}' > det_trans_stock_${DD}${MM}_${HH}${mm_}.txt
rm -f det_trans_stock_${DD}${MM}_${HH}${mm_}.sql

# ------------------ DETALLE BT ------------------
cat > det_bt_${DD}${MM}_${HH}${mm_}.sql <<EOF
select a.fec_emb, a.cod_pto, a.con_emb, a.num_scn, sum(num_pzas) as num_pzas
  from ora_bt2_emb a
 where a.cod_pto = 870
   and a.fec_emb between '$FECHA1' and '$FECHA2'
 group by 1,2,3,4
  into temp paso_bt with no log;

select con_emb, num_scn, sum(num_pzas) as total_genesix
  from paso_bt
 group by 1,2
  into temp genesix_bt with no log;

select a.con_emb, a.num_scn, sum(uni_rep) as total_wms
  from paso_bt a, outer (tra_cab b, tra_det c)
 where b.cod_emp = 1
   and b.num_scn = a.num_scn
   and b.bt_tra  = 'S'
   and c.cod_emp = b.cod_emp
   and c.pto_emi = b.pto_emi
   and c.num_rep = b.num_rep
   and c.cod_pto = b.cod_pto
   and c.pto_tra = b.pto_tra
   and c.con_emb[4,10] = a.con_emb
 group by 1,2
  into temp wms_bt with no log;

select
    coalesce(g.con_emb, w.con_emb) as con_emb,
    coalesce(g.num_scn, w.num_scn) as num_scn,
    g.total_genesix,
    w.total_wms,
    (nvl(g.total_genesix,0) - nvl(w.total_wms,0)) as diferencia,
    case
        when g.con_emb is null then 'Solo en WMS'
        when w.con_emb is null then 'Solo en Genesix'
        when g.total_genesix != w.total_wms then 'Cantidad diferente'
    end as tipo_diferencia
  from genesix_bt g
  full outer join wms_bt w
    on g.con_emb = w.con_emb
   and g.num_scn = w.num_scn
 where not (
        g.total_genesix = w.total_wms
    and g.con_emb is not null
    and w.con_emb is not null
);
EOF

dbaccess gen det_bt_${DD}${MM}_${HH}${mm_}.sql 2>/dev/null | awk 'BEGIN{OFS="|"} /^[^ ]/ && !/^$/ {gsub(/[[:space:]]+/, "|"); print}' > det_bt_${DD}${MM}_${HH}${mm_}.txt
rm -f det_bt_${DD}${MM}_${HH}${mm_}.sql

# ------------------ DETALLE TRANSFERENCIAS PREDISTRIBUIDAS ------------------
cat > det_trans_predis_${DD}${MM}_${HH}${mm_}.sql <<EOF
SELECT
    COALESCE(g.con_emb, w.con_emb) AS con_emb,
    COALESCE(g.ped_com, w.ped_com) AS ped_com,
    COALESCE(g.emi_com, w.emi_com) AS emi_com,
    g.total_genesix,
    w.total_wms,
    NVL(g.total_genesix,0) - NVL(w.total_wms,0) AS diferencia,
    CASE
        WHEN g.con_emb IS NULL THEN 'Solo en WMS'
        WHEN w.con_emb IS NULL THEN 'Solo en Genesix'
        WHEN g.total_genesix != w.total_wms THEN 'Cantidad diferente'
    END AS tipo_diferencia
FROM
    (
        SELECT con_emb, ped_com, emi_com, SUM(num_pzas) AS total_genesix
        FROM ora_oc2_emb
        WHERE cod_pto = 870
          AND fec_emb BETWEEN '$FECHA1' AND '$FECHA2'
        GROUP BY con_emb, ped_com, emi_com
    ) g
FULL OUTER JOIN
    (
        SELECT a.con_emb, a.ped_com, a.emi_com, SUM(c.uni_rep) AS total_wms
        FROM ora_oc2_emb a
        INNER JOIN tra_cab b ON
            b.cod_emp = 1
            AND b.pto_emi = 999
            AND b.cod_pto = a.cod_pto
            AND b.pto_tra = a.pto_tra
            AND b.emi_com = a.emi_com
            AND b.ped_com = a.ped_com
        INNER JOIN tra_det c ON
            c.cod_emp = b.cod_emp
            AND c.pto_emi = b.pto_emi
            AND c.num_rep = b.num_rep
            AND c.cod_pto = b.cod_pto
            AND c.pto_tra = b.pto_tra
        INNER JOIN tra_env d ON
            d.cod_emp = c.cod_emp
            AND d.pto_emi = c.pto_emi
            AND d.num_rep = c.num_rep
            AND d.cod_pto = c.cod_pto
            AND d.pto_tra = c.pto_tra
            AND d.con_emb = c.con_emb
        WHERE a.cod_pto = 870
          AND a.fec_emb BETWEEN '$FECHA1' AND '$FECHA2'
        GROUP BY a.con_emb, a.ped_com, a.emi_com
    ) w
ON g.con_emb = w.con_emb
   AND g.ped_com = w.ped_com
   AND g.emi_com = w.emi_com
WHERE NOT (
    g.total_genesix = w.total_wms
    AND g.con_emb IS NOT NULL
    AND w.con_emb IS NOT NULL
);
EOF

dbaccess gen det_trans_predis_${DD}${MM}_${HH}${mm_}.sql 2>/dev/null | awk 'BEGIN{OFS="|"} /^[^ ]/ && !/^$/ {gsub(/[[:space:]]+/, "|"); print}' > det_trans_predis_${DD}${MM}_${HH}${mm_}.txt
rm -f det_trans_predis_${DD}${MM}_${HH}${mm_}.sql

# Guardar nombres de archivos para comprimir y adjuntar
ADJ1=det_trans_stock_${DD}${MM}_${HH}${mm_}.txt
ADJ2=det_bt_${DD}${MM}_${HH}${mm_}.txt
ADJ3=det_trans_predis_${DD}${MM}_${HH}${mm_}.txt

# Comprimir los tres archivos en uno solo con máxima compresión
ARCHCOMP="archvios_${DD}${MM}${HH}${mm_}.tar.gz"
tar -cvf - "$ADJ1" "$ADJ2" "$ADJ3" | gzip -9 > "$ARCHCOMP"

###############################################################################
# GENERACIÓN DEL REPORTE HTML
###############################################################################

generar_html() {
cat <<EOF
<html>
<head>
<style>
body { font-family: Arial, sans-serif; }
table { border-collapse: collapse; width: 100%; max-width: 600px; margin-bottom: 20px; }
th, td { border: 1px solid #0a4c6a; padding: 6px 10px; text-align: center; }
th { background: #0a4c6a; color: #fff; }
caption { background: #0a4c6a; color: #fff; font-weight: bold; padding: 6px; text-align: left; }
</style>
</head>
<body>
<table>
<caption>Transferencias STOCK</caption>
<tr>
    <th>Fecha_emb</th>
    <th>Genesix</th>
    <th>WMS</th>
    <th>Gnx Vs WMS</th>
</tr>
EOF
for mes in "${meses[@]}"; do
    printf "<tr><td>%s</td><td>%d</td><td>%d</td><td>%d</td></tr>\n" \
        "$mes" "${tras_genesix[$mes]%%.*}" "${tras_wms[$mes]%%.*}" "${tras_dif[$mes]%%.*}"
done
cat <<EOF
</table>
<table>
<caption>Transferencias Predistribuidas</caption>
<tr>
    <th>Fecha_emb</th>
    <th>Genesix</th>
    <th>WMS</th>
    <th>Gnx Vs WMS</th>
</tr>
EOF
for mes in "${meses[@]}"; do
    printf "<tr><td>%s</td><td>%d</td><td>%d</td><td>%d</td></tr>\n" \
        "$mes" "${predis_genesix[$mes]%%.*}" "${predis_wms[$mes]%%.*}" "${predis_dif[$mes]%%.*}"
done
cat <<EOF
</table>
<table>
<caption>BT</caption>
<tr>
    <th>Fecha_emb</th>
    <th>Genesix</th>
    <th>WMS</th>
    <th>Gnx Vs WMS</th>
</tr>
EOF
for mes in "${meses[@]}"; do
    printf "<tr><td>%s</td><td>%d</td><td>%d</td><td>%d</td></tr>\n" \
        "$mes" "${bt_genesix[$mes]%%.*}" "${bt_wms[$mes]%%.*}" "${bt_dif[$mes]%%.*}"
done
cat <<EOF
</table>
</body>
</html>
EOF
}

###############################################################################
# ENVÍO DE CORREO CON UN SOLO ARCHIVO COMPRIMIDO COMO ADJUNTO
###############################################################################

(
echo "From: desa@sears33.sanborns.net"
echo "To: $destinatarios"
echo "Subject: PROBANDO CIFRAS"
echo "MIME-Version: 1.0"
BOUNDARY="====MULTIPART_BOUNDARY_$(date +%s)===="
echo "Content-Type: multipart/mixed; boundary=\"$BOUNDARY\""
echo
echo "--$BOUNDARY"
echo "Content-Type: text/html; charset=UTF-8"
echo "Content-Disposition: inline"
echo
generar_html
echo "--$BOUNDARY"
echo "Content-Type: application/gzip; name=\"$ARCHCOMP\""
echo "Content-Transfer-Encoding: x-uuencode"
echo "Content-Disposition: attachment; filename=\"$ARCHCOMP\""
uuencode "$ARCHCOMP" "$ARCHCOMP"
echo "--$BOUNDARY--"
) | /usr/sbin/sendmail -t

###############################################################################
# FIN DEL SCRIPT
###############################################################################
