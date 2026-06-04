#!/usr/bin/bash
set -u

WORK_DIR="/respaldo_migracion/reportes_gnx/reportes/int74"
MAIL_FROM="sistemas_mercaderias@sears.com.mx"
MAIL_TO="gresendiz@sears.com.mx,ortegac@sanborns.com.mx mdominguezc@sears.com.mx,dmolina@sears.com.mx,eacastillo@sears.com.mx,balderramaj@sanborns.com.mx,rescobedog@sears.com.mx,kjecheverria@sears.com.mx,josanchez@sears.com.mx,facturacionst@sears.com.mx,mramirezv@sears.com.mx"
SUBJECT_NOT_FOUND="EL Archivo de no sports no existe"
SUBJECT_SUCCESS="Proceso NoSpot completado y depositado correctamente"

FTP_HOST="140.240.11.6"
FTP_USER="ftpusr01"
FTP_PASS="ftpgral1"
FTP_REMOTE_DIR="/syp_interfases/"
FTP_REMOTE_FILE="NoSpot187.txt"

LOG_FILE="${WORK_DIR}/nospot187_$(date '+%Y%m%d').log"
FTP_LAST_ERROR=""

log() {
  printf '%s | %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$*" | tee -a "$LOG_FILE"
}

send_mail() {
  subject="$1"
  body="$2"
  printf '%b\n' "$body" | mailx -r "$MAIL_FROM" -s "$subject" $MAIL_TO
}

build_expected_pattern() {
  local day_np month_np year4
  month_np=$((10#$(date '+%m')))
  day_np=$((10#$(date '+%d')))
  year4=$(date '+%Y')
  printf 'GPOSAN_VAL_OCS_Reportes Custom_INT74_%s_%s_%s_*.csv' "$month_np" "$day_np" "$year4"
}

find_daily_file() {
  local month_np day_np month_p day_p year4
  local -a files
  local f

  files=()
  month_np=$((10#$(date '+%m')))
  day_np=$((10#$(date '+%d')))
  month_p=$(date '+%m')
  day_p=$(date '+%d')
  year4=$(date '+%Y')

  while IFS= read -r f; do
    files+=("$f")
  done < <(find "$WORK_DIR" -type f \
    \( -name "GPOSAN_VAL_OCS_Reportes Custom_INT74_${month_np}_${day_np}_${year4}_*.csv" \
    -o -name "GPOSAN_VAL_OCS_Reportes Custom_INT74_${month_p}_${day_p}_${year4}_*.csv" \
    -o -name "GPOSAN_VAL_OCS_Reportes Custom_INT74_${month_np}_${day_p}_${year4}_*.csv" \
    -o -name "GPOSAN_VAL_OCS_Reportes Custom_INT74_${month_p}_${day_np}_${year4}_*.csv" \) -print)

  if [ "${#files[@]}" -eq 0 ]; then
    return 1
  fi

  ls -1t "${files[@]}" 2>/dev/null | head -n 1
}

process_csv() {
  src_file="$1"
  out_file="$2"

  : > "$out_file"

  awk -v out="$out_file" '
function trim(s) {
  sub(/^[ \t\r\n\"]+/, "", s)
  sub(/[ \t\r\n\"]+$/, "", s)
  return s
}

function csv_split(line, arr,   i,c,field,inq,n,nextc) {
  n = 0
  field = ""
  inq = 0

  for (i = 1; i <= length(line); i++) {
    c = substr(line, i, 1)

    if (c == "\"") {
      nextc = (i < length(line)) ? substr(line, i + 1, 1) : ""
      if (inq && nextc == "\"") {
        field = field "\""
        i++
      } else {
        inq = !inq
      }
    } else if (c == "," && !inq) {
      arr[++n] = field
      field = ""
    } else {
      field = field c
    }
  }

  arr[++n] = field
  return n
}

function split_name(full, outa,   t,n,i) {
  gsub(/^[ \t]+|[ \t]+$/, "", full)

  if (full == "") {
    outa[1] = "X"
    outa[2] = "X"
    outa[3] = "X"
    return
  }

  n = split(full, t, /[ ]+/)

  if (n == 1) {
    outa[1] = t[1]
    outa[2] = "X"
    outa[3] = "X"
    return
  }

  if (n == 2) {
    outa[1] = t[1]
    outa[2] = t[2]
    outa[3] = "X"
    return
  }

  outa[1] = t[1]
  for (i = 2; i <= n - 2; i++) {
    outa[1] = outa[1] " " t[i]
  }

  outa[2] = t[n - 1]
  outa[3] = t[n]
}

BEGIN {
  line_no = 0
  idx_sc = -1
  idx_ns = -1
  idx_nom = -1
  idx_cant = -1
  written = 0
}

{
  line_no++
  nf = csv_split($0, col)

  for (i = 1; i <= nf; i++) {
    col[i] = trim(col[i])
  }

  if (line_no == 2) {
    for (i = 1; i <= nf; i++) {
      if (col[i] == "Sales Check") idx_sc = i
      else if (col[i] == "No Spots") idx_ns = i
      else if (col[i] == "Nombre del Cliente") idx_nom = i
      else if (col[i] == "Cantidad") idx_cant = i
    }

    if (idx_sc < 0 || idx_ns < 0 || idx_nom < 0 || idx_cant < 0) {
      print "Faltan columnas requeridas en el CSV" > "/dev/stderr"
      exit 20
    }
    next
  }

  if (line_no < 3) next

  empty = 1
  for (i = 1; i <= nf; i++) {
    if (col[i] != "") {
      empty = 0
      break
    }
  }
  if (empty) next

  max_idx = idx_sc
  if (idx_ns > max_idx) max_idx = idx_ns
  if (idx_nom > max_idx) max_idx = idx_nom
  if (idx_cant > max_idx) max_idx = idx_cant

  if (nf < max_idx) next

  sales_check = col[idx_sc]
  no_spots = col[idx_ns]
  nombre_raw = col[idx_nom]
  cantidad = col[idx_cant]

  if (sales_check == "" && no_spots == "") next

  while (length(sales_check) < 16) {
    sales_check = "0" sales_check
  }

  if (toupper(substr(no_spots, 1, 3)) == "SRS") {
    no_spots = substr(no_spots, 4)
  }

  split_name(nombre_raw, parts)

  print sales_check "|" no_spots "|" parts[1] "|" parts[2] "|" parts[3] "|" cantidad "|" >> out
  written++
}

END {
  print written
}
' "$src_file"
}

ftp_upload() {
  local_file="$1"
  local_file_name=$(basename "$local_file")
  local ftp_output ftp_rc error_line

  FTP_LAST_ERROR=""

  ftp_output=$(ftp -v -n "$FTP_HOST" 2>&1 <<EOF
user $FTP_USER $FTP_PASS
binary
cd $FTP_REMOTE_DIR
put "$local_file_name" "$FTP_REMOTE_FILE"
bye
EOF
)
  ftp_rc=$?

  printf '%s\n' "----- INICIO TRANSCRIPCION FTP -----" >> "$LOG_FILE"
  printf '%s\n' "$ftp_output" >> "$LOG_FILE"
  printf '%s\n' "----- FIN TRANSCRIPCION FTP (rc=${ftp_rc}) -----" >> "$LOG_FILE"

  if [ "$ftp_rc" -ne 0 ]; then
    FTP_LAST_ERROR="El comando ftp termino con codigo ${ftp_rc}."
    return 1
  fi

  if printf '%s\n' "$ftp_output" | grep -qi "Transfer complete"; then
    return 0
  fi

  error_line=$(printf '%s\n' "$ftp_output" | grep -Ei '(^5[0-9][0-9]|not connected|unknown host|timed out|refused|login incorrect|permission denied|no such file|cannot|error)' | tail -n 1)
  if [ -n "$error_line" ]; then
    FTP_LAST_ERROR="$error_line"
  else
    FTP_LAST_ERROR="No se encontro la confirmacion 'Transfer complete' en la sesion FTP."
  fi

  return 1
}

main() {
  local expected_pattern daily_file out_file line_count now_text success_body
  expected_pattern=$(build_expected_pattern)

  if ! cd "$WORK_DIR"; then
    printf 'No se pudo entrar al directorio %s\n' "$WORK_DIR" >&2
    exit 1
  fi

  log "Inicio de ejecucion en $WORK_DIR"

  daily_file=$(find_daily_file || true)

  if [ -z "$daily_file" ]; then
    now_text=$(date '+%d/%m/%Y %H:%M:%S')
    if ! send_mail "$SUBJECT_NOT_FOUND" "el archivo ${expected_pattern} no se encuentra en la ruta ${WORK_DIR} aun, se reintenta en una hora. Fecha y hora: ${now_text}."; then
      log "Error enviando correo de no encontrado (intento 1)."
      exit 5
    fi
    log "Archivo no encontrado en intento 1. Se duerme 1 hora. Patron esperado: $expected_pattern"

    sleep 3600

    daily_file=$(find_daily_file || true)
    if [ -z "$daily_file" ]; then
      now_text=$(date '+%d/%m/%Y %H:%M:%S')
      if ! send_mail "$SUBJECT_NOT_FOUND" "se volvio a intentar el procesamiento del archivo ${expected_pattern} y no se encontro en la ruta ${WORK_DIR}. Fecha y hora: ${now_text}."; then
        log "Error enviando correo de no encontrado (intento 2)."
        exit 6
      fi
      log "Archivo no encontrado en intento 2. Fin con error."
      exit 2
    fi
  fi

  out_file="${WORK_DIR}/NoSpot187_$(date '+%d%m%Y').txt"
  log "Archivo diario seleccionado: $daily_file"
  log "Archivo de salida local: $out_file"

  if ! line_count=$(process_csv "$daily_file" "$out_file"); then
    now_text=$(date '+%d/%m/%Y %H:%M:%S')
    if ! send_mail "Falla en procesamiento NoSpot" "No se pudo procesar el archivo ${daily_file}. Fecha y hora: ${now_text}. Revisar log: ${LOG_FILE}."; then
      log "Error enviando correo por falla de procesamiento."
      exit 7
    fi
    log "Error en transformacion CSV."
    exit 3
  fi

  line_count=$(printf '%s\n' "$line_count" | tail -n 1 | tr -d '[:space:]')
  if [ -z "$line_count" ]; then
    line_count=0
  fi

  if ! ftp_upload "$out_file"; then
    now_text=$(date '+%d/%m/%Y %H:%M:%S')
    if ! send_mail "Falla en deposito FTP NoSpot" "Se proceso el archivo ${daily_file}, pero fallo el deposito FTP de ${out_file} hacia ${FTP_HOST}:${FTP_REMOTE_DIR}${FTP_REMOTE_FILE}.\nFecha y hora: ${now_text}.\nFalla detectada: ${FTP_LAST_ERROR}\nRevisar log: ${LOG_FILE}."; then
      log "Error enviando correo por falla FTP."
      exit 8
    fi
    log "Fallo el FTP. Revisar ${LOG_FILE}"
    exit 4
  fi

  now_text=$(date '+%d/%m/%Y %H:%M:%S')
  success_body=$(cat <<EOF
Proceso completado y depositado correctamente.
Fecha y hora: ${now_text}
Archivo origen: ${daily_file}
Ruta origen: ${WORK_DIR}
Archivo generado: ${out_file}
Destino FTP: ${FTP_HOST}:${FTP_REMOTE_DIR}${FTP_REMOTE_FILE}
Numero de lineas: ${line_count}
EOF
)
  if ! send_mail "$SUBJECT_SUCCESS" "$success_body"; then
    log "Error enviando correo de exito."
    exit 9
  fi
  log "Proceso exitoso. Lineas generadas: $line_count"

  exit 0
}

main "$@"

