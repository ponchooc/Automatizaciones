#!/usr/bin/bash
set -u

SENDER="sistemas_mercaderias@sears.com.mx"

usage() {
  cat <<EOF
Uso: $0 <destinatarios> <asunto> <cuerpo> [adjunto]
  destinatarios: uno o varios (separar por comas o espacios)
  asunto: texto (obligatorio)
  cuerpo: texto (puede ser "")
  adjunto: ruta opcional a archivo a adjuntar
Ejemplo:
  $0 "a@dom.com,b@dom.com" "Asunto" "Cuerpo del mensaje" /tmp/archivo.zip
  printf 'Linea1\nLinea2\n' | $0 "a@dom.com" "Asunto" ""    # cuerpo vacío
EOF
  exit 1
}

# Parámetros posicionales
if [ $# -lt 2 ]; then
  usage
fi

recipients_raw="$1"
subject="$2"
body="${3:-}"
attachment="${4:-}"

# Validaciones mínimas
if [ -z "$recipients_raw" ] || [ -z "$subject" ]; then
  usage
fi

# Normalizar destinatarios: convertir comas en espacios
recipients=$(echo "$recipients_raw" | tr ',' ' ')

# Comprobar mailx
if ! command -v mailx >/dev/null 2>&1; then
  echo "ERROR: 'mailx' no encontrado en el sistema." >&2
  exit 2
fi

# Envío
if [ -n "$attachment" ]; then
  if [ ! -f "$attachment" ]; then
    echo "ERROR: archivo adjunto no encontrado: $attachment" >&2
    exit 3
  fi

  (
    printf '%s\n' "$body"
    uuencode "$attachment" "$(basename "$attachment")"
  ) | mailx -s "$subject" -r "$SENDER" $recipients
  rc=$?
else
  printf '%s\n' "$body" | mailx -s "$subject" -r "$SENDER" $recipients
  rc=$?
fi

if [ $rc -eq 0 ]; then
  echo "Correo enviado correctamente."
else
  echo "ERROR: fallo al enviar correo (código $rc)." >&2
fi

exit $rc
