#!/bin/bash
# script para consultar informacion de transferencias y proceso de reenvio
#
#Autor: Carlos Alfonso Ortega Molina
#############

#############################################################
# funcion para mostrar un titulo en pantalla
# no recibe parametros
# solo imprime el banner del sistema
#############################################################
mostrar_titulo() {
  clear
  echo "+-------------------------------------------------------+"
  echo "|             CONSULTA DE TRANSFERENCIAS                |"
  echo "+-------------------------------------------------------+"
  echo ""
}

#############################################################
# funcion para leer los datos basicos del usuario
# pide punto emisor tienda emisora tienda receptora y numeros de transferencia
# formatea los numeros de transferencia para usar en consultas sql
# permite dejar en blanco cualquiera de los tres filtros
#############################################################
leer_datos() {
  echo "Por favor, ingrese los siguientes datos:"
  echo "(puede dejar en blanco cualquier campo para no filtrarlo)"
  echo ""

  read -p "Punto emisor (pto_emi): " pto_emi
  read -p "Tienda emisora (cod_pto): " cod_pto
  read -p "Tienda receptora (pto_tra): " pto_tra

  echo -n "Número(s) de transferencia (num_rep), separados por coma si son varios: "
  read num_rep_input

  # CAMBIO: Permitir todos los campos en blanco
  if [[ -z "$pto_emi" && -z "$cod_pto" && -z "$pto_tra" && -z "$num_rep_input" ]]; then
    # Todos los filtros en blanco, flujo especial
    return 2
  fi

  # validar que al menos se ingreso un numero de transferencia
  # CAMBIO: Solo si todos los campos están en blanco, se permite continuar
  if [[ -z "$num_rep_input" ]]; then
    # Puede continuar si hay algún filtro, aunque no haya num_rep_input
    return 0
  fi

  # formatear numeros de transferencia para la consulta sql
  num_rep_list=$(echo $num_rep_input | tr ',' '\n' | tr -d ' ' | tr '\n' ',' | sed 's/,$//')

  return 0
}

#############################################################
# funcion para crear la condicion sql basada en los filtros
# genera una condicion sql apropiada segun los datos ingresados
# agrega solo los campos que tienen valor
#############################################################
generar_condicion_sql() {
  local condiciones=()

  # agregar condiciones solo para los campos con valor
  if [[ -n "$pto_emi" ]]; then
    condiciones+=("pto_emi = $pto_emi")
  fi

  if [[ -n "$cod_pto" ]]; then
    condiciones+=("cod_pto = $cod_pto")
  fi

  if [[ -n "$pto_tra" ]]; then
    condiciones+=("pto_tra = $pto_tra")
  fi

  # si no hay condiciones adicionales retornar cadena vacia
  if [ ${#condiciones[@]} -eq 0 ]; then
    echo ""
    return
  fi

  # si hay una sola condicion retornarla directamente
  if [ ${#condiciones[@]} -eq 1 ]; then
    echo "AND ${condiciones[0]}"
    return
  fi

  # si hay multiples condiciones unirlas con OR
  local condicion="AND ("
  local primero=true

  for c in "${condiciones[@]}"; do
    if [ "$primero" = true ]; then
      condicion="$condicion$c"
      primero=false
    else
      condicion="$condicion OR $c"
    fi
  done

  condicion="$condicion)"
  echo "$condicion"
}

#############################################################
# funcion para crear y ejecutar consultas sql en informix
# parametros:
# $1 - tabla a consultar
# $2 - campos a seleccionar
# $3 - archivo sql temporal
# $4 - archivo salida temporal
# $5 - mensaje para caso de no encontrar resultados
# retorna 0 si encuentra datos 1 si no hay datos
#############################################################
ejecutar_sql() {
  local tabla=$1
  local campos=$2
  local sql_file=$3
  local out_file=$4
  local mensaje_no_datos=$5
  local condicion=$(generar_condicion_sql)
  local encontrado=0

  # crear archivo sql para la consulta
  cat > $sql_file << EOF
DATABASE gen;
SET ISOLATION TO DIRTY READ;

SELECT $campos
  FROM $tabla
 WHERE num_rep IN ($num_rep_list)
 $condicion
 ORDER BY 2;
EOF

  # ejecutar la consulta y capturar la salida
  dbaccess - $sql_file > $out_file 2>&1

  # mostrar solo los datos sin mensajes del sistema ni encabezados de columnas
  while IFS= read -r linea; do
    # saltar lineas con mensajes del sistema
    if [[ $linea == *"Database selected"* ]] || [[ $linea == *"Isolation level set"* ]] ||
       [[ $linea == *"row(s) retrieved"* ]] || [[ $linea == *"Database closed"* ]]; then
        continue
    fi

    # detectar y saltar la linea de encabezados de columnas
    if [[ $linea == *"pto_emi"*"num_rep"* ]] || [[ $linea == *"No rows found"* ]]; then
        continue
    fi

    # si la linea no esta vacia mostrarla y actualizar la bandera
    if [[ -n $linea ]]; then
        # formatear la salida para alinear mejor las columnas
        # contar el numero de campos para aplicar el formato correcto
        num_campos=$(echo $campos | tr ',' '\n' | wc -l)

        if [ $num_campos -eq 4 ]; then
          formatted=$(echo "$linea" | awk '{printf "   %-5s      %-9s    %-5s    %-5s\n", $1, $2, $3, $4}')
        elif [ $num_campos -eq 6 ]; then
          formatted=$(echo "$linea" | awk '{printf "   %-5s      %-9s    %-5s    %-5s     %-5s  %-5s\n", $1, $2, $3, $4, $5, $6}')
        fi

        echo "$formatted"
        encontrado=1
    fi
  done < $out_file

  # si no se encontraron datos mostrar mensaje
  if [ $encontrado -eq 0 ]; then
    echo "   $mensaje_no_datos"
    return 1
  fi

  return 0
}

#############################################################
# funcion para ejecutar todas las consultas
# utiliza las variables globales para las consultas
# crea archivos temporales para los resultados
# imprime encabezados y resultados formateados
#############################################################
ejecutar_consultas() {
  # archivos temporales
  local sql_file="/tmp/consulta_$$.sql"
  local out_file="/tmp/output_$$.txt"

  mostrar_titulo
  echo "Ejecutando consultas para:"
  echo "- Punto emisor: ${pto_emi:-"(todos)"}"
  echo "- Tienda emisora: ${cod_pto:-"(todas)"}"
  echo "- Tienda receptora: ${pto_tra:-"(todas)"}"
  echo "- Número(s) de transferencia: $num_rep_input"
  echo ""
  echo "Por favor espere..."
  echo ""

  # primera consulta - tra_env
  echo "+===========================================================+"
  echo "|                    DATOS DE TRA_ENV                       |"
  echo "+===========================================================+"
  echo ""
  echo "+---------+---------+---------+---------+"
  echo "| PTO_EMI | NUM_REP | COD_PTO | PTO_TRA |"
  echo "+---------+---------+---------+---------+"

  # ejecutar consulta a tra_env
  ejecutar_sql "tra_env" "pto_emi, num_rep, cod_pto, pto_tra" "$sql_file" "$out_file" "No se encontraron datos en tra_env"

  echo ""
  echo "+===========================================================+"
  echo ""

  # segunda consulta - tra_cab
  echo "+===========================================================+"
  echo "|                    DATOS DE TRA_CAB                       |"
  echo "+===========================================================+"
  echo ""
  echo "+---------+---------+---------+---------+---------+---------+"
  echo "| PTO_EMI | NUM_REP | COD_PTO | PTO_TRA | NUM_BUL | CDT_TRA |"
  echo "+---------+---------+---------+---------+---------+---------+"

  # ejecutar consulta a tra_cab
  ejecutar_sql "tra_cab" "pto_emi, num_rep, cod_pto, pto_tra, num_bul, cdt_tra" "$sql_file" "$out_file" "No se encontraron datos en tra_cab"

  echo ""
  echo "+===========================================================+"
  echo ""

  # tercera consulta - repcab
  echo "+===========================================================+"
  echo "|                     DATOS DE REPCAB                       |"
  echo "+===========================================================+"
  echo ""
  echo "+---------+---------+---------+---------+---------+---------+"
  echo "| PTO_EMI | NUM_REP | COD_PTO | PTO_TRA | MAR_REP | MAR_INT |"
  echo "+---------+---------+---------+---------+---------+---------+"

  # ejecutar consulta a repcab
  ejecutar_sql "repcab" "pto_emi, num_rep, cod_pto, pto_tra, mar_rep, mar_int" "$sql_file" "$out_file" "No se encontraron datos en repcab"

  echo ""
  echo "+===========================================================+"

  # eliminar archivos temporales
  rm -f $sql_file $out_file
}

#############################################################
# funcion para verificar que un archivo exista y tenga permisos
# parametros:
# $1 - ruta del archivo a verificar
# retorna 0 si el archivo existe y tiene permisos 1 en otro caso
#############################################################
verificar_archivo() {
  local archivo=$1

  # verificar si existe el archivo
  if [ ! -f "$archivo" ]; then
    touch "$archivo" 2>/dev/null || {
      echo "error no se puede crear el archivo $archivo"
      return 1
    }
  fi

  # verificar permisos de escritura
  if [ ! -w "$archivo" ]; then
    echo "error no hay permisos para escribir en $archivo"
    return 1
  fi

  return 0
}

#############################################################
# CAMBIO: Nueva función para validar/capturar transferencias
# Si el archivo tiene líneas, preguntar si se usa o se captura uno nuevo
# Si se captura uno nuevo, vaciar el archivo y permitir ingresar transferencias
#############################################################
validar_o_capturar_transferencias() {
  local archivo_reproceso="/gnx_prod/manto/desa/trabajo/MesaControl/i40_repro.txt"
  verificar_archivo "$archivo_reproceso" || return 1

  local num_lineas=$(wc -l < "$archivo_reproceso")
  echo ""
  echo "El archivo de reproceso es: $archivo_reproceso"
  echo "Actualmente contiene $num_lineas línea(s)."
  echo ""

  if [ "$num_lineas" -gt 0 ]; then
    echo "Opciones para el archivo de reproceso:"
    echo "1) Trabajar con el archivo actual"
    echo "2) Capturar uno nuevo (vaciar el archivo y capturar transferencias)"
    read -p "Seleccione una opción [1/2]: " opcion_archivo

    case "$opcion_archivo" in
      1)
        echo "Se usará el archivo actual."
        ;;
      2)
        # Vaciar el archivo
        > "$archivo_reproceso"
        echo ""
        echo "Ingrese los números de transferencia uno por línea."
        echo "Presione 'x' o 'X' para finalizar la captura."
        while true; do
          read -p "Número de transferencia: " linea
          if [[ "$linea" == "x" || "$linea" == "X" ]]; then
            break
          fi
          # Solo guardar si no está vacío
          if [[ -n "$linea" ]]; then
            echo "$linea" >> "$archivo_reproceso"
          fi
        done
        ;;
      *)
        echo "Opción no válida. Se usará el archivo actual."
        ;;
    esac
  else
    echo "El archivo está vacío. Debe capturar al menos un número de transferencia."
    echo "Ingrese los números de transferencia uno por línea."
    echo "Presione 'x' o 'X' para finalizar la captura."
    while true; do
      read -p "Número de transferencia: " linea
      if [[ "$linea" == "x" || "$linea" == "X" ]]; then
        break
      fi
      if [[ -n "$linea" ]]; then
        echo "$linea" >> "$archivo_reproceso"
      fi
    done
  fi

  # Mostrar resumen
  local num_lineas_final=$(wc -l < "$archivo_reproceso")
  echo ""
  echo "El archivo contiene $num_lineas_final número(s) de transferencia."
  echo ""
  echo "Primeras líneas del archivo:"
  head -10 "$archivo_reproceso"
  echo ""
}

#############################################################
# funcion para generar el archivo de reproceso automaticamente
# genera un archivo con pto_emi y num_rep separados por pipes
# utiliza los datos de la consulta original
#############################################################
generar_archivo_reproceso() {
  local archivo_reproceso=$1
  local sql_file="/tmp/reproceso_$$.sql"
  local condicion=$(generar_condicion_sql)

  echo "Generando archivo de reproceso a partir de la consulta..."

  # crear consulta sql para generar archivo
  cat > $sql_file << EOF
DATABASE gen;
SET ISOLATION TO DIRTY READ;
UNLOAD TO '$archivo_reproceso'
DELIMITER '|'
SELECT pto_emi, num_rep
  FROM tra_env
 WHERE num_rep IN ($num_rep_list)
 $condicion
 ORDER BY 2;
EOF

  # ejecutar la consulta
  dbaccess - $sql_file > /dev/null 2>&1

  if [ $? -ne 0 ]; then
    echo "Error: No se pudo generar el archivo de reproceso"
    rm -f $sql_file
    return 1
  fi

  rm -f $sql_file
  return 0
}

#############################################################
# funcion para editar el archivo de reproceso
# permite al usuario elegir entre vaciar el archivo o usar el actual
# si elige vaciar genera el archivo de reproceso automaticamente
# cuenta las lineas en el archivo y muestra el resultado
#############################################################
editar_archivo_reproceso() {
  # ruta del archivo de reproceso
  local archivo_reproceso="/gnx_prod/manto/desa/trabajo/MesaControl/i40_repro.txt"

  # verificar si existe el archivo y permisos
  verificar_archivo "$archivo_reproceso" || return 1

  echo ""
  echo "El archivo de reproceso se encuentra en: $archivo_reproceso"

  # preguntar si desea vaciar el archivo o usar el actual
  echo ""
  echo "Opciones para el archivo de reproceso:"
  echo "1) Vaciar el archivo y crear uno nuevo con los datos de la consulta"
  echo "2) Usar el archivo actual"
  read -p "Seleccione una opción [1/2]: " opcion_archivo

  case "$opcion_archivo" in
    1)
      # vaciar el archivo sin borrarlo
      > "$archivo_reproceso"

      # generar archivo con datos de la consulta
      generar_archivo_reproceso "$archivo_reproceso" || return 1
      ;;
    2)
      # usar el archivo actual no hacer nada
      echo "Se usará el archivo actual."
      ;;
    *)
      echo "Opción no válida. Se usará el archivo actual."
      ;;
  esac

  # contar lineas en el archivo
  local num_lineas=$(wc -l < "$archivo_reproceso")
  echo ""
  echo "El archivo contiene $num_lineas número(s) de transferencia."

  # mostrar las primeras 10 lineas
  echo ""
  echo "Primeras líneas del archivo:"
  head -10 "$archivo_reproceso"
  echo ""

  # preguntar si desea continuar
  read -p "¿Desea continuar con el proceso? [S/N]: " respuesta
  if [[ "$respuesta" != "S" && "$respuesta" != "s" ]]; then
    echo "Proceso cancelado por el usuario."
    return 1
  fi

  return 0
}

#############################################################
# funcion para ejecutar un programa fgl y manejar errores
# parametros:
# $1 - nombre del programa a ejecutar
# $2 - mensaje de error en caso de fallo
# retorna el codigo de salida del programa ejecutado
#############################################################
ejecutar_programa() {
  local programa=$1
  local mensaje_error=$2

  echo "Ejecutando $programa..."
  fglgo $programa

  # verificar resultado
  if [ $? -ne 0 ]; then
    echo "error $mensaje_error"
    return 1
  fi

  return 0
}

#############################################################
# funcion para cambiar el nombre de archivos
# parametros:
# $1 - nombre del archivo original
# $2 - prefijo viejo a reemplazar
# $3 - prefijo nuevo
# $4 - mensaje para mostrar
# retorna 0 si tiene exito 1 en caso de error
#############################################################
cambiar_nombre_archivo() {
  local archivo_original=$1
  local prefijo_viejo=$2
  local prefijo_nuevo=$3
  local mensaje=$4

  if [ -n "$archivo_original" ]; then
    local nuevo_nombre=${archivo_original/$prefijo_viejo/$prefijo_nuevo}
    cp "$archivo_original" "$nuevo_nombre" || {
      echo "error al copiar $archivo_original a $nuevo_nombre"
      return 1
    }
    echo "$mensaje $archivo_original copiado a $nuevo_nombre"
    return 0
  fi

  return 1
}

#############################################################
# funcion para procesar archivos csv y cambiar sus nombres
# busca los archivos mas recientes del propietario desa
# copia los archivos cambiando el prefijo en a se
#############################################################
procesar_archivos_csv() {
  # cambiar a la ruta de reportes
  cd "/respaldo_migracion/reportes_gnx" || {
    echo "error no se puede acceder a la ruta de reportes"
    return 1
  }


  # validar que se encontraron los archivos
  if [ -z "$asn_to_tie" ] || [ -z "$ordentie" ] || [ -z "$or_files" ]; then
    echo "error no se encontraron todos los archivos necesarios"
    return 1
  fi

  # cambiar nombres de archivos
  echo "Cambiando nombres de archivos..."



  # mostrar listado y conteo de lineas
  echo ""
  echo "Listado de archivos generados:"
  ls -lt SE_*.csv | head -3

  echo ""
  echo "Conteo de líneas en archivos generados:"
  wc -l SE_*.csv | head -4

  return 0
}

#############################################################
# funcion para el proceso de reenvio
# ejecuta los programas fgl necesarios para el reproceso
# permite al usuario elegir si desea entrar al proceso forzado
# procesa los archivos csv cambiando nombres
#############################################################
proceso_reenvio() {
  # ruta de trabajo
  local ruta_trabajo="/gnx_prod/manto/desa/trabajo/MesaControl"

  echo ""
  read -p "¿Desea proceder con los reprocesos? [S/N]: " respuesta

  if [[ "$respuesta" == "S" || "$respuesta" == "s" ]]; then
    echo ""
    # cambiar al directorio de trabajo
    cd "$ruta_trabajo" || {
      echo "error no se puede acceder a la ruta de trabajo $ruta_trabajo"
      return 1
    }

    # ejecutar limpieza de tablas
    ejecutar_programa "Limpia_Tabla" "fallo la ejecucion de limpia_tabla" || return 1

    # ejecutar reproceso de transferencias
    echo ""
    ejecutar_programa "integracion71_forzado" "fallo la ejecucion de integracion71_forzado" || return 1

    echo ""
    read -p "¿Desea entrar al proceso forzado? [S/N]: " proceso_forzado

    if [[ "$proceso_forzado" == "S" || "$proceso_forzado" == "s" ]]; then
      echo ""
      echo "Comienza el forzado..."

      # ejecutar limpieza de tablas nuevamente
      ejecutar_programa "Limpia_Tabla" "fallo la ejecucion de limpia_tabla" || return 1

      # ejecutar proceso ora_tran_40_10
      echo ""
      ejecutar_programa "Integracion_40_23_forzado" "fallo la ejecucion de Integracion_40_23_forzado" || return 1

      # procesar archivos csv
      echo ""
      echo "Procesando archivos de reportes..."
      procesar_archivos_csv || return 1
    fi

    return 0
  fi

  echo "Proceso de reenvío cancelado por el usuario."
  return 0
}

#############################################################
# funcion principal del programa
# coordina la ejecucion de las distintas funciones
# maneja el flujo general del programa
#############################################################
main() {
  mostrar_titulo
  leer_datos
  resultado_leer_datos=$?

  if [ "$resultado_leer_datos" -eq 2 ]; then
    # CAMBIO: Si todos los filtros están en blanco, ir directo a validación/captura de archivo
    validar_o_capturar_transferencias

    # Preguntar si desea comenzar el proceso de reenvío
    read -p "¿Desea comenzar el proceso de reenvío? [S/N]: " iniciar_reenvio
    if [[ "$iniciar_reenvio" == "S" || "$iniciar_reenvio" == "s" ]]; then
      proceso_reenvio
    fi

    echo ""
    echo "Proceso completado."
    echo ""
    read -p "Presione Enter para salir..." tecla
    return
  fi

  # Si hay filtros o números de transferencia, flujo normal
  ejecutar_consultas

  echo ""
  echo "Consulta finalizada."
  echo ""

  # preguntar si desea comenzar el proceso de reenvio
  read -p "¿Desea comenzar el proceso de reenvío? [S/N]: " iniciar_reenvio

  if [[ "$iniciar_reenvio" == "S" || "$iniciar_reenvio" == "s" ]]; then
    editar_archivo_reproceso && proceso_reenvio
  fi

  echo ""
  echo "Proceso completado."
  echo ""
  read -p "Presione Enter para salir..." tecla
}

# ejecutar programa principal
main
