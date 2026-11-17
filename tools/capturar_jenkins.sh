#!/usr/bin/env bash
# Evidencias de terminal de la Entrega 2 (Jenkins). Igual que tools/capturar.sh: ejecuta de verdad cada comando,
# guarda la salida en evidencias/texto/ y la convierte en PNG con tools/render_captura.py.
# Uso: bash tools/capturar_jenkins.sh <evidencia> [numero_de_ejecucion]
#   levantar | cliente | espejo | consola N | dind
# La consola se lee con la API de Jenkins; el usuario y la clave salen del archivo local de credenciales
# (variable CRED, fuera del repositorio) y nunca se imprimen.
cd "$(dirname "$0")/.." || exit 1
mkdir -p evidencias/texto
CRED="${CRED:-/c/Users/fgogu/poli-ic-proyecto2-jenkins-local/credenciales.txt}"

cap() {  # cap <nombre> <comando visible> [comando real]
  local f="evidencias/texto/$1.txt"
  { echo "\$ $2"; bash -c "${3:-$2}" 2>&1; } > "$f"
  echo "[$1] listo ($(wc -l < "$f") líneas)"
  python tools/render_captura.py "$f"
}

auth() { echo "$(grep '^usuario=' "$CRED" | cut -d= -f2):$(grep '^clave=' "$CRED" | cut -d= -f2)"; }

case "$1" in
  levantar)
    cap j01-levantar "docker network ls --filter name=jenkins; docker image ls tareas-jenkins; docker ps --filter name=jenkins --format 'table {{.Names}}\t{{.Image}}\t{{.Status}}\t{{.Ports}}'"
    ;;
  cliente)
    cap j09-cliente-docker "docker exec jenkins docker version --format 'cliente {{.Client.Version}} -> motor {{.Server.Version}} ({{.Server.Os}})'; docker exec jenkins docker compose version; docker exec jenkins git --version; docker exec jenkins bash -c 'echo DOCKER_HOST=\$DOCKER_HOST'"
    ;;
  espejo)
    cap j13-espejo "git push jenkins-local poli-trabajos/ic-proyecto2-jenkins; git ls-remote jenkins-local poli-trabajos/ic-proyecto2-jenkins; git log -1 --format='%h %s'"
    ;;
  consola)
    n="$2"
    cap "j-consola-$n" "curl -s -u <usuario>:<clave> http://localhost:8181/job/tareas-equipo/$n/consoleText | grep -E -v '^\\s*\$' | ..." \
      "curl -s -u \"$(auth)\" http://localhost:8181/job/tareas-equipo/$n/consoleText | sed -E 's/^\\[[0-9T:.Z-]+\\] //' | grep -E '^(Started|Obtained|Checking out|Commit|\\[Pipeline\\] (stage|\\{ \\()|v?[0-9]+\\.[0-9]+\\.[0-9]+$|added|# |ok |not ok|.*Built|.*Healthy|.*Started|.*Running|NAME|tareas-|\\{\"estado|.*Removed|Integración|La integración|Finished)' | head -80"
    ;;
  dind)
    cap j-dind "docker exec jenkins-docker docker ps --format 'table {{.Names}}\t{{.Image}}\t{{.Status}}'; docker exec jenkins-docker docker image ls --format 'table {{.Repository}}\t{{.Tag}}\t{{.Size}}'"
    ;;
esac
