#!/usr/bin/env bash
# Ejecuta de verdad cada comando de la evidencia y guarda su salida en evidencias/texto/NN-nombre.txt.
# Después tools/render_captura.py convierte cada .txt en PNG. Uso: bash tools/capturar.sh
cd "$(dirname "$0")/.." || exit 1
mkdir -p evidencias/texto

cap() {  # cap <nombre> <comando>
  local f="evidencias/texto/$1.txt"
  { echo "\$ $2"; bash -c "$2" 2>&1; } > "$f"
  echo "[$1] listo ($(wc -l < "$f") líneas)"
}
bonito() { python -c "import sys,json;print(json.dumps(json.load(sys.stdin),indent=2,ensure_ascii=False))"; }
export -f bonito

cap 01-versiones "docker --version; docker compose version; git --version; node --version"
cap 02-estructura "git ls-files | grep -v -e '^evidencias/' -e '^docs/' -e '^tools/' -e package-lock"
cap 03-compose-up "docker compose up -d --build 2>&1 | tail -30"
cap 04-compose-ps "docker compose ps"
cap 05-salud-db "docker inspect tareas-db --format 'contenedor={{.Name}} estado={{.State.Status}} salud={{.State.Health.Status}}'; docker inspect tareas-app --format 'contenedor={{.Name}} estado={{.State.Status}}'"
cap 06-red "docker network inspect tareas-equipo_red-tareas --format '{{.Name}} driver={{.Driver}}{{range \$id,\$c := .Containers}}{{println}}  {{\$c.Name}}  ip={{\$c.IPv4Address}}{{end}}'"
cap 07-volumen "docker volume inspect tareas-equipo_tareas-datos --format 'volumen={{.Name}} driver={{.Driver}} montaje={{.Mountpoint}}'"
cap 08-curl-salud "curl -s http://localhost:8095/salud | bonito"
cap 09-curl-tareas "curl -s http://localhost:8095/tareas | bonito"
cap 10-psql "docker compose exec -T db sh -c 'psql -U \$POSTGRES_USER -d \$POSTGRES_DB -c \"SELECT id, titulo, responsable, estado FROM tareas ORDER BY id;\"'"
cap 11-dns-desde-app "docker compose exec -T app node -e \"require('dns').lookup('db',(e,ip)=>console.log('desde el contenedor tareas-app, el nombre db resuelve a', ip))\""
cap 12-crud "curl -s -X POST http://localhost:8095/tareas -H 'Content-Type: application/json' -d '{\"titulo\":\"Entregar la Entrega 1 en Canvas\",\"responsable\":\"Grupo 14\"}' | bonito; echo ---; curl -s -X PUT http://localhost:8095/tareas/6 -H 'Content-Type: application/json' -d '{\"estado\":\"en_progreso\"}' | bonito"
cap 13-psql-verifica "docker compose exec -T db sh -c 'psql -U \$POSTGRES_USER -d \$POSTGRES_DB -c \"SELECT id, titulo, estado FROM tareas WHERE id = 6;\"'"
cap 14-validacion "curl -s -o /dev/null -w 'POST sin titulo -> HTTP %{http_code}\n' -X POST http://localhost:8095/tareas -H 'Content-Type: application/json' -d '{}'; curl -s -o /dev/null -w 'GET /tareas/9999 -> HTTP %{http_code}\n' http://localhost:8095/tareas/9999"
cap 15-pruebas "docker compose exec -T app npm test 2>&1 | grep -E '^(ok|not ok)|^# (tests|pass|fail|skipped)'"
cap 16-usuario-no-root "docker compose exec -T app id; docker compose exec -T app node --version"
cap 17-persistencia "docker compose down; docker compose up -d; sleep 12; docker compose ps; curl -s http://localhost:8095/tareas/6 | bonito"
cap 18-git-log "git log --stat --format='%h %an %ad %s' --date=short | head -40"

python tools/render_captura.py evidencias/texto/*.txt
