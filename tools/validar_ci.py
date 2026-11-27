"""Valida, sin ejecutar nada, la configuración de Travis CI y de GitHub Actions del proyecto.

Uso: python tools/validar_ci.py <schema_travis.json> <schema_github_workflow.json>
  - schema_travis.json: el esquema oficial de Travis CI, archivo schema.json del repositorio travis-ci/travis-yml.
  - schema_github_workflow.json: el esquema de flujos de GitHub Actions publicado en https://json.schemastore.org/github-workflow.json
Imprime cada control con OK o FALLA y termina con el número de fallas (código de salida 1 si hay alguna).

    python tools/validar_ci.py --comandos-travis
Imprime los comandos de .travis.yml como un script de Bash, para revisarlos con ShellCheck.
"""
import json
import sys
from pathlib import Path

import yaml
from jsonschema import Draft4Validator, Draft7Validator

RAIZ = Path(__file__).resolve().parent.parent
TRAVIS = RAIZ / ".travis.yml"
FLUJO = RAIZ / ".github" / "workflows" / "integracion-continua.yml"
fallas = 0


def control(nombre, ok, detalle=""):
    global fallas
    fallas += 0 if ok else 1
    print(("OK    " if ok else "FALLA ") + nombre + (f": {detalle}" if detalle else ""))


def errores(validador, datos):
    return [f"{'/'.join(map(str, e.absolute_path)) or '(raíz)'}: {e.message[:120]}" for e in validador.iter_errors(datos)]


if sys.argv[1:] == ["--comandos-travis"]:
    sys.stdout.reconfigure(newline="\n")
    print("#!/usr/bin/env bash")
    for job in yaml.safe_load(TRAVIS.read_text(encoding="utf-8"))["jobs"]["include"]:
        print(f"# etapa {job['stage']}")
        for fase in ("install", "before_script", "script", "after_script"):
            v = job.get(fase, [])
            for cmd in [] if v == "skip" else ([v] if isinstance(v, str) else v):
                print(cmd)
    sys.exit(0)

# 1. Sintaxis YAML
travis = yaml.safe_load(TRAVIS.read_text(encoding="utf-8"))
flujo = yaml.safe_load(FLUJO.read_text(encoding="utf-8"))
control("YAML válido: .travis.yml", isinstance(travis, dict))
control("YAML válido: .github/workflows/integracion-continua.yml", isinstance(flujo, dict))
# YAML 1.1 lee la clave "on" como el booleano True; GitHub la interpreta como texto.
if True in flujo:
    flujo["on"] = flujo.pop(True)

# 2. Esquemas
esq_travis = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
esq_flujo = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))
e = errores(Draft4Validator(esq_travis), travis)
control("Esquema oficial de Travis CI (travis-ci/travis-yml)", not e, "; ".join(e[:5]))
e = errores(Draft7Validator(esq_flujo), flujo)
control("Esquema de flujos de GitHub Actions (SchemaStore)", not e, "; ".join(e[:5]))

# 3. Que los dos archivos hagan lo mismo que el Jenkinsfile, en el mismo orden
etapas = [j.get("stage") for j in travis["jobs"]["include"]]
control("Travis: etapas en orden", etapas == ["pruebas unitarias", "contenedores"], str(etapas))
control("GitHub Actions: contenedores espera a pruebas-unitarias", flujo["jobs"]["contenedores"].get("needs") == "pruebas-unitarias")


def comandos_travis():
    salida = []
    for j in travis["jobs"]["include"]:
        for fase in ("install", "before_script", "script"):
            v = j.get(fase, [])
            salida += [v] if isinstance(v, str) else v
    return " ".join(salida)


def comandos_flujo():
    return " ".join(p.get("run", "") for j in flujo["jobs"].values() for p in j["steps"])


PASOS = ["npm ci", "npm test", "docker compose build app", "docker compose up -d --wait", "/salud",
         "docker compose exec -T app npm test", "docker compose down -v"]
for nombre, texto in (("Travis", comandos_travis() + " " + " ".join(travis["jobs"]["include"][1]["after_script"])),
                      ("GitHub Actions", comandos_flujo())):
    faltan = [p for p in PASOS if p not in texto]
    control(f"{nombre}: contiene los pasos del Jenkinsfile", not faltan, f"faltan {faltan}" if faltan else "")

ultimo = flujo["jobs"]["contenedores"]["steps"][-1]
control("GitHub Actions: la limpieza corre siempre", ultimo.get("if") == "always()" and "down -v" in ultimo.get("run", ""))
control("Ninguna clave escrita en los archivos", "POSTGRES_PASSWORD=%s" in TRAVIS.read_text(encoding="utf-8")
        and "POSTGRES_PASSWORD=%s" in FLUJO.read_text(encoding="utf-8"))

print("Total de fallas:", fallas)
sys.exit(1 if fallas else 0)
