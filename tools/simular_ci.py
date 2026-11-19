"""Ejecuta en el equipo local los comandos de .travis.yml o del flujo de GitHub Actions, leídos de los propios archivos.

Sirve para comprobar, antes de conectar los servicios, que lo que dicen los archivos de configuración funciona. No
reemplaza la ejecución en Travis CI ni en GitHub Actions: no usa sus máquinas virtuales, sino Bash y Docker del equipo.

Uso:
    python tools/simular_ci.py travis  [nombre de etapa ...]
    python tools/simular_ci.py actions [nombre de job ...]
Sin nombres, ejecuta todas las etapas o jobs en orden y se detiene en la primera que falle, como los dos servicios.

Reglas que imita:
  - Travis CI: cada etapa es un único script de Bash. Si falla before_install, install o before_script, el job queda
    "errored" y no sigue; si falla un comando de script, los demás corren igual y el job queda "failed"; after_script
    corre siempre (docs.travis-ci.com/user/job-lifecycle/).
  - GitHub Actions: cada paso run es un proceso de Bash con -e; si un paso falla, se saltan los siguientes, salvo los
    que tienen if: always(). Los pasos uses (actions/checkout) se omiten: el código ya está en la carpeta.
Diferencia propia de Windows: Git Bash convierte las rutas que parecen de Unix; se desactiva con MSYS_NO_PATHCONV=1 y
$PWD se pasa en formato C:/..., que es el que entiende Docker Desktop.
"""
import os
import shlex
import subprocess
import sys
import tempfile
from pathlib import Path

import yaml

RAIZ = Path(__file__).resolve().parent.parent
BASH = os.environ.get("BASH_CI", r"C:\Program Files\Git\bin\bash.exe" if os.name == "nt" else "/bin/bash")
PREAMBULO = 'export MSYS_NO_PATHCONV=1; export PWD="$(pwd -W 2>/dev/null || pwd)"\n' if os.name == "nt" else ""


def bash(script, entorno):
    with tempfile.NamedTemporaryFile("w", suffix=".sh", delete=False, encoding="utf-8", newline="\n") as f:
        f.write(PREAMBULO + script)
    try:
        return subprocess.run([BASH, f.name], cwd=RAIZ, env={**os.environ, **entorno}).returncode
    finally:
        os.unlink(f.name)


def eco(cmd):
    return f"printf '\\n$ %s\\n' {shlex.quote(cmd)}\n"


def lista(v):
    return [] if v in (None, "skip") else ([v] if isinstance(v, str) else list(v))


def travis(nombres):
    cfg = yaml.safe_load((RAIZ / ".travis.yml").read_text(encoding="utf-8"))
    for job in cfg["jobs"]["include"]:
        etapa = job["stage"]
        if nombres and etapa not in nombres:
            continue
        print(f"\n===== Travis CI, etapa «{etapa}»: {job.get('name', '')}", flush=True)
        entorno = dict(p.split("=", 1) for p in shlex.split(job.get("env", "")))
        s = "estado=0\n"
        for fase in ("before_install", "install", "before_script"):
            for cmd in lista(job.get(fase)):
                s += f"if [ $estado -eq 0 ]; then {eco(cmd)}{cmd}\n[ $? -eq 0 ] || {{ echo '== falló {fase}: el job queda errored'; estado=2; }}\nfi\n"
        for cmd in lista(job.get("script")):
            s += f"if [ $estado -ne 2 ]; then {eco(cmd)}{cmd}\n[ $? -eq 0 ] || estado=1\nfi\n"
        for cmd in lista(job.get("after_script")):
            s += f"{eco(cmd)}{cmd}\n"
        s += "case $estado in 0) r=passed;; 1) r=failed;; *) r=errored;; esac\necho \"== resultado de la etapa: $r\"\nexit $estado\n"
        if bash(s, entorno) != 0:
            print("== Travis cancela las etapas siguientes", flush=True)
            return 1
    return 0


def actions(nombres):
    cfg = yaml.safe_load((RAIZ / ".github" / "workflows" / "integracion-continua.yml").read_text(encoding="utf-8"))
    for nombre, job in cfg["jobs"].items():
        if nombres and nombre not in nombres:
            continue
        print(f"\n===== GitHub Actions, job «{nombre}»: {job.get('name', '')}", flush=True)
        entorno = {k: str(v) for k, v in (job.get("env") or {}).items()}
        fallo = False
        for paso in job["steps"]:
            if "run" not in paso:
                print(f"\n-- paso «{paso.get('name', paso.get('uses'))}» ({paso.get('uses')}): omitido, el código ya está en la carpeta", flush=True)
                continue
            if fallo and paso.get("if") != "always()":
                print(f"\n-- paso «{paso['name']}»: saltado por la falla anterior", flush=True)
                continue
            print(f"\n-- paso «{paso['name']}»", flush=True)
            cuerpo = paso["run"].rstrip("\n")
            rc = bash("set -e\n" + eco(cuerpo) + cuerpo + "\n", entorno)
            fallo = fallo or rc != 0
        print(f"== resultado del job: {'failure' if fallo else 'success'}", flush=True)
        if fallo:
            print("== los jobs que dependen de este (needs) se saltan", flush=True)
            return 1
    return 0


if __name__ == "__main__":
    modo, nombres = sys.argv[1], sys.argv[2:]
    sys.exit({"travis": travis, "actions": actions}[modo](nombres))
