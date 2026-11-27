"""Extracto de un registro real: conserva las líneas que coinciden con los patrones y marca con [...] las omitidas.

Uso: python tools/extracto_log.py <registro_completo.txt> <extracto.txt>
No cambia ninguna línea: solo elige cuáles se conservan.
"""
import re
import sys
from pathlib import Path

CLAVE = re.compile(
    r"^\$ |^=====|^== |^-- paso|^v\d+\.\d+|^\d+\.\d+\.\d+$|^added \d+ packages|^Docker (version|Compose)|^not ok|"
    r"^# (tests|suites|pass|fail|cancelled|skipped|todo) \d|Image tareas-ci-app Built|Container tareas-(db|app) Healthy|"
    r"^NAME |^tareas-(app|db) |\"estado\":|Volume tareas-ci_tareas-datos Removed|Network tareas-ci_red-tareas Removed|"
    r"^\(código de salida|^  \"node --version|^docker compose ps$|^rm -f \.env$")

origen, destino = Path(sys.argv[1]), Path(sys.argv[2])
salida, omitidas = [], False
for linea in origen.read_text(encoding="utf-8").replace("\r", "").split("\n"):
    if CLAVE.search(linea):
        if omitidas and salida:
            salida.append("[...]")
        salida.append(linea)
        omitidas = False
    elif linea.strip():
        omitidas = True
destino.write_text("\n".join(salida) + "\n", encoding="utf-8")
print(destino, len(salida), "líneas")
