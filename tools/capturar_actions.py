"""Captura REAL de la página pública de una ejecución de GitHub Actions, sin iniciar sesión (Playwright + Microsoft Edge).

Uso (con el entorno virtual que tiene playwright instalado):
    python tools/capturar_actions.py <id_de_la_ejecucion> <archivo.png>
La página no se modifica: se espera a que aparezcan los dos jobs y se guarda la vista.
"""
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

REPO = "Owito/poli-ic-proyecto1"
EVID = Path(__file__).resolve().parent.parent / "evidencias"
JOBS = ("Instalar dependencias y pruebas unitarias", "pruebas de integración")

ejecucion, archivo = sys.argv[1], sys.argv[2]
url = f"https://github.com/{REPO}/actions/runs/{ejecucion}"

with sync_playwright() as p:
    nav = p.chromium.launch(channel="msedge", headless=True)
    ctx = nav.new_context(viewport={"width": 1400, "height": 1000}, locale="es-CO")
    page = ctx.new_page()
    page.goto(url, wait_until="networkidle", timeout=60000)
    for job in JOBS:
        page.get_by_text(job, exact=False).first.wait_for(state="visible", timeout=30000)
    page.wait_for_timeout(1500)
    page.screenshot(path=str(EVID / archivo), full_page=False)
    print("captura", archivo, "desde", page.url, "| título:", page.title())
    nav.close()
