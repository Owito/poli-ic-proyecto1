"""Capturas REALES de la interfaz de Jenkins con un navegador sin ventana (Playwright + Microsoft Edge).

Uso (con el entorno virtual que tiene playwright instalado):
    python tools/jenkins_capturas.py asistente <archivo_credenciales>
    python tools/jenkins_capturas.py pantalla <archivo_credenciales> <ruta_jenkins> <nombre_png> [selector_a_esperar]

El archivo de credenciales vive FUERA del repositorio y tiene dos líneas: usuario=... y clave=...
Las claves se escriben en campos de tipo password, así que en las capturas salen enmascaradas.
"""
import subprocess
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

RAIZ = Path(__file__).resolve().parent.parent
EVID = RAIZ / "evidencias"
URL = "http://localhost:8181"
VISTA = {"width": 1366, "height": 820}


def credenciales(ruta):
    datos = dict(l.split("=", 1) for l in Path(ruta).read_text(encoding="utf-8").splitlines() if "=" in l)
    return datos["usuario"], datos["clave"]


def foto(page, nombre, completa=False):
    page.wait_for_timeout(800)
    page.screenshot(path=str(EVID / nombre), full_page=completa)
    print("captura", nombre)


def asistente(ruta_cred):
    usuario, clave = credenciales(ruta_cred)
    inicial = subprocess.run(["docker", "exec", "jenkins", "cat", "/var/jenkins_home/secrets/initialAdminPassword"],
                             capture_output=True, text=True, check=True).stdout.strip()
    with sync_playwright() as p:
        nav = p.chromium.launch(channel="msedge", headless=True)
        page = nav.new_page(viewport=VISTA)
        page.goto(URL + "/", wait_until="networkidle")
        page.wait_for_selector("#security-token", timeout=120_000)
        # Si el asistente se reanuda, no se repiten las capturas que ya se tomaron la primera vez.
        if not (EVID / "j02-desbloqueo.png").exists():
            foto(page, "j02-desbloqueo.png")
        page.fill("#security-token", inicial)
        page.click("input[type=submit], button[type=submit]")
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(4000)
        if page.query_selector("text=Install suggested plugins"):
            foto(page, "j03-personalizar.png")
            page.click("text=Install suggested plugins")
            # Progreso de la instalación: se captura a mitad de camino.
            page.wait_for_selector(".plugin-list, .install-console", timeout=120_000)
            page.wait_for_timeout(15_000)
            foto(page, "j04-instalando-plugins.png")
        marco = None
        for _ in range(180):  # hasta 15 minutos
            for f in page.frames:
                if "setupWizardFirstUser" in f.url:
                    marco = f
            if marco:
                break
            time.sleep(5)
        if not marco:
            raise SystemExit("No apareció el formulario del primer usuario")
        marco.wait_for_selector("input[name=username]")
        marco.fill("input[name=username]", usuario)
        marco.fill("input[name=password1]", clave)
        marco.fill("input[name=password2]", clave)
        marco.fill("input[name=fullname]", "Administrador del proyecto")
        marco.fill("input[name=email]", "admin@tareas.local")
        foto(page, "j05-primer-administrador.png")
        page.click("button.save-first-user")
        conf = None
        for _ in range(60):
            for f in page.frames:
                if "setupWizardConfigureInstance" in f.url:
                    conf = f
            if conf:
                break
            time.sleep(2)
        conf.wait_for_selector("input[name=rootUrl]")
        conf.fill("input[name=rootUrl]", URL + "/")
        foto(page, "j06-url-instancia.png")
        page.click("button.save-configure-instance")
        page.wait_for_selector("text=Jenkins is ready", timeout=120_000)
        foto(page, "j07-listo.png")
        page.click("button.install-done")
        page.wait_for_load_state("networkidle")
        foto(page, "j08-tablero-inicial.png")
        nav.close()


def cierre_asistente(ruta_cred):
    """Termina el asistente cuando el primer administrador ya existe (Instance Configuration y Jenkins is ready)."""
    usuario, clave = credenciales(ruta_cred)
    with sync_playwright() as p:
        nav = p.chromium.launch(channel="msedge", headless=True)
        page = nav.new_page(viewport=VISTA)
        page.goto(URL + "/login", wait_until="networkidle")
        page.fill("input[name=j_username]", usuario)
        page.fill("input[name=j_password]", clave)
        page.click("button[type=submit]")
        page.wait_for_load_state("networkidle")
        page.goto(URL + "/", wait_until="networkidle")
        page.wait_for_timeout(3000)
        conf = next(f for f in page.frames if "setupWizardConfigureInstance" in f.url)
        conf.wait_for_selector("input[name=rootUrl]")
        conf.fill("input[name=rootUrl]", URL + "/")
        foto(page, "j06-url-instancia.png")
        page.click("button.save-configure-instance")
        page.wait_for_selector("text=Jenkins is ready", timeout=120_000)
        foto(page, "j07-listo.png")
        page.click("button.install-done")
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(3000)
        foto(page, "j08-tablero-inicial.png")
        nav.close()


def pantalla(ruta_cred, ruta, nombre, esperar=None, completa="no"):
    usuario, clave = credenciales(ruta_cred)
    with sync_playwright() as p:
        nav = p.chromium.launch(channel="msedge", headless=True)
        page = nav.new_page(viewport=VISTA)
        page.goto(URL + "/login", wait_until="networkidle")
        page.fill("input[name=j_username]", usuario)
        page.fill("input[name=j_password]", clave)
        page.click("button[name=Submit], input[name=Submit], button[type=submit]")
        page.wait_for_load_state("networkidle")
        page.goto(URL + ruta, wait_until="networkidle")
        if esperar:
            page.wait_for_selector(esperar, timeout=60_000)
        foto(page, nombre, completa == "si")
        nav.close()


if __name__ == "__main__":
    EVID.mkdir(exist_ok=True)
    if sys.argv[1] == "asistente":
        asistente(sys.argv[2])
    elif sys.argv[1] == "cierre":
        cierre_asistente(sys.argv[2])
    elif sys.argv[1] == "pantalla":
        pantalla(*sys.argv[2:])
