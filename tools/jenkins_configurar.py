"""Configura el Jenkins local para el proyecto y toma las capturas de cada paso.

Uso (entorno virtual con playwright):
    python tools/jenkins_configurar.py credencial <archivo_credenciales>
    python tools/jenkins_configurar.py trabajo <archivo_credenciales>

- credencial: crea la credencial "tareas-db-password" (tipo Secret text) con la clave de PostgreSQL que usará el
  pipeline. La clave sale del archivo local de credenciales (línea clave_db_ci=...) y no se imprime.
- trabajo: crea desde la interfaz web el trabajo "tareas-equipo" de tipo Pipeline, leído desde Git (Jenkinsfile).
"""
import json
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

RAIZ = Path(__file__).resolve().parent.parent
EVID = RAIZ / "evidencias"
URL = "http://localhost:8181"
VISTA = {"width": 1366, "height": 820}
REPO = "file:///repos/poli-ic-proyecto1.git"
RAMA = "*/poli-trabajos/ic-proyecto2-jenkins"


def datos(ruta):
    return dict(l.split("=", 1) for l in Path(ruta).read_text(encoding="utf-8").splitlines() if "=" in l)


def entrar(p, d):
    nav = p.chromium.launch(channel="msedge", headless=True)
    page = nav.new_page(viewport=VISTA)
    page.goto(URL + "/login", wait_until="networkidle")
    page.fill("input[name=j_username]", d["usuario"])
    page.fill("input[name=j_password]", d["clave"])
    page.click("button[type=submit]")
    page.wait_for_load_state("networkidle")
    return nav, page


def foto(page, nombre, completa=False):
    page.wait_for_timeout(800)
    page.screenshot(path=str(EVID / nombre), full_page=completa)
    print("captura", nombre)


def credencial(ruta):
    d = datos(ruta)
    with sync_playwright() as p:
        nav, page = entrar(p, d)
        crumb = page.request.get(URL + "/crumbIssuer/api/json").json()
        script = (
            "import com.cloudbees.plugins.credentials.*\n"
            "import com.cloudbees.plugins.credentials.domains.Domain\n"
            "import org.jenkinsci.plugins.plaincredentials.impl.StringCredentialsImpl\n"
            "import hudson.util.Secret\n"
            "def store = SystemCredentialsProvider.getInstance().getStore()\n"
            "def viejas = store.getCredentials(Domain.global()).findAll { it.id == 'tareas-db-password' }\n"
            "viejas.each { store.removeCredentials(Domain.global(), it) }\n"
            "store.addCredentials(Domain.global(), new StringCredentialsImpl(CredentialsScope.GLOBAL, "
            "'tareas-db-password', 'Clave de PostgreSQL para el pipeline de tareas', "
            f"Secret.fromString({json.dumps(d['clave_db_ci'])})))\n"
            "println 'credencial creada: tareas-db-password'\n"
        )
        r = page.request.post(URL + "/scriptText", form={"script": script},
                              headers={crumb["crumbRequestField"]: crumb["crumb"]})
        print(r.status, r.text().strip())
        page.goto(URL + "/manage/credentials/", wait_until="networkidle")
        foto(page, "j10-credencial.png")
        nav.close()


def trabajo(ruta):
    d = datos(ruta)
    with sync_playwright() as p:
        nav, page = entrar(p, d)
        existe = page.request.get(URL + "/job/tareas-equipo/api/json").ok
        if existe:
            # Reanudación: el trabajo ya se creó desde "New Item"; se completa su configuración.
            page.goto(URL + "/job/tareas-equipo/configure", wait_until="networkidle")
        else:
            page.goto(URL + "/view/all/newJob", wait_until="networkidle")
            page.fill("#name", "tareas-equipo")
            page.locator("input[name=mode][value='org.jenkinsci.plugins.workflow.job.WorkflowJob']").locator("xpath=..").click()
            foto(page, "j11-nuevo-trabajo.png")
            page.click("#ok-button")
            page.wait_for_load_state("networkidle")
        try:
            page.wait_for_selector("textarea[name=description]", timeout=20_000)
        except Exception:
            page.screenshot(path=str(Path(sys.argv[2]).parent / "diagnostico_config.png"), full_page=True)
            print("selects:", [s.inner_text()[:60] for s in page.query_selector_all("select")])
            raise
        page.fill("textarea[name=description]",
                  "Integración continua de la API de tareas: pruebas, imagen, contenedores y verificación de /salud.")
        # Sección Pipeline: definición desde SCM, con Git.
        definicion = page.locator("select.dropdownList >> visible=true").filter(has_text="Pipeline script from SCM").first
        definicion.scroll_into_view_if_needed()
        definicion.select_option(label="Pipeline script from SCM")
        page.wait_for_timeout(1500)
        scm = page.locator("select >> visible=true").filter(has_text="Git").last
        scm.select_option(label="Git")
        page.wait_for_timeout(2000)
        page.locator("input[name='_.url'] >> visible=true").first.fill(REPO)
        page.locator("input[name='_.url'] >> visible=true").first.press("Tab")
        rama = page.locator("input[name='_.name'] >> visible=true").first
        rama.fill(RAMA)
        rama.press("Tab")
        page.locator("input[name='_.url'] >> visible=true").first.scroll_into_view_if_needed()
        page.evaluate("window.scrollBy(0, -120)")
        foto(page, "j12-configuracion-scm.png")
        rama.scroll_into_view_if_needed()
        page.evaluate("window.scrollBy(0, 200)")
        assert page.locator("input[name='_.scriptPath'] >> visible=true").first.input_value() == "Jenkinsfile"
        foto(page, "j12b-ruta-jenkinsfile.png")
        page.click("button[name=Submit]")
        page.wait_for_load_state("networkidle")
        print("trabajo guardado en", page.url)
        nav.close()


def trabajo_github(ruta):
    """Segundo trabajo: script en línea que descarga main desde GitHub (jenkins/tareas-equipo-github.groovy)."""
    d = datos(ruta)
    script = (RAIZ / "jenkins" / "tareas-equipo-github.groovy").read_text(encoding="utf-8")
    with sync_playwright() as p:
        nav, page = entrar(p, d)
        page.goto(URL + "/view/all/newJob", wait_until="networkidle")
        page.fill("#name", "tareas-equipo-github")
        page.locator("input[name=mode][value='org.jenkinsci.plugins.workflow.job.WorkflowJob']").locator("xpath=..").click()
        page.wait_for_timeout(1500)
        foto(page, "j11-nuevo-trabajo.png")
        page.click("#ok-button")
        page.wait_for_load_state("networkidle")
        page.wait_for_selector("textarea[name=description]", timeout=20_000)
        page.fill("textarea[name=description]", "Comprueba la conexión con GitHub: descarga main del repositorio público y corre las pruebas unitarias.")
        editor = page.locator(".ace_editor >> visible=true").first
        editor.scroll_into_view_if_needed()
        page.evaluate("s => { const e = ace.edit(document.querySelector('.ace_editor')); e.setValue(s, -1); }", script)
        page.wait_for_timeout(1000)
        page.evaluate("window.scrollBy(0, -150)")
        foto(page, "j16-script-github.png")
        page.click("button[name=Submit]")
        page.wait_for_load_state("networkidle")
        print("trabajo guardado en", page.url)
        nav.close()


if __name__ == "__main__":
    EVID.mkdir(exist_ok=True)
    {"credencial": credencial, "trabajo": trabajo, "github": trabajo_github}[sys.argv[1]](sys.argv[2])
