"""Genera docs/Entrega3_IC_Proyecto.docx (APA 7): el documento del proyecto con la primera parte (Docker, Entrega 1),
la segunda parte (Jenkins, Entrega 2) y la tercera parte (Travis CI, Codeship y GitHub Actions, historial, solución de
problemas, responsabilidades y opiniones, Entrega 3), a partir de las evidencias reales.

Uso:
    python tools/diagrama.py; python tools/diagrama_jenkins.py; python tools/diagrama_tres_ci.py   (diagramas)
    bash tools/capturar.sh            (evidencias de la primera parte; requiere Docker)
    bash tools/capturar_jenkins.sh ...   (evidencias de terminal de la segunda parte; ver el script)
    python tools/jenkins_capturas.py ... (capturas de la interfaz de Jenkins; requiere Playwright)
    python tools/web_capturas.py      (capturas de la documentación oficial de la tercera parte; requiere Playwright)
    python tools/validar_ci.py ...; python tools/simular_ci.py ...   (evidencias de terminal de la tercera parte)
    python build_documento.py         (arma el .docx; la tabla del historial sale de git log)
    powershell -File tools/exportar_pdf.ps1 -Nombre Entrega3_IC_Proyecto   (tabla de contenido y PDF con Word)

Los campos resaltados en amarillo, que empiezan por COMPLETAR, son datos que debe aportar el grupo.
"""
import re
import subprocess
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT  # noqa: F401
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_COLOR_INDEX, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

RAIZ = Path(__file__).resolve().parent
EVID = RAIZ / "evidencias"
SALIDA = RAIZ / "docs" / "Entrega3_IC_Proyecto.docx"
FUENTE = "Times New Roman"

# Estado del repositorio en GitHub. El texto del documento cambia según estos dos valores, para que siempre diga lo
# que de verdad está hecho. Cambiar a True solo después de comprobarlo con gh:
#   REPO_PUBLICO: gh repo view Owito/poli-ic-proyecto1 --json visibility  ->  "PUBLIC" (y recapturar 19-github)
#   COLABORADORES_INVITADOS: gh api repos/Owito/poli-ic-proyecto1/collaborators --jq '.[].login'  ->  lista a los integrantes
REPO_PUBLICO = True   # ✅ 6-nov-2026: gh repo view -> "PUBLIC"; curl sin sesión -> HTTP 200
COLABORADORES_INVITADOS = False

doc = Document()

# ---------------------------------------------------------------- estilos
sec = doc.sections[0]
sec.page_width, sec.page_height = Inches(8.5), Inches(11)
for lado in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(sec, lado, Inches(1))


def fijar_fuente(estilo, tam=12, negrita=None, cursiva=None):
    estilo.font.name = FUENTE
    estilo.font.size = Pt(tam)
    estilo.font.color.rgb = RGBColor(0, 0, 0)
    if negrita is not None:
        estilo.font.bold = negrita
    if cursiva is not None:
        estilo.font.italic = cursiva
    rpr = estilo.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(a), FUENTE)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(a)) is not None:
            del rfonts.attrib[qn(a)]


normal = doc.styles["Normal"]
fijar_fuente(normal)
normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
normal.paragraph_format.space_after = Pt(0)
normal.paragraph_format.space_before = Pt(0)

h1, h2, h3 = doc.styles["Heading 1"], doc.styles["Heading 2"], doc.styles["Heading 3"]
for h, neg, cur, al in ((h1, True, False, WD_ALIGN_PARAGRAPH.CENTER), (h2, True, False, WD_ALIGN_PARAGRAPH.LEFT),
                        (h3, True, True, WD_ALIGN_PARAGRAPH.LEFT)):
    fijar_fuente(h, 12, neg, cur)
    h.paragraph_format.alignment = al
    h.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    h.paragraph_format.space_before = Pt(0)
    h.paragraph_format.space_after = Pt(0)
    h.paragraph_format.keep_with_next = True

# ---------------------------------------------------------------- utilidades
MARCAS = re.compile(r"(\*\*.+?\*\*|\*.+?\*|\[\[.+?\]\])")


def rico(p, texto, base_cursiva=False):
    """*cursiva*, **negrita**, [[texto resaltado en amarillo: pendiente de confirmar]]."""
    for trozo in MARCAS.split(texto):
        if not trozo:
            continue
        if trozo.startswith("**"):
            r = p.add_run(trozo[2:-2]); r.bold = True
        elif trozo.startswith("[["):
            r = p.add_run(trozo[2:-2]); r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        elif trozo.startswith("*"):
            r = p.add_run(trozo[1:-1]); r.italic = True
        else:
            r = p.add_run(trozo)
        if base_cursiva:
            r.italic = True
    return p


def parrafo(texto, sangria=True, alin=None, simple=False, keep=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    if sangria:
        pf.first_line_indent = Inches(0.5)
    if alin is not None:
        p.alignment = alin
    if simple:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    if keep:
        pf.keep_with_next = True
    return rico(p, texto)


def titulo(texto, nivel):
    return doc.add_heading(texto, level=nivel)


def salto():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def campo(par, instruccion, ejemplo="1"):
    r = par.add_run()
    for tipo, extra in (("begin", None), (None, instruccion), ("separate", None), ("txt", ejemplo), ("end", None)):
        if tipo in ("begin", "separate", "end"):
            e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), tipo); r._r.append(e)
        elif tipo is None:
            e = OxmlElement("w:instrText"); e.set(qn("xml:space"), "preserve"); e.text = extra; r._r.append(e)
        else:
            e = OxmlElement("w:t"); e.text = extra; r._r.append(e)
    r.font.name = FUENTE
    return r


# número de página arriba a la derecha (APA 7)
hp = sec.header.paragraphs[0]
hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
campo(hp, "PAGE")

contador = {"fig": 0, "lis": 0, "tab": 0}
ORDEN = ["01-versiones", "02-estructura", "18-git-log", "19-github", "diagrama-arquitectura", "16-usuario-no-root", "03-compose-up",
         "04-compose-ps", "05-salud-db", "06-red", "11-dns-desde-app", "08-curl-salud", "09-curl-tareas", "10-psql",
         "12-crud", "13-psql-verifica", "14-validacion", "17-persistencia", "15-pruebas", "21-clon-github",
         # Segunda parte (Jenkins)
         "diagrama-jenkins", "j01-levantar", "j09-cliente-docker", "j02-desbloqueo", "j03-personalizar",
         "j04-instalando-plugins", "j05-primer-administrador", "j08-tablero-inicial", "j10-credencial",
         "j12-configuracion-scm", "j12b-ruta-jenkinsfile", "j15-etapas-verde", "j21-consola-verde", "j13-espejo",
         "j17-etapas-rojo", "j22-consola-rojo", "j20-etapas-historial", "j11-nuevo-trabajo", "j16-script-github",
         "j18-github-ejecucion", "j23-consola-github",
         # Tercera parte (Travis CI, Codeship y GitHub Actions)
         "t01-travis-prueba-gratuita", "t02-codeship-fin-de-vida", "t03-codeship-cierre", "diagrama-tres-ci",
         "t04-actions-gratis-publicos", "t05-validacion-ci", "t06-travis-local", "t07-actions-local"]
# Evidencias de la segunda parte que son salida de terminal (las demás "j" son capturas de la interfaz de Jenkins).
J_TERMINAL = {"j01-levantar", "j09-cliente-docker", "j13-espejo"}
J_CONSOLA = {"j21-consola-verde", "j22-consola-rojo", "j23-consola-github"}
# Notas propias de las figuras de la tercera parte que no son salida directa de un comando.
WEB = "tomada el 19 de noviembre de 2026 con un navegador sin ventana; se recortó la sección citada, sin modificarla."
EXTRACTO = ("Extracto de la salida real del comando, generada durante la elaboración de este trabajo; [...] marca las líneas "
            "omitidas. El registro completo está en el archivo evidencias/texto/{} del repositorio.")
NOTA_FIG = {
    "t01-travis-prueba-gratuita": "Captura de la página oficial https://docs.travis-ci.com/user/billing-overview/, " + WEB,
    "t02-codeship-fin-de-vida": ("Captura de la página a la que redirige https://docs.cloudbees.com/docs/cloudbees-codeship/latest/, " + WEB),
    "t04-actions-gratis-publicos": "Captura de la página oficial https://docs.github.com/en/actions/concepts/billing-and-usage, " + WEB,
    "t06-travis-local": EXTRACTO.format("t06-travis-local-completo.txt"),
    "t07-actions-local": EXTRACTO.format("t07-actions-local-completo.txt"),
}


def F(nombre):
    """Número de figura de una evidencia, según el orden en que aparece en el documento."""
    return ORDEN.index(nombre) + 1



def figura(png, titulo_fig, explicacion, ancho=6.5):
    """Figura APA 7: número en negrita, título en cursiva, imagen, nota y explicación."""
    contador["fig"] += 1
    n = contador["fig"]
    assert ORDEN[n - 1] == png[:-4], f"orden de figuras: {png} no es la figura {n}"
    p = doc.add_paragraph(); p.paragraph_format.keep_with_next = True
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    r = p.add_run(f"Figura {n}"); r.bold = True
    p2 = doc.add_paragraph(); p2.paragraph_format.keep_with_next = True
    p2.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p2.paragraph_format.space_after = Pt(6)
    r2 = p2.add_run(titulo_fig); r2.italic = True
    ruta = EVID / png
    if not ruta.exists():
        raise SystemExit(f"Falta la evidencia {ruta}. Ejecute tools/capturar.sh")
    pi = doc.add_paragraph(); pi.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pi.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pi.paragraph_format.keep_with_next = True
    pi.add_run().add_picture(str(ruta), width=Inches(ancho))
    pn = doc.add_paragraph(); pn.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pn.paragraph_format.space_after = Pt(10)
    rn = pn.add_run("Nota. "); rn.italic = True
    nombre = png[:-4]
    if png.startswith("diagrama"):
        nota = "Elaboración propia."
    elif nombre in NOTA_FIG:
        nota = NOTA_FIG[nombre]
    elif nombre in J_CONSOLA:
        nota = ("Extracto de la salida de consola real de la ejecución, leída con la API de Jenkins; se conservan solo "
                "las líneas clave y se ocultan el usuario y la clave.")
    elif nombre.startswith("j") and nombre not in J_TERMINAL:
        nota = ("Captura de pantalla de la interfaz web de Jenkins (http://localhost:8181), tomada durante la "
                "elaboración de este trabajo.")
    else:
        nota = "Captura de la salida real del comando, generada durante la elaboración de este trabajo."
    pn.add_run(nota)
    parrafo(f"**Qué muestra la Figura {n}.** {explicacion}")
    return n


def listado(titulo_lis, ruta_archivo, explicacion_previa=None):
    contador["lis"] += 1
    n = contador["lis"]
    p = doc.add_paragraph(); p.paragraph_format.keep_with_next = True
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    r = p.add_run(f"Listado {n}"); r.bold = True
    p2 = doc.add_paragraph(); p2.paragraph_format.keep_with_next = True
    p2.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p2.add_run(titulo_lis).italic = True
    contenido = (RAIZ / ruta_archivo).read_text(encoding="utf-8").rstrip("\n")
    for linea in contenido.split("\n"):
        pc = doc.add_paragraph()
        pc.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pc.paragraph_format.left_indent = Inches(0.2)
        rr = pc.add_run(linea if linea else " ")
        rr.font.name = "Consolas"; rr.font.size = Pt(9)
        rr._r.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Consolas")
        sh = OxmlElement("w:shd"); sh.set(qn("w:val"), "clear"); sh.set(qn("w:fill"), "F2F2F2")
        pc._p.get_or_add_pPr().append(sh)
    pn = doc.add_paragraph(); pn.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pn.paragraph_format.space_after = Pt(10)
    pn.add_run("Nota. ").italic = True
    pn.add_run(f"Contenido del archivo {ruta_archivo.replace(chr(92), '/')} del repositorio.")
    return n


def tabla(titulo_tab, encabezados, filas, anchos):
    contador["tab"] += 1
    n = contador["tab"]
    p = doc.add_paragraph(); p.paragraph_format.keep_with_next = True
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.add_run(f"Tabla {n}").bold = True
    p2 = doc.add_paragraph(); p2.paragraph_format.keep_with_next = True
    p2.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p2.paragraph_format.space_after = Pt(6)
    p2.add_run(titulo_tab).italic = True
    t = doc.add_table(rows=1, cols=len(encabezados))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = "Table Grid"
    for i, e in enumerate(encabezados):
        c = t.rows[0].cells[i]; c.text = ""
        pr = c.paragraphs[0]; pr.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pr.paragraph_format.keep_with_next = True  # el encabezado no queda solo al pie de una página
        pr.add_run(e).bold = True
    for fila in filas:
        cs = t.add_row().cells
        for i, v in enumerate(fila):
            cs[i].text = ""
            pr = cs[i].paragraphs[0]; pr.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            rico(pr, v)
    for k, fila in enumerate(t.rows):
        # Ninguna fila se parte entre dos páginas y el encabezado se repite si la tabla continúa en la página siguiente.
        trpr = fila._tr.get_or_add_trPr()
        trpr.append(OxmlElement("w:cantSplit"))
        if k == 0:
            trpr.append(OxmlElement("w:tblHeader"))
        for i, w in enumerate(anchos):
            fila.cells[i].width = Inches(w)
        for c in fila.cells:
            for pr in c.paragraphs:
                for r in pr.runs:
                    r.font.size = Pt(10.5)
    sp = doc.add_paragraph(); sp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    return n


def centrado(texto, negrita=False, resaltar=False, simple=False):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if simple:
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    r = p.add_run(texto); r.bold = negrita
    if resaltar:
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return p


# ================================================================ PORTADA
for _ in range(3):
    doc.add_paragraph()
centrado("Proyecto de software basado en herramientas de integración continua: API de tareas del equipo con Docker, Jenkins, "
         "Travis CI y GitHub Actions", negrita=True)
centrado("Entrega 3 y sustentación, escenarios 7 y 8")
doc.add_paragraph()
centrado("Leonardo Espinosa Rivera")
centrado("Daniel Santiago Estupiñán Cruz")
centrado("Anyelo Brayan Galindo Delgado")
centrado("Jerlinson Fabián Ramos Galíndez")
centrado("Carlos Fernando Guerra Arango")
doc.add_paragraph()
centrado("Programa de Ingeniería de Software")
centrado("Facultad de Ingeniería, Diseño e Innovación")
centrado("Politécnico Grancolombiano")
centrado("TIC 42711: Integración Continua, grupo B04")
centrado("Grupo de trabajo colaborativo 14")
centrado("Tutor: Edward Reyes Corredor")
centrado("Diciembre de 2026")

# ================================================================ TABLA DE CONTENIDO
# Título con formato directo (no Heading 1) para que el campo TOC no lo liste; el salto de página va en el propio
# párrafo para no dejar una hoja en blanco cuando la portada llena la página 1.
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.page_break_before = True
p.add_run("Tabla de contenido").bold = True
ptoc = doc.add_paragraph(); ptoc.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
campo(ptoc, 'TOC \\o "1-3" \\h \\z \\u', "Haga clic derecho y elija Actualizar campo para generar la tabla de contenido.")
salto()

# ================================================================ INTRODUCCIÓN
titulo("Introducción", 1)
parrafo("La integración continua es una práctica del desarrollo de software en la que los integrantes de un equipo "
        "integran su trabajo con frecuencia y cada integración se verifica de forma automática, de modo que los errores se "
        "detecten pronto (Reyes Corredor, 2026). Para que esa práctica sea posible, el proyecto "
        "debe poder construirse y ejecutarse de la misma manera en cualquier equipo. Docker resuelve ese punto: empaqueta la "
        "aplicación y sus dependencias en contenedores que se comportan igual en el computador de cada integrante y, más "
        "adelante, en un servidor de integración continua.")
parrafo("A diferencia de una máquina virtual, que instala un sistema operativo completo sobre un hipervisor, Docker aprovecha el núcleo del sistema "
        "operativo del equipo anfitrión y empaqueta solo la aplicación con sus librerías, su configuración y sus dependencias, lo que permite que funcione "
        "igual en cualquier computador (Reyes Corredor, 2026).")
parrafo("Este documento reúne las tres entregas del proyecto del módulo. Como indicó el tutor, cada entrega se agrega como una "
        "parte nueva del mismo documento, de modo que al final todo el proyecto quede en uno solo (Reyes Corredor, 2026). La primera "
        "parte corresponde a la Entrega 1: el enunciado indica que el proyecto se crea en un repositorio de GitHub y que, con Docker, se "
        "construyen dos contenedores comunicados entre sí (Politécnico Grancolombiano, s. f.). Para cumplirlo se desarrolló una API REST de "
        "registro de tareas del equipo con Node.js y Express, que guarda sus datos en una base de datos PostgreSQL. La API corre en un "
        "contenedor y la base de datos en otro; ambos se orquestan con Docker Compose. El dominio es deliberadamente sencillo, porque las "
        "entregas siguientes necesitan una aplicación con pruebas automatizadas sobre la cual construir los flujos de integración continua.")
parrafo("La segunda parte corresponde a la Entrega 2. El enunciado pide aplicar la integración continua implementando Jenkins como gestor "
        "de operaciones y entregar un documento con las características requeridas para su implementación (Politécnico Grancolombiano, "
        "s. f.). Jenkins se instaló con Docker, se configuró desde su interfaz web y se conectó con el repositorio mediante un pipeline "
        "escrito en un archivo Jenkinsfile que, ante cada cambio, ejecuta las pruebas, construye la imagen de la API, levanta los dos "
        "contenedores y verifica que la API responda. El repositorio del proyecto está disponible en https://github.com/Owito/poli-ic-proyecto1.")
parrafo("La tercera parte corresponde a la Entrega 3. El enunciado pide que la plataforma quede totalmente integrada con contenedores, "
        "Jenkins, Travis CI y Codeship, y que el documento consolide el historial de cambios, las sugerencias para la solución de "
        "problemas y las responsabilidades y opiniones (Politécnico Grancolombiano, s. f.). Se agregaron las configuraciones de Travis CI "
        "y de GitHub Actions, que ejecutan las mismas etapas que Jenkins; GitHub Actions reemplaza a Codeship, que CloudBees retiró en "
        "enero de 2026, como se explica y se evidencia en esa parte.")
parrafo("Cada captura de pantalla del documento es la salida real de un comando, una vista real de la interfaz de Jenkins o una página "
        "oficial de documentación, obtenida durante la elaboración del trabajo, y va acompañada de una explicación de lo que muestra y de "
        "por qué importa para el proyecto.")
titulo("Objetivo general", 2)
parrafo("Construir una aplicación de dos contenedores, una API en Node.js y una base de datos PostgreSQL, y automatizar su integración "
        "continua con Jenkins, Travis CI y GitHub Actions, de modo que cada cambio en el repositorio se pruebe, se construya y se "
        "verifique sin pasos manuales, documentando el proceso con evidencias.")
titulo("Objetivos específicos", 2)
parrafo("Primero, organizar el proyecto en un repositorio de control de versiones con una estructura clara, publicarlo en GitHub y compartirlo con los integrantes del grupo y con el tutor.")
parrafo("Segundo, definir un Dockerfile para la API y un archivo de Docker Compose que levante los dos servicios en una red propia, con volumen "
        "persistente para la base de datos, comprobación de salud y orden de arranque.")
parrafo("Tercero, demostrar con comandos reales que los dos contenedores se comunican: que la API consulta la base de datos y que los datos "
        "devueltos provienen de ella.")
parrafo("Cuarto, escribir pruebas automatizadas que puedan ejecutarse desde las herramientas de integración continua.")
parrafo("Quinto, instalar y configurar Jenkins con Docker, e identificar las características y los requisitos que exige su implementación en el proyecto.")
parrafo("Sexto, definir en un Jenkinsfile, versionado junto con el código, un pipeline que descargue el proyecto, ejecute las pruebas, construya "
        "la imagen, levante los contenedores y verifique la API.")
parrafo("Séptimo, comprobar que el pipeline se ejecuta solo cuando cambia el repositorio y que detiene la integración cuando una prueba falla.")
parrafo("Octavo, integrar el proyecto con Travis CI y con un servicio equivalente a Codeship, después de verificar el estado de cada "
        "servicio en su fuente oficial, de modo que ambos ejecuten las mismas etapas que Jenkins.")
parrafo("Noveno, consolidar el historial de cambios, las sugerencias para la solución de problemas y las responsabilidades y opiniones "
        "de los integrantes.")

# ================================================================ PRIMERA PARTE
salto()
titulo("Primera parte: el proyecto en dos contenedores con Docker", 1)
titulo("Repositorio en GitHub y estructura del proyecto", 2)
titulo("Elección del proyecto y de las herramientas", 3)
parrafo("El tutor indicó que cada grupo escoge libremente el proyecto, siempre que sea proporcional y permita evidenciar que las herramientas del módulo se "
        "integran, y que se prefieren herramientas libres, instalando solo las necesarias para el proyecto (Reyes Corredor, 2026). "
        "Por eso se eligió un registro de tareas del equipo: tiene un modelo de datos de una sola tabla, un CRUD completo y suficiente "
        "lógica de validación para escribir pruebas útiles. Todas las herramientas usadas son libres o de uso gratuito: Git para el control de versiones, "
        "GitHub como servicio del repositorio remoto, Docker y Docker Compose para los contenedores, Node.js 22 con el marco Express para la API "
        "(Express, s. f.) y PostgreSQL 16 como base de datos (PostgreSQL Global Development Group, s. f.). Node.js y Git se instalan en el equipo "
        "porque se usan para escribir el código y las pruebas; PostgreSQL no se instala en el equipo, solo corre dentro de su contenedor. "
        f"Como indicó el tutor, no se documenta el paso detallado de la instalación de Docker y de Git; la Figura {F("01-versiones")} deja constancia de que están "
        "instalados y de las versiones con las que se trabajó (Reyes Corredor, 2026).")
figura("01-versiones.png", "Versiones de las herramientas utilizadas",
       "El comando imprime las versiones de Docker, Docker Compose, Git y Node.js instaladas en el equipo donde se construyó el proyecto. "
       "Sirve para que cualquier persona que reproduzca el trabajo sepa con qué versiones se verificó, y se relaciona con la recomendación del tutor de "
       "que todos trabajen con la misma versión de las herramientas.")
titulo("Estructura del repositorio", 3)
parrafo(f"El código se organiza de manera que cada archivo tenga una sola responsabilidad. La Figura {F("02-estructura")} lista los archivos versionados con Git, "
        "y la Tabla 1 explica para qué sirve cada uno.", keep=True)
figura("02-estructura.png", "Archivos versionados en el repositorio",
       "Es la salida de git ls-files, es decir, solo los archivos que Git está siguiendo. Se observa que el archivo .env con los valores reales "
       "no aparece en la lista, mientras que .env.example sí: las claves de la base de datos no se publican, solo su plantilla.")
tabla("Archivos principales del proyecto", ["Archivo", "Función"], [
    ["Dockerfile", "Receta para construir la imagen de la API en dos etapas, con usuario sin privilegios."],
    ["docker-compose.yml", "Define los servicios app y db, la red red-tareas, el volumen tareas-datos, el healthcheck y el orden de arranque."],
    [".env.example", "Plantilla de variables de entorno sin secretos reales; cada integrante la copia a .env."],
    ["db/init.sql", "Crea la tabla tareas y carga cinco filas de ejemplo la primera vez que se crea el volumen."],
    ["src/app.js, db.js, server.js", "Rutas de la API, conexión a PostgreSQL y arranque del servidor."],
    ["tests/", "Pruebas unitarias (base de datos simulada) y de integración (base de datos real)."],
    [".dockerignore, .gitignore", "Evitan que node_modules, .env y las evidencias entren a la imagen o al repositorio."],
    ["README.md", "Instrucciones en español para levantar el proyecto."],
], [2.0, 4.5])
titulo("Control de versiones y publicación en GitHub", 3)
parrafo(f"El proyecto se desarrolló en un repositorio Git con commits pequeños y mensajes descriptivos que siguen la convención de tipos "
        f"(feat, docs, test). La Figura {F("18-git-log")} muestra el historial local con los archivos que cambió cada commit, un tipo de evidencia que el tutor mencionó como valiosa "
        "para ver el avance del equipo (Reyes Corredor, 2026).", keep=True)
figura("18-git-log.png", "Historial de commits del repositorio",
       "Cada línea de cabecera es un commit con su identificador corto, autor, fecha y mensaje, seguida de los archivos que modificó. Confirma que el trabajo "
       "quedó registrado por etapas: primero la API, luego los contenedores y después la documentación.")
parrafo("El repositorio remoto se creó el 5 de noviembre de 2026 con la herramienta de línea de comandos de GitHub, gh, y después se enlazó con el "
        "repositorio local para subir el historial. La Tabla 2 resume ese paso a paso. El mismo resultado se obtiene desde la página web de GitHub, con la "
        "opción New repository del menú que abre el signo más, arriba a la derecha, como lo mostró el tutor en el encuentro del 5 de noviembre "
        "(Reyes Corredor, 2026).", keep=True)
PASOS_REPO = [
    ["1", "gh repo create Owito/poli-ic-proyecto1 --private --description \"Proyecto de Integración Continua: API de tareas en Node.js y PostgreSQL con Docker Compose\"",
     "Crea el repositorio vacío en la cuenta de GitHub, con su descripción."],
    ["2", "git remote add origin https://github.com/Owito/poli-ic-proyecto1.git",
     "Registra el repositorio de GitHub como remoto origin del repositorio local."],
    ["3", "git push -u origin HEAD:main",
     "Sube los commits a la rama main de GitHub y deja configurado el seguimiento entre la rama local y la remota."],
]
if REPO_PUBLICO:
    PASOS_REPO.append([str(len(PASOS_REPO) + 1), "gh repo edit Owito/poli-ic-proyecto1 --visibility public --accept-visibility-change-consequences",
                       "Cambia la visibilidad a pública, como recomendó el tutor, para que pueda entrar con el enlace."])
if COLABORADORES_INVITADOS:
    PASOS_REPO.append([str(len(PASOS_REPO) + 1), "Settings, Collaborators, Add people (página web de GitHub)",
                       "Invita a cada integrante como colaborador; el acceso queda activo cuando el integrante acepta la invitación."])
tabla("Pasos para crear el repositorio en GitHub y publicar el proyecto", ["Paso", "Comando", "Qué hace"], PASOS_REPO, [0.5, 3.2, 2.8])
if REPO_PUBLICO:
    parrafo("Los repositorios de GitHub se describen en su documentación oficial (GitHub, s. f.-e). El proyecto está publicado en un repositorio público de "
            "GitHub, https://github.com/Owito/poli-ic-proyecto1, enlazado como remoto origin del repositorio local. Se creó como privado y se cambió a público "
            "por recomendación del tutor, para que él pueda entrar a validarlo y para que las herramientas de integración continua muestren los errores donde "
            f"los encuentren (Reyes Corredor, 2026). La Figura {F("19-github")} muestra la consulta del repositorio publicado y su historial en GitHub.", keep=True)
    figura("19-github.png", "Repositorio publicado en GitHub",
           "La primera consulta confirma el nombre, la visibilidad pública, la rama principal main y la dirección del repositorio. La segunda muestra que la "
           "rama main de GitHub apunta al mismo commit que el repositorio local, y la tercera lista el historial tal como quedó publicado. La cuarta pide la "
           "página del repositorio sin iniciar sesión y recibe la respuesta HTTP 200: cualquier persona con el enlace, incluido el tutor, puede abrirlo.")
else:
    parrafo("Los repositorios de GitHub se describen en su documentación oficial (GitHub, s. f.-e). El proyecto se publicó en un repositorio privado de "
            f"GitHub, https://github.com/Owito/poli-ic-proyecto1, enlazado como remoto origin del repositorio local. La Figura {F("19-github")} muestra la "
            "consulta del repositorio publicado y su historial en GitHub.", keep=True)
    figura("19-github.png", "Repositorio publicado en GitHub",
           "La primera consulta confirma el nombre, la visibilidad privada, la rama principal main y la dirección del repositorio. La segunda muestra que la "
           "rama main de GitHub apunta al mismo commit que el repositorio local, y la tercera lista el historial tal como quedó publicado.")
ACCESO = ("Para revisar el proyecto, el tutor pidió entrar al repositorio y ver cómo están conformados el repositorio y los contenedores. En el encuentro del "
          "5 de noviembre de 2026 aclaró que no comparte su usuario de GitHub y que acepta tres alternativas: crearle un usuario, dejar el repositorio público "
          "o enviarle un enlace que le permita ingresar (Reyes Corredor, 2026). ")
if REPO_PUBLICO:
    ACCESO += "Con el repositorio público, el enlace anterior basta para que el tutor lo consulte sin cuenta ni invitación."
else:
    ACCESO += ("[[DECISIÓN DEL GRUPO: el tutor recomendó dejarlo público; al cambiar la visibilidad, poner REPO_PUBLICO = True en "
               "build_documento.py, recapturar la Figura de GitHub y regenerar.]]")
if COLABORADORES_INVITADOS:
    ACCESO += (" Los integrantes del grupo, en cambio, se registraron como colaboradores en Settings, Collaborators, la opción que el tutor "
               "indicó para conformar el equipo de desarrollo y que GitHub documenta en su procedimiento de invitación (GitHub, s. f.-d).")
else:
    ACCESO += (" Para que cada integrante pueda subir sus propios cambios, se le registra como colaborador en Settings, Collaborators con su "
               "usuario de GitHub, la opción que el tutor indicó para conformar el equipo de desarrollo y que GitHub documenta en su "
               "procedimiento de invitación (GitHub, s. f.-d).")
parrafo(ACCESO)

titulo("Arquitectura de los dos contenedores", 2)
parrafo("La solución tiene dos servicios que Docker Compose crea y conecta en una misma red. El primero, app, ejecuta la API de Node.js. "
        "El segundo, db, ejecuta PostgreSQL 16. Solo la API se publica hacia el equipo anfitrión, por el puerto 8095; la base de datos queda "
        "accesible únicamente dentro de la red, lo que reduce la superficie expuesta.", keep=True)
figura("diagrama-arquitectura.png", "Diagrama de la arquitectura de la solución",
       "El cliente (navegador o curl) habla por HTTP con el contenedor tareas-app, que atiende en el puerto interno 3000 y está publicado en el 8095 del equipo. "
       "Ese contenedor se conecta por SQL con tareas-db usando el nombre de servicio db como dirección. Ambos viven en la red red-tareas. La base de datos "
       "guarda sus archivos en el volumen tareas-datos y recibe el script db/init.sql en el primer arranque. La flecha de depends_on indica que la API no arranca hasta que la base "
       "de datos pase su comprobación de salud.")
parrafo(f"La Tabla {contador['tab'] + 1} resume la imagen, los puertos y la responsabilidad de cada contenedor.", keep=True)
tabla("Los dos contenedores del proyecto", ["Servicio", "Contenedor", "Imagen", "Puerto", "Responsabilidad"], [
    ["app", "tareas-app", "Construida desde el Dockerfile (Node.js 22 sobre Alpine)", "3000 interno, 8095 publicado", "API REST de tareas y endpoint /salud"],
    ["db", "tareas-db", "postgres:16-alpine", "5432 solo dentro de la red", "Almacena las tareas de forma persistente"],
], [0.8, 1.0, 1.9, 1.3, 1.5])

titulo("El Dockerfile explicado", 2)
parrafo("El Dockerfile es la receta con la que Docker construye la imagen de la API (Docker Inc., s. f.-c). Se escribió en dos etapas, como muestra el Listado 1.", keep=True)
listado("Dockerfile de la API", "Dockerfile")
parrafo("En la primera etapa, llamada dependencias, se copian solo package.json y package-lock.json y se ejecuta npm ci --omit=dev. Así se instalan "
        "exactamente las versiones bloqueadas y únicamente las dependencias de producción. Como Docker guarda cada instrucción en una capa, mientras esos dos archivos no "
        "cambien, la instalación se reutiliza de la memoria de construcción y las reconstrucciones son rápidas.")
parrafo(f"La segunda etapa parte de una imagen limpia de node:22-alpine y copia desde la primera únicamente la carpeta node_modules, junto con el código de src y tests. "
        "El resultado es una imagen final pequeña, sin herramientas de construcción. La instrucción USER node hace que el proceso corra con un usuario sin privilegios, en lugar de root, "
        f"una práctica de seguridad que limita el daño si la aplicación llegara a ser vulnerada. La Figura {F("16-usuario-no-root")} lo comprueba dentro del contenedor en ejecución.")
figura("16-usuario-no-root.png", "Usuario con el que corre la API dentro del contenedor",
       "El comando id, ejecutado dentro del contenedor tareas-app, devuelve uid=1000(node): el proceso no es root. El segundo comando confirma que la versión de Node.js del "
       "contenedor es la 22, la que pide el proyecto.")
parrafo("El archivo .dockerignore evita que node_modules, el repositorio .git, el archivo .env y las evidencias se envíen al motor de Docker al construir. Con ello el contexto "
        "de construcción es pequeño y los secretos locales no terminan dentro de la imagen.")

titulo("El archivo docker-compose.yml explicado", 2)
parrafo("Docker Compose permite describir varios contenedores y sus relaciones en un solo archivo (Docker Inc., s. f.-a; Docker Inc., s. f.-b). "
        "El Listado 2 contiene el archivo del proyecto.", keep=True)
listado("Archivo docker-compose.yml", "docker-compose.yml")
parrafo("Los puntos clave son los siguientes. **Red definida:** ambos servicios se conectan a red-tareas, una red de tipo bridge creada por el propio archivo. "
        "En una red definida por el usuario, Docker resuelve el nombre de cada servicio como dirección, por eso la API usa DB_HOST=db (Docker Inc., s. f.-d). "
        "**Volumen:** tareas-datos se monta en /var/lib/postgresql/data, de modo que las tareas sobreviven aunque se borren los contenedores (Docker Inc., s. f.-f). "
        "**Script de inicialización:** el archivo db/init.sql se monta en la carpeta docker-entrypoint-initdb.d, y la imagen oficial de PostgreSQL lo ejecuta una sola vez, "
        "cuando el volumen está vacío (Docker Inc., s. f.-e). **Healthcheck:** la base de datos se declara saludable cuando pg_isready responde, con una prueba cada cinco segundos. "
        "**Orden de arranque:** depends_on con condition: service_healthy hace que la API espere a que la base de datos esté lista, evitando errores de conexión al inicio. "
        "**Variables de entorno:** el usuario, la clave y el nombre de la base se leen del archivo .env, que no se sube al repositorio; el repositorio solo contiene .env.example.")
parrafo(f"La Figura {F("03-compose-up")} muestra el arranque real. Docker construye la imagen, crea la red y el volumen, inicia primero la base de datos, espera a que sea saludable y solo entonces inicia la API.", keep=True)
figura("03-compose-up.png", "Construcción y arranque con docker compose up",
       "Se ven los últimos pasos de la construcción de la imagen tareas-equipo-app y, a continuación, la creación de la red tareas-equipo_red-tareas, del volumen "
       "tareas-equipo_tareas-datos y de los dos contenedores. El orden de las líneas es la prueba de la dependencia: tareas-db pasa por Waiting y luego Healthy, y recién después "
       "tareas-app pasa a Starting y Started.")
figura("04-compose-ps.png", "Estado de los dos contenedores",
       "docker compose ps lista los dos servicios del proyecto. tareas-db aparece con estado healthy, lo que significa que su healthcheck está pasando, y tareas-app aparece en ejecución "
       "con el puerto 8095 del equipo redirigido al 3000 del contenedor. La base de datos solo muestra su puerto interno 5432, sin publicarlo hacia afuera.")
parrafo(f"La Figura {F("04-compose-ps")} también permite distinguir imagen y contenedor. La imagen es la plantilla a partir de la cual se crea el contenedor; "
        "el tutor la comparó con el plano de una casa, y al contenedor con la casa ya construida (Reyes Corredor, 2026). En este proyecto hay dos imágenes, "
        "tareas-equipo-app, construida desde el Dockerfile, y postgres:16-alpine, descargada de Docker Hub, y de cada una se crea un contenedor: tareas-app "
        "y tareas-db.")
figura("05-salud-db.png", "Estado de salud reportado por Docker",
       "docker inspect lee directamente el estado de salud que Docker registra para tareas-db: running y healthy. Es la misma información que usa depends_on para decidir cuándo arrancar la API.")

titulo("Comunicación entre los contenedores", 2)
parrafo("Que los dos contenedores estén en ejecución no prueba por sí solo que se comuniquen. Por eso se verificó la comunicación desde cuatro ángulos: la red de Docker, la "
        "resolución de nombres, una respuesta de la API que contiene datos de la base y la comprobación cruzada directamente contra la base de datos.")
titulo("Red y resolución de nombres", 3)
figura("06-red.png", "Contenedores conectados a la red red-tareas",
       "docker network inspect lista los miembros de la red tareas-equipo_red-tareas, de tipo bridge. Aparecen tareas-app y tareas-db, cada uno con su dirección IP "
       "dentro de la misma subred. Estar en la misma red es la condición para que puedan hablar entre sí.")
figura("11-dns-desde-app.png", "Resolución del nombre db desde el contenedor de la API",
       f"Se ejecuta un pequeño programa de Node.js dentro de tareas-app que pregunta por la dirección del nombre db. La respuesta coincide con la IP que la Figura {F("06-red")} asigna a tareas-db. "
       "Es la prueba de que la API no necesita conocer direcciones numéricas: le basta con el nombre del servicio.")
titulo("La API consulta la base de datos", 3)
figura("08-curl-salud.png", "Respuesta del endpoint /salud",
       "El endpoint /salud ejecuta una consulta en la base de datos y devuelve su nombre (tareas) y la versión de PostgreSQL (16.x). Esa información no existe en la API: solo puede provenir "
       "del otro contenedor. Si la base dejara de responder, el endpoint devolvería HTTP 503 con estado error, comportamiento cubierto por una de las pruebas automatizadas.")
figura("09-curl-tareas.png", "Listado de tareas devuelto por la API",
       "La API devuelve en formato JSON las cinco tareas de ejemplo que el script db/init.sql cargó en la base de datos al crearse el volumen. Los datos del navegador o de curl viajan por HTTP "
       "hasta tareas-app y esta los obtiene de tareas-db con una consulta SQL.")
titulo("Comprobación cruzada contra la base de datos", 3)
figura("10-psql.png", "Consulta directa a PostgreSQL con psql",
       f"Con docker compose exec se abre psql dentro del contenedor tareas-db y se consulta la tabla tareas. Las cinco filas coinciden con las que entregó la API en la Figura {F("09-curl-tareas")}, "
       "lo que confirma que la API lee de esta base de datos y no de datos escritos en su propio código.")
figura("12-crud.png", "Creación y actualización de una tarea por la API",
       "Con curl se envía un POST que crea la tarea 6 y un PUT que cambia su estado a en_progreso. La API responde con el registro completo, incluida la fecha creada_en que asigna la base de datos.")
figura("13-psql-verifica.png", "La tarea creada por la API existe en la base de datos",
       f"Una consulta con psql a la fila 6 devuelve la tarea creada en la Figura {F("12-crud")} con su estado actualizado. Es la comprobación de ida y vuelta: lo que la API escribe queda guardado en el contenedor de la base de datos.")
figura("14-validacion.png", "Respuestas de error de la API",
       "Un POST sin título responde HTTP 400 y la consulta de una tarea inexistente responde HTTP 404. La validación ocurre en la API antes de tocar la base, y las rutas inexistentes se manejan con códigos HTTP correctos.")
titulo("Persistencia de los datos", 3)
parrafo("El volumen tareas-datos guarda los datos fuera del ciclo de vida de los contenedores. Para comprobarlo se ejecutó docker compose down, que borra los contenedores y la red pero no el volumen, "
        "y luego docker compose up -d.", keep=True)
figura("17-persistencia.png", "Los datos sobreviven a la recreación de los contenedores",
       "Se ven la eliminación y la nueva creación de los dos contenedores y de la red. Después de que tareas-db vuelve a estar healthy, la consulta a la tarea 6, creada antes del reinicio, "
       "sigue respondiendo. Según la documentación de la imagen oficial, el script de inicialización solo se ejecuta cuando el volumen está vacío, por lo que no se vuelve a ejecutar en este reinicio (Docker Inc., s. f.-e).")

titulo("Pruebas automatizadas", 2)
parrafo("Las pruebas usan el ejecutor integrado de Node.js, node:test, que no requiere instalar dependencias adicionales (Node.js, s. f.). Hay dos grupos. Las pruebas unitarias "
        "reemplazan la base de datos por un objeto simulado y verifican la validación de entradas (HTTP 400), los casos de recurso inexistente (HTTP 404) y el comportamiento de /salud cuando la base "
        "falla (HTTP 503). La prueba de integración usa la base de datos real y recorre el ciclo completo: salud, listado de las tareas de ejemplo, creación, actualización, borrado y verificación del borrado. "
        "La prueba de integración se omite automáticamente cuando no existe la variable DB_HOST, de modo que el comando npm test también funciona en un equipo sin Docker.")
figura("15-pruebas.png", "Ejecución de las pruebas dentro del contenedor de la API",
       "Dentro de tareas-app, donde la variable DB_HOST apunta a la base de datos, las cinco pruebas se ejecutan y pasan, sin pruebas omitidas, incluida la de integración contra el contenedor tareas-db. "
       "El resumen final reporta cinco pruebas, cinco aprobadas y cero fallidas.")
parrafo("Estas mismas pruebas son las que ejecuta Jenkins en la segunda parte de este documento, y las que ejecutan Travis CI y GitHub Actions en la "
        "tercera, ya que el comando npm test no depende de ninguna herramienta específica.")

titulo("Conexión entre GitHub y Docker", 2)
parrafo("El tutor precisó que el alcance de esta entrega es trabajar con GitHub, construir los dos contenedores con Docker y conectar Docker con GitHub "
        "(Reyes Corredor, 2026). En este proyecto la conexión se da a través de los archivos que definen los contenedores. GitHub no guarda las imágenes ni "
        "los contenedores, sino las instrucciones para construirlos: el Dockerfile de la API, el archivo docker-compose.yml con los dos servicios, la red y el "
        f"volumen, el script db/init.sql y la plantilla .env.example, todos visibles en la Figura {F("02-estructura")}. A partir de esos archivos, Docker construye "
        "la imagen de la API en el equipo de quien ejecuta el proyecto y descarga de Docker Hub la imagen oficial postgres:16-alpine (Docker Inc., s. f.-e).")
parrafo("Por eso el repositorio es el punto de encuentro del equipo: cualquier integrante, o el tutor, clona el repositorio y ejecuta docker compose up -d --build "
        "para obtener los mismos dos contenedores, con la misma red y la misma configuración. Cuando alguien modifica el Dockerfile o el docker-compose.yml, el "
        "cambio queda en el historial de commits y los demás lo reciben al actualizar su copia, lo que evita que cada integrante trabaje con una versión distinta "
        "del ambiente. Es también la conexión que aprovechan Jenkins en la segunda parte, y Travis CI y GitHub Actions en la tercera, "
        "que toman el código del repositorio para construir los contenedores y ejecutar las pruebas.")
parrafo(f"Para comprobarlo, el proyecto se clonó desde el repositorio público en una carpeta nueva, sin iniciar sesión en GitHub, y se levantó "
        f"con otro nombre de proyecto de Docker Compose, tareas-clon, para no reutilizar el volumen del ambiente original. La Figura {F("21-clon-github")} "
        "muestra el resultado.", keep=True)
figura("21-clon-github.png", "El proyecto clonado desde GitHub levanta los mismos dos contenedores",
       "El comando git clone descarga el repositorio público; después se copia la plantilla .env.example como .env y se levanta el proyecto. Docker "
       "construye la imagen de la API a partir del Dockerfile del repositorio y descarga postgres:16-alpine: los dos contenedores quedan saludables, el "
       "endpoint /salud responde con la versión de PostgreSQL leída de la base de datos y las cinco pruebas automatizadas pasan dentro del contenedor de la "
       "API, sin pruebas omitidas. Es la evidencia de que lo que está en GitHub basta para reconstruir el ambiente completo.")

titulo("Cómo reproducir el proyecto", 2)
parrafo("El archivo README.md del repositorio lo explica en español. En resumen: se clona el repositorio con git clone https://github.com/Owito/poli-ic-proyecto1.git, "
        "se copia .env.example como .env y se cambia la clave, se ejecuta docker compose up -d --build, se verifica con docker compose ps "
        "y se consultan http://localhost:8095/salud y http://localhost:8095/tareas. Para detenerlo se usa docker compose down.")

# ================================================================ SEGUNDA PARTE (ENTREGA 2)
def sig(tipo):
    """Número que tendrá la próxima figura, tabla o listado (para citarlo antes de que aparezca)."""
    return contador[tipo] + 1


salto()
titulo("Segunda parte: integración continua con Jenkins", 1)
parrafo("El enunciado de la Entrega 2 dice que es momento de aplicar la integración continua, y que para ello se implementa Jenkins como "
        "gestor de operaciones, acompañado de un documento con las características requeridas para su implementación (Politécnico "
        "Grancolombiano, s. f.). El tutor precisó el alcance: saber qué hace la herramienta, qué características tiene y cómo se implementa "
        "en el proyecto que se viene desarrollando, explorándola e instalándola, y explicar el paso a paso para demostrar que se entiende, no "
        "solo pegar capturas (Reyes Corredor, 2026). Esta parte responde a esos tres puntos sobre la misma API de tareas de la primera parte.")

titulo("Qué es Jenkins y qué papel cumple en el proyecto", 2)
parrafo("Jenkins es un servidor de automatización de código abierto. Su sitio oficial lo describe como un programa autónomo escrito en Java que "
        "puede funcionar como un servidor de integración continua sencillo o convertirse en el centro de entrega continua de un proyecto, y que "
        "se configura desde una interfaz web (Jenkins, s. f.-c). Sobre esa base, Jenkins Pipeline es un conjunto de plugins que permite definir "
        "todo el proceso de integración y entrega como código, en un archivo de texto llamado Jenkinsfile que se guarda en el control de "
        "versiones junto con la aplicación (Jenkins, s. f.-d).")
parrafo("En este proyecto Jenkins cumple el papel de gestor de operaciones que pide el enunciado: cada vez que alguien del equipo integra un "
        "cambio en la rama del proyecto, Jenkins lo descarga y ejecuta, sin intervención de nadie, la misma secuencia de verificación: "
        "pruebas unitarias, construcción de la imagen de la API, arranque de los dos contenedores, consulta del endpoint /salud y pruebas de "
        "integración contra la base de datos real. Es el ciclo de código, commit, construcción y despliegue que el tutor asoció con la cultura "
        "DevOps (Reyes Corredor, 2026). Si alguna etapa falla, la integración se detiene y el equipo sabe exactamente dónde está el problema.")

titulo("Características de Jenkins", 2)
parrafo(f"El tutor resumió las características de Jenkins en el encuentro del escenario 4: es gratuito, sirve para integración y despliegue "
        f"continuos, trabaja sobre la línea de código, commit, construcción y despliegue, cuenta con más de 1000 plugins, guarda el historial de los "
        f"cambios y se integra con herramientas como Docker y con la nube (Reyes Corredor, 2026). La Tabla {sig('tab')} contrasta esas características "
        f"con la documentación oficial y explica cómo se aprovechó cada una en el proyecto.", keep=True)
tabla("Características de Jenkins y su uso en el proyecto", ["Característica", "Qué significa", "Cómo se aprovecha en el proyecto"], [
    ["Integración y entrega continuas", "Puede funcionar como servidor de integración continua o como centro de entrega continua (Jenkins, s. f.-c).",
     "Ejecuta el pipeline del proyecto cada vez que cambia la rama de trabajo."],
    ["Instalación sencilla", "Es un programa Java autónomo, con paquetes para Windows, Linux y macOS (Jenkins, s. f.-c) y una imagen oficial de Docker (Jenkins, s. f.-a).",
     "Se instaló como contenedor, sin instalar Java en el equipo."],
    ["Configuración por interfaz web", "Se configura desde el navegador, con verificación de errores y ayuda en línea (Jenkins, s. f.-c).",
     "El asistente inicial, la credencial y los trabajos se configuraron desde la interfaz."],
    ["Plugins", "El sitio oficial reporta más de 2300 plugins aportados por la comunidad (Jenkins, s. f.-c).",
     "Se usan Pipeline, Git, Docker Pipeline, Credentials Binding, Timestamper y Pipeline Graph View."],
    ["Pipeline como código", "El pipeline se escribe en un Jenkinsfile que se versiona con el código (Jenkins, s. f.-d).",
     "El Jenkinsfile vive en la raíz del repositorio; cualquier cambio al proceso queda en el historial de Git."],
    ["Integración con Docker", "Puede usar imágenes de Docker como entorno de ejecución de una etapa o de todo el pipeline (Jenkins, s. f.-g).",
     "Las pruebas unitarias corren en un contenedor node:22-alpine y el pipeline levanta los dos contenedores del proyecto."],
    ["Disparadores", "Puede revisar el repositorio con una frecuencia definida (pollSCM) y ejecutar el pipeline si encuentra cambios (Jenkins, s. f.-e).",
     "Revisa la rama cada dos minutos."],
    ["Credenciales", "Guarda las credenciales de forma cifrada en el controlador (Jenkins, s. f.-f).",
     "La clave de PostgreSQL del pipeline no está en el repositorio."],
    ["Historial y trabajo distribuido", "Conserva el historial de cada ejecución y puede repartir el trabajo entre varias máquinas (Jenkins, s. f.-c).",
     "El historial muestra las ejecuciones en verde y en rojo. El trabajo distribuido no se usa: un solo nodo basta para este proyecto."],
], [1.4, 2.6, 2.5])

titulo("Requisitos para implementar Jenkins en el proyecto", 2)
parrafo(f"El enunciado pide documentar las características requeridas para implementar Jenkins (Politécnico Grancolombiano, s. f.). La Tabla "
        f"{sig('tab')} reúne lo que fue necesario en este proyecto, con el valor que se usó y el motivo.", keep=True)
tabla("Requisitos de la implementación de Jenkins", ["Requisito", "Qué se usó", "Motivo o fuente"], [
    ["Java", "Java 21, incluido en la imagen oficial jenkins/jenkins:2.580.2-jdk21",
     "La guía oficial de instalación pide Java 21 o superior (Jenkins, s. f.-a); el tutor mostró la instalación en Windows con JDK 21 o 25 (Reyes Corredor, 2026)."],
    ["Motor de contenedores", "Docker Desktop con contenedores Linux",
     "Es el requisito de la instalación con Docker en Windows (Jenkins, s. f.-a) y ya se usaba en la primera parte."],
    ["Puertos", "8181 del equipo para la interfaz web (8080 dentro del contenedor) y 50000 para agentes",
     "El puerto 8080 del equipo ya lo usaba otra aplicación."],
    ["Almacenamiento persistente", "Volumen jenkins-data montado en /var/jenkins_home",
     "Conserva la configuración, los trabajos y el historial aunque se recree el contenedor (Jenkins, s. f.-a)."],
    ["Acceso a Docker desde Jenkins", "Contenedor docker:dind y cliente de Docker con el plugin de Docker Compose",
     "El pipeline construye la imagen de la API y levanta los dos contenedores."],
    ["Acceso al código", "Repositorio Git del proyecto",
     "En esta entrega Jenkins lee una copia local de la rama de trabajo; un segundo trabajo descarga la rama main desde GitHub."],
    ["Plugins", "Pipeline, Git, Docker Pipeline, Credentials Binding, Timestamper y Pipeline Graph View",
     f"Ver la Tabla {sig('tab') + 1}."],
    ["Usuario administrador", "admin-ci, creado en el asistente inicial", "Su clave no se publica en el documento ni en el repositorio."],
    ["Credencial", "tareas-db-password, de tipo Secret text", "Es la clave de la base de datos que usa el pipeline."],
    ["Recursos", "Medidos en reposo, después de las ejecuciones: 1,234 GiB de memoria para el contenedor jenkins y 57,84 MiB para jenkins-docker; la imagen propia de Jenkins ocupa 1,06 GB en disco",
     "Lectura de docker stats y docker image ls en el equipo del proyecto."],
], [1.4, 2.4, 2.7])

titulo("Instalación de Jenkins con Docker", 2)
titulo("Por qué con Docker", 3)
parrafo("El material de apoyo del tutor muestra la instalación de Jenkins en Windows como un servicio, con el instalador del sistema, un JDK, la "
        "prueba del puerto 8080, la excepción del firewall y el desbloqueo inicial (Reyes Corredor, 2026). El grupo optó por la instalación con "
        "Docker que documenta el sitio oficial de Jenkins (Jenkins, s. f.-a), por tres razones: da continuidad a la primera parte, en la que todo "
        "el proyecto ya corre en contenedores; no exige instalar Java en el equipo; y deja la instalación escrita en archivos del repositorio, de "
        "modo que cualquier integrante la reproduce con un solo comando. El tutor también destacó la integración con Docker como una de las más "
        "valiosas, porque facilita el manejo de contenedores y ambientes (Reyes Corredor, 2026).")
parrafo(f"La guía oficial usa dos contenedores, como muestra la Figura {sig('fig')}. El contenedor jenkins ejecuta Jenkins. El contenedor "
        "jenkins-docker ejecuta la imagen docker:dind, que es un motor de Docker completo dentro de un contenedor; Jenkins le envía los comandos "
        "de Docker por la red jenkins, en el puerto 2376 y con certificados TLS. Los dos comparten el volumen jenkins-data, así que las rutas del "
        "espacio de trabajo existen para ambos: cuando el pipeline ejecuta docker compose, el motor encuentra los mismos archivos que descargó "
        "Jenkins, por ejemplo db/init.sql.")
figura("diagrama-jenkins.png", "Arquitectura de Jenkins con Docker en el proyecto",
       "El navegador entra a Jenkins por el puerto 8181 del equipo. Jenkins descarga el código del repositorio Git y le pide al motor docker:dind "
       "que ejecute cada etapa: un contenedor temporal de Node.js para las pruebas unitarias y, después, los dos contenedores del proyecto, "
       "tareas-app y tareas-db, con su red y su volumen. Jenkins consulta /salud a través del nombre docker, que es el alias del motor en la red "
       "jenkins. Al final de cada ejecución el pipeline borra los contenedores y el volumen, de modo que la siguiente empieza con la base de datos limpia.")
titulo("La imagen de Jenkins del proyecto", 3)
parrafo(f"El Listado {sig('lis')} es el Dockerfile de la imagen de Jenkins. Parte de la receta oficial (Jenkins, s. f.-a): la imagen "
        "jenkins/jenkins:2.580.2-jdk21, el repositorio de paquetes de Docker, el cliente docker-ce-cli y los plugins docker-workflow y "
        "json-path-api.", keep=True)
listado("Dockerfile de la imagen de Jenkins", r"jenkins\Dockerfile")
parrafo("A la receta oficial se le agregaron dos líneas. La primera instala docker-compose-plugin, porque el pipeline levanta los dos "
        "contenedores del proyecto con docker compose, igual que en la primera parte. La segunda marca como seguro cualquier repositorio para "
        "Git, porque el repositorio que Jenkins lee en esta entrega se monta desde el equipo anfitrión y pertenece a otro usuario; sin esa línea, "
        "Git se niega a leerlo.")
titulo("El script de instalación", 3)
parrafo(f"El Listado {sig('lis')} es el script que levanta todo. Sigue los pasos de la guía oficial para Windows: crea la red jenkins, inicia "
        "el contenedor docker:dind, construye la imagen y crea el contenedor de Jenkins con las variables que le indican dónde está el motor de "
        "Docker (Jenkins, s. f.-a).", keep=True)
listado("Script jenkins/levantar.ps1", r"jenkins\levantar.ps1")
parrafo("Frente a la guía oficial hay tres ajustes, todos comentados en el script: la interfaz se publica en el puerto 8181 porque el 8080 "
        "del equipo estaba ocupado; se monta en /repos, solo para lectura, una copia local del repositorio; y se activa la propiedad "
        "ALLOW_LOCAL_CHECKOUT del plugin Git para que Jenkins pueda leer esa copia, por las razones que se explican en la sección sobre la "
        "conexión con el repositorio.")
parrafo(f"La Figura {sig('fig')} muestra el resultado del script: la red, la imagen construida y los dos contenedores en ejecución.", keep=True)
figura("j01-levantar.png", "Red, imagen y contenedores de Jenkins",
       "La red jenkins es de tipo bridge. La imagen tareas-jenkins:2.580.2-1 es la que construyó el script a partir del Dockerfile. Los dos "
       "contenedores están en ejecución: jenkins publica la interfaz web en el puerto 8181 del equipo y el puerto 50000 para agentes, y "
       "jenkins-docker publica el puerto 2376 del motor de Docker.")
figura("j09-cliente-docker.png", "Jenkins se comunica con su motor de Docker",
       "Los comandos se ejecutan dentro del contenedor jenkins. El cliente de Docker de Jenkins responde con su versión y con la del motor al que "
       "está conectado, que es Linux; también están disponibles Docker Compose y Git. La variable DOCKER_HOST confirma que el motor es el "
       "contenedor docker:dind, al que Jenkins llama docker en el puerto 2376.")

titulo("Configuración inicial", 2)
parrafo("La primera vez que se abre Jenkins aparece un asistente de cuatro pasos. Se recorrió completo desde el navegador, en http://localhost:8181.")
parrafo(f"**Paso 1, desbloqueo.** Jenkins genera una clave de administrador de un solo uso y pide escribirla para continuar (Figura {sig('fig')}). "
        "La clave está en el archivo /var/jenkins_home/secrets/initialAdminPassword del contenedor y también aparece en su registro (Jenkins, "
        "s. f.-a); se leyó con docker exec. Así se garantiza que solo quien tiene acceso al servidor puede terminar la instalación.", keep=True)
figura("j02-desbloqueo.png", "Pantalla de desbloqueo de Jenkins",
       "Jenkins indica la ruta del archivo con la clave inicial y pide pegarla en el campo Administrator password. Mientras no se desbloquee, "
       "nadie puede usar la instalación.")
parrafo(f"**Paso 2, plugins.** Jenkins ofrece instalar los plugins sugeridos o elegirlos uno por uno (Figura {sig('fig')}). Se eligió la opción "
        "sugerida, porque incluye los que el proyecto necesita: Pipeline, Git, Credentials Binding y Timestamper, entre otros. El plugin de Docker "
        "Pipeline ya venía en la imagen propia.", keep=True)
figura("j03-personalizar.png", "Elección de los plugins iniciales",
       "La opción Install suggested plugins instala el conjunto que la comunidad recomienda para la mayoría de los casos. La otra opción, "
       "Select plugins to install, permite elegirlos; los plugins también se pueden agregar después desde Manage Jenkins.")
figura("j04-instalando-plugins.png", "Instalación de los plugins sugeridos",
       "Cada recuadro es un plugin; los verdes ya se instalaron y los demás están en cola. A la derecha, Jenkins lista lo que va instalando, "
       "incluidas las dependencias, marcadas con dos asteriscos. Según el registro del contenedor, la instalación terminó con 43 plugins en 52 segundos.")
parrafo(f"**Paso 3, primer administrador.** Se creó el usuario admin-ci (Figura {sig('fig')}), con el que se hizo todo lo demás. La clave se "
        "guardó en un archivo local fuera del repositorio.", keep=True)
figura("j05-primer-administrador.png", "Creación del primer usuario administrador",
       "El formulario pide usuario, clave, nombre y correo. Los campos de clave aparecen enmascarados. El correo es ficticio, porque esta "
       "instalación no envía notificaciones.")
parrafo(f"**Paso 4, dirección y arranque.** Se confirmó http://localhost:8181/ como dirección de la instancia y Jenkins quedó listo. La Figura "
        f"{sig('fig')} muestra el tablero inicial, todavía sin trabajos.", keep=True)
figura("j08-tablero-inicial.png", "Tablero inicial de Jenkins",
       "El tablero ofrece crear el primer trabajo (Create a job o New Item) y configurar agentes o nubes para repartir el trabajo. A la izquierda "
       "están la cola de ejecuciones y el estado de los ejecutores: dos, ambos libres. Abajo a la derecha se lee la versión instalada, 2.580.2.")

titulo("Plugins utilizados", 2)
parrafo(f"Después del asistente, Jenkins quedó con 95 plugins activos, contando dependencias, según su API de administración. La Tabla "
        f"{sig('tab')} lista los que usa el pipeline del proyecto y para qué.", keep=True)
tabla("Plugins que usa el pipeline", ["Plugin", "Versión", "Para qué se usa"], [
    ["Pipeline", "608.v67378e9d3db_1", "Permite definir el proceso en un Jenkinsfile con sintaxis declarativa (Jenkins, s. f.-d)."],
    ["Git", "5.11.0", "Descarga el código del repositorio y revisa si hay commits nuevos (Jenkins, s. f.-b)."],
    ["Docker Pipeline", "653.v2f2c08eff0ec", "Habilita agent { docker } para correr una etapa dentro de un contenedor (Jenkins, s. f.-g)."],
    ["Credentials Binding", "728.v902a_273b_8947", "Entrega la clave de la base de datos al pipeline con withCredentials, sin escribirla en el código."],
    ["Timestamper", "1.30", "Agrega la hora a cada línea de la consola, lo que ayuda a ver cuánto tarda cada paso."],
    ["Pipeline Graph View", "1053.v9df19e6ceed2", "Dibuja las etapas de cada ejecución y su estado."],
], [1.6, 1.6, 3.3])

titulo("Credencial de la base de datos", 2)
parrafo(f"El pipeline levanta una base de datos PostgreSQL y necesita su clave. Para no escribirla en el repositorio, se guardó en Jenkins como una "
        "credencial de tipo Secret text, con alcance global e identificador tareas-db-password. Jenkins guarda las credenciales cifradas en el "
        f"controlador (Jenkins, s. f.-f). La Figura {sig('fig')} muestra la credencial registrada: se ven su identificador y su descripción, nunca su valor.")
figura("j10-credencial.png", "Credencial tareas-db-password en Jenkins",
       "La credencial aparece en el almacén System con dominio Global. Durante cada ejecución, el pipeline la recibe en una variable con "
       "withCredentials, escribe el archivo .env que necesita Docker Compose y lo borra al terminar.")

titulo("El pipeline del proyecto: archivo Jenkinsfile", 2)
parrafo("El pipeline se escribió con la sintaxis declarativa, pensada para que el código del pipeline sea fácil de escribir y de leer (Jenkins, "
        "s. f.-d). En ella, un bloque stage agrupa un conjunto de tareas con un propósito, como construir o probar, y cada step es una tarea "
        f"puntual, como ejecutar un comando con sh (Jenkins, s. f.-d). El Listado {sig('lis')} contiene el archivo completo.")
listado("Archivo Jenkinsfile del proyecto", "Jenkinsfile")
parrafo("Las secciones del archivo cumplen funciones distintas (Jenkins, s. f.-e). **agent any** indica que el pipeline puede correr en "
        "cualquier ejecutor disponible. **options** fija reglas de la ejecución: la hora en cada línea de la consola, que no corran dos "
        "ejecuciones a la vez, que se conserven las últimas treinta y un tiempo límite de quince minutos, después del cual Jenkins aborta la "
        "ejecución. **triggers** con pollSCM('H/2 * * * *') le pide a Jenkins revisar el repositorio cada dos minutos y ejecutar el pipeline si "
        "encuentra commits nuevos; la letra H reparte las revisiones en el tiempo según el nombre del trabajo, para que no coincidan todas en el "
        "mismo minuto. **environment** define el nombre del proyecto de Docker Compose y el puerto de la API. **post** define lo que se hace al "
        "final: always se ejecuta siempre, sin importar el resultado, mientras que success y failure solo cuando la ejecución termina bien o mal.")
parrafo(f"La Tabla {sig('tab')} explica cada etapa y qué comprueba.", keep=True)
tabla("Etapas del pipeline", ["Etapa", "Qué hace", "Qué comprueba"], [
    ["Descargar código", "Descarga la rama configurada en el trabajo e imprime el commit que se va a probar.", "Que Jenkins accede al repositorio y sabe qué versión prueba."],
    ["Instalar dependencias y pruebas unitarias", "Dentro de un contenedor node:22-alpine ejecuta npm ci y npm test, con un tiempo límite de tres minutos.",
     "La validación de entradas y los códigos HTTP de la API, sin base de datos."],
    ["Construir la imagen de la API", "Ejecuta docker compose build app con el Dockerfile de la primera parte.", "Que la imagen se construye desde el repositorio."],
    ["Levantar los dos contenedores", "Escribe el .env con la credencial y ejecuta docker compose up -d --wait, que espera a que los servicios estén listos.",
     "Que los dos contenedores arrancan y la base de datos queda saludable."],
    ["Verificar /salud", "Consulta http://docker:8095/salud con curl; si la respuesta no es exitosa, la etapa falla.", "Que la API responde y llega a la base de datos."],
    ["Pruebas de integración", "Ejecuta npm test dentro del contenedor de la API, donde existe DB_HOST.", "El ciclo completo de crear, leer, actualizar y borrar contra la base real."],
    ["Acciones finales (post)", "Ejecuta docker compose down -v y borra el .env, pase lo que pase.", "Que no quedan contenedores, volúmenes ni claves de una ejecución a otra."],
], [1.6, 2.7, 2.2])
parrafo("Node.js no está instalado en Jenkins. La segunda etapa usa agent { docker { image 'node:22-alpine' } }, que corre esa etapa dentro de un "
        "contenedor de esa imagen; la opción reuseNode true mantiene el mismo espacio de trabajo de las demás etapas, así que el contenedor ve el "
        "código ya descargado (Jenkins, s. f.-g). Esa es la primera forma en que Jenkins se integra con Docker en este proyecto; la segunda es que el "
        "propio pipeline construye y levanta los contenedores de la aplicación.")

titulo("Creación del trabajo y conexión con el repositorio", 2)
parrafo("En Jenkins, un trabajo de tipo Pipeline es el que ejecuta un Jenkinsfile. Se creó desde la interfaz con estos pasos: New Item, nombre "
        "tareas-equipo, tipo Pipeline y OK; luego, en la sección Pipeline de la configuración, Definition en Pipeline script from SCM, SCM en Git, "
        "la dirección del repositorio, la rama que se debe construir y la ruta del Jenkinsfile, que se dejó en su valor por defecto, Jenkinsfile. "
        f"Las Figuras {sig('fig')} y {sig('fig') + 1} muestran esa configuración.", keep=True)
figura("j12-configuracion-scm.png", "Configuración del trabajo: el pipeline se lee del repositorio",
       "Con Pipeline script from SCM, Jenkins no guarda el pipeline: lo lee del Jenkinsfile del repositorio en cada ejecución. El SCM es Git y la "
       "dirección es file:///repos/poli-ic-proyecto1.git, la copia local del repositorio montada en el contenedor. No se necesitan credenciales "
       "para leerla.")
figura("j12b-ruta-jenkinsfile.png", "Rama que construye el trabajo",
       "En Branch Specifier se indicó */poli-trabajos/ic-proyecto2-jenkins, la rama donde se desarrolló esta entrega. Jenkins solo construye "
       "esa rama y solo vigila sus cambios.")
parrafo("La dirección del repositorio merece una explicación. La rama de esta entrega, que contiene el Jenkinsfile, todavía no estaba publicada "
        "en GitHub cuando se tomaron las capturas. Por eso Jenkins lee una copia local de esa rama, montada solo para lectura. Por seguridad, el "
        "plugin Git de las versiones actuales no permite leer repositorios locales, según la alerta SECURITY-2478; la documentación del plugin "
        "indica que el administrador puede permitirlo, si acepta el riesgo, con la propiedad hudson.plugins.git.GitSCM.ALLOW_LOCAL_CHECKOUT=true "
        "al arrancar Jenkins (Jenkins, s. f.-b). Se aceptó porque la instalación es local, la copia es de solo lectura y la medida es temporal: al "
        "publicar la rama, basta con cambiar la dirección por https://github.com/Owito/poli-ic-proyecto1.git, sin tocar nada más. La sección "
        "sobre la integración con GitHub muestra que esa conexión ya funciona.")

titulo("Ejecuciones del pipeline", 2)
parrafo(f"La Tabla {sig('tab')} resume las ejecuciones del trabajo tareas-equipo. La primera se lanzó a mano con la API de Jenkins, lo que "
        "equivale al botón Build Now; las demás las lanzó Jenkins solo, al encontrar commits nuevos en la rama. Los datos salen de la consola y "
        "de la API de cada ejecución.", keep=True)
tabla("Ejecuciones del trabajo tareas-equipo", ["N.º", "Cómo se inició", "Commit probado", "Resultado", "Duración", "Qué mostró"], [
    ["1", "A mano", "50c76c4, el Jenkinsfile", "SUCCESS", "1 min 5 s", "El pipeline completo funciona; la primera vez descarga la imagen de Node.js."],
    ["2", "Cambio en el repositorio", "335a28e, error introducido a propósito", "ABORTED", "7 min 6 s", "La prueba falló, pero la ejecución no terminaba; se abortó a mano."],
    ["3", "Cambio en el repositorio", "34592c9, pruebas y tiempo límite corregidos", "FAILURE", "7,8 s", "El error sigue y el pipeline se detiene en las pruebas unitarias."],
    ["4", "Cambio en el repositorio", "8e708a8, revierte el error", "SUCCESS", "19 s", "Todo en verde otra vez."],
], [0.4, 1.0, 1.5, 0.9, 0.8, 1.9])
parrafo("Los identificadores de commit de esta sección son los de la rama publicada en GitHub. Antes de publicarla se ajustaron los mensajes de sus commits, sin cambiar el código, y por eso cambiaron sus identificadores. Las consolas y las capturas de Jenkins conservan los que tenían en la copia local del repositorio: 50c76c4 aparece allí como 35489a5, 335a28e como d06d969, 34592c9 como 25bcac6 y 8e708a8 como 53d6fe3.")
titulo("Una ejecución completa en verde", 3)
parrafo(f"La Figura {sig('fig')} muestra las etapas de la ejecución 4 y la Figura {sig('fig') + 1} las líneas clave de su consola.", keep=True)
figura("j15-etapas-verde.png", "Etapas de la ejecución 4 en verde",
       "Arriba se lee que la ejecución la inició un cambio en el repositorio (Started by an SCM change) y que tardó 19 segundos. El grafo "
       "muestra las seis etapas del Jenkinsfile, más la descarga inicial del Jenkinsfile (Checkout SCM) y las acciones finales, todas en verde. "
       "A la izquierda se ve cuánto tardó cada etapa; por ejemplo, levantar los dos contenedores tomó 7 segundos.")
figura("j21-consola-verde.png", "Consola de la ejecución 4",
       "Se ven el commit probado, la versión de Node.js del contenedor de pruebas, las cinco pruebas unitarias (cuatro aprobadas y la de "
       "integración omitida, porque en ese contenedor no hay base de datos), la imagen construida, los dos contenedores en ejecución con "
       "tareas-db saludable, la respuesta de /salud con la versión de PostgreSQL leída de la base y las pruebas de integración, que esta vez sí "
       "corren y pasan. La última línea es Finished: SUCCESS.")
titulo("Ejecución automática ante un cambio", 3)
parrafo("Para comprobar el disparador, se hizo un commit en la rama y se subió a la copia del repositorio que lee Jenkins (Figura "
        f"{sig('fig')}). Sin que nadie entrara a Jenkins, en la siguiente revisión del repositorio apareció la ejecución 2, iniciada por el cambio.", keep=True)
figura("j13-espejo.png", "Envío de un commit a la rama que vigila Jenkins",
       "git push envía el commit 335a28e a la copia local del repositorio; git ls-remote confirma que la rama apunta a ese commit. Desde ese "
       "momento, la próxima revisión de pollSCM encuentra un commit que no se ha construido y lanza el pipeline.")
titulo("El pipeline detiene la integración cuando una prueba falla", 3)
parrafo("El commit 335a28e se hizo a propósito para comprobar que Jenkins detecta errores: cambia la API para que una tarea sin título reciba un "
        "título por defecto, lo que rompe la regla de negocio de que el título es obligatorio. La prueba unitaria POST /tareas sin titulo espera "
        "la respuesta HTTP 400 y, con el cambio, recibió 500.")
parrafo("En la ejecución 2, Jenkins detectó el fallo, pero la ejecución no terminó: la prueba fallida dejaba abierto el servidor de prueba, porque "
        "la instrucción que lo cerraba estaba después de la aserción y nunca se ejecutaba, y Node.js se quedaba esperando. Tras siete minutos se "
        "abortó a mano. El hallazgo llevó a dos correcciones en el commit 34592c9: las pruebas cierran el servidor con t.after(), que se ejecuta "
        "aunque la prueba falle, y el Jenkinsfile tiene ahora un tiempo límite general de quince minutos y uno de tres minutos para las pruebas "
        f"unitarias. Con eso, la ejecución 3 terminó en rojo en 7,8 segundos (Figuras {sig('fig')} y {sig('fig') + 1}).")
figura("j17-etapas-rojo.png", "Etapas de la ejecución 3: falla en las pruebas unitarias",
       "La etapa de dependencias y pruebas unitarias está en rojo. Las cuatro siguientes aparecen saltadas: Jenkins no construyó la imagen ni "
       "levantó los contenedores, porque no tiene sentido desplegar código que no pasa sus pruebas. Las acciones finales sí se ejecutaron. A la "
       "derecha, el resumen de npm test: cinco pruebas, tres aprobadas, una fallida y una omitida, y el código de salida 1 que marca la falla.")
figura("j22-consola-rojo.png", "Consola de la ejecución 3",
       "La prueba 1 reporta not ok, con el valor esperado, 400, y el obtenido, 500. Después, cada etapa siguiente aparece como skipped due to "
       "earlier failure(s), se imprime el mensaje de falla definido en post y la ejecución termina con Finished: FAILURE.")
parrafo(f"Finalmente se revirtió el commit con el error (commit 8e708a8) y la ejecución 4 volvió a quedar en verde. La Figura {sig('fig')} "
        "muestra las cuatro ejecuciones, etapa por etapa.", keep=True)
figura("j20-etapas-historial.png", "Historial de ejecuciones por etapa",
       "Cada fila es una ejecución y cada círculo una etapa: el primero es la descarga del Jenkinsfile, siguen las seis etapas en el orden del "
       "Jenkinsfile y el último son las acciones finales. La 1 y la 4 están completas en verde. La 3 falla en el tercer círculo, la etapa de "
       "pruebas unitarias, y salta las siguientes. La 2 quedó abortada en esa misma etapa. Es la clase de historial que "
       "el tutor mostró en el encuentro como una de las ventajas de Jenkins (Reyes Corredor, 2026).")

titulo("Integración con GitHub", 2)
parrafo("Para comprobar que Jenkins se conecta con el repositorio público de GitHub, se creó un segundo trabajo, tareas-equipo-github. Como la "
        "rama main de GitHub todavía no tiene Jenkinsfile, su pipeline se escribió en el propio trabajo, con la opción Pipeline script; el mismo "
        f"script está versionado en el archivo jenkins/tareas-equipo-github.groovy. La Figura {sig('fig')} muestra la creación del trabajo.")
figura("j11-nuevo-trabajo.png", "Creación de un trabajo de tipo Pipeline",
       "En New Item se escribe el nombre del trabajo y se elige el tipo. Pipeline es el tipo que ejecuta un pipeline como código; los otros, como "
       "Freestyle project o Multibranch Pipeline, sirven para casos distintos. El trabajo tareas-equipo se creó de la misma forma.")
figura("j16-script-github.png", "Script del trabajo tareas-equipo-github",
       "Con Definition en Pipeline script, el pipeline se escribe en el editor del trabajo. La etapa Descargar desde GitHub usa el paso git con la "
       "dirección https://github.com/Owito/poli-ic-proyecto1.git y la rama main. Como el repositorio es público, no hace falta credencial.")
figura("j18-github-ejecucion.png", "Ejecución del trabajo tareas-equipo-github",
       "Las dos etapas, Descargar desde GitHub y Pruebas unitarias, terminaron en verde. A la derecha se ve el resumen de npm test con cuatro "
       "pruebas aprobadas, ninguna fallida y una omitida.")
figura("j23-consola-github.png", "Consola del trabajo tareas-equipo-github",
       "Jenkins clona el repositorio desde GitHub y prueba el commit 190bf46, que es el último de la rama main. Es la evidencia de que la conexión "
       "con GitHub funciona: cuando la rama de esta entrega se publique, el trabajo tareas-equipo solo necesita cambiar la dirección del repositorio.")

titulo("Dificultades y cómo se resolvieron", 2)
parrafo("**El puerto 8080 estaba ocupado.** Es el puerto por defecto de Jenkins, pero en el equipo del proyecto lo usaba otra aplicación. Se "
        "publicó la interfaz en el 8181, que es un parámetro del script de instalación.")
parrafo("**La rama no estaba en GitHub.** Jenkins necesitaba leer el Jenkinsfile antes de publicar la rama. Se resolvió con la copia local de "
        "solo lectura y la propiedad del plugin Git que lo permite, como medida temporal, y se comprobó la conexión con GitHub con un segundo trabajo.")
parrafo("**Una prueba fallida dejaba la ejecución esperando.** Fue la dificultad más útil, porque la descubrió el propio Jenkins: sin un tiempo "
        "límite, una prueba que no termina bloquea la integración indefinidamente. Se corrigieron las pruebas para que siempre cierren el servidor "
        "y se agregaron tiempos límite al pipeline.")

titulo("Cómo reproducir la segunda parte", 2)
parrafo("Con Docker Desktop en ejecución, se crea una copia local del repositorio con git clone --bare y se ejecuta powershell -File "
        "jenkins\\levantar.ps1 -Espejo con la ruta de esa copia. Después se abre http://localhost:8181, se completa el asistente con la clave "
        "inicial, se crea la credencial tareas-db-password y el trabajo tareas-equipo con la configuración de las Figuras "
        f"{F('j12-configuracion-scm')} y {F('j12b-ruta-jenkinsfile')}. La primera ejecución se lanza con Build Now; desde ahí, cada commit nuevo "
        "en la rama lanza el pipeline solo.")

# ================================================================ TERCERA PARTE (ENTREGA 3)
# Letras de las referencias sin fecha de cada autor, en el orden alfabético de sus títulos (APA 7). Se usan en las citas.
TR = {"faq": "a", "facturacion": "b", "node": "c", "validacion": "d", "etapas": "e", "precios": "f", "ciclo": "g",
      "inicio": "h", "noble": "i", "docker": "j"}
CB = {"free": "a", "aviso": "b", "faq": "c", "unify": "d", "gha": "e", "resumen": "f", "precio": "g", "jet": "h",
      "services": "i", "steps": "j", "flujos": "k"}
GH = {"facturacion": "a", "expresiones": "b", "pack": "c", "sintaxis": "f"}


def tr(k):
    return f"Travis CI, s. f.-{TR[k]}"


def cb(k):
    return f"CloudBees, s. f.-{CB[k]}"


def gh(k):
    return f"GitHub, s. f.-{GH[k]}"


def historial():
    """Historial real del repositorio: una fila por commit, del primero al más reciente."""
    salida = subprocess.run(["git", "-C", str(RAIZ), "log", "--reverse", "--date=format:%Y-%m-%d", "--format=%h|%ad|%an|%s"],
                            capture_output=True, text=True, encoding="utf-8", check=True).stdout
    filas = []
    for i, linea_git in enumerate(salida.strip().splitlines(), 1):
        corto, fecha, autor, asunto = linea_git.split("|", 3)
        tipo, _, descripcion = asunto.partition(": ")
        entrega = "1" if fecha < "2026-11-10" else ("2" if fecha < "2026-11-19" else "3")
        filas.append([str(i), corto, fecha, autor, tipo, descripcion, entrega])
    return filas


HISTORIAL = historial()

salto()
titulo("Tercera parte: integración con Travis CI y Codeship, historial y solución de problemas", 1)
parrafo("El enunciado de la Entrega 3 pide que, en este punto, la plataforma de software esté totalmente integrada con contenedores, "
        "Jenkins, Travis CI y Codeship, y que en un documento se consoliden el historial de cambios y las sugerencias para la solución "
        "de problemas, así como las responsabilidades y opiniones (Politécnico Grancolombiano, s. f.). Esta parte agrega al proyecto dos "
        "integraciones continuas en la nube que ejecutan las mismas etapas que el Jenkinsfile de la segunda parte, explica qué ocurrió con "
        "Codeship y cómo se reemplazó, y reúne el historial, los problemas encontrados en las tres entregas con su solución, y las "
        "responsabilidades y opiniones del grupo.")
parrafo("Para elegir las herramientas se tuvo en cuenta una indicación que el tutor repitió en varios encuentros: si una herramienta "
        "obliga a pagar una licencia, el grupo debe buscar otra que cumpla la misma función y que sea libre, porque “no se trata de que "
        "paguen la herramienta” (Reyes Corredor, 2026).")

titulo("Estado de Travis CI y Codeship en noviembre de 2026", 2)
parrafo(f"Los servicios de integración continua en la nube cambian sus planes con el tiempo, así que antes de escribir cualquier "
        f"configuración se consultó, el 19 de noviembre de 2026, la documentación oficial de cada servicio. La Tabla {sig('tab')} resume "
        f"lo encontrado y las Figuras {sig('fig')}, {sig('fig') + 1} y {sig('fig') + 2} muestran las fuentes.")
tabla("Estado de los servicios verificado el 19 de noviembre de 2026", ["Servicio", "Qué se encontró", "Consecuencia para el proyecto"], [
    ["Travis CI",
     f"Sigue en operación. Los planes publicados son de pago; el más económico, Usage Based, cuesta 15 dólares al mes por 35.000 créditos de Linux, o 13,75 dólares al mes con pago anual ({tr('precios')}). "
     f"La documentación de facturación indica que la prueba gratuita no está disponible por ahora y que un minuto de Linux cuesta 10 "
     f"créditos ({tr('facturacion')}). Ofrece créditos para código abierto, solo para repositorios públicos, que se solicitan al soporte "
     f"({tr('faq')}), y el GitHub Student Developer Pack incluye compilaciones privadas gratuitas de Travis CI mientras se es estudiante ({gh('pack')}).",
     "Se puede integrar. El archivo .travis.yml queda listo y validado; para ejecutarlo en la nube, la cuenta de Travis CI necesita créditos por una de esas vías."],
    ["Codeship",
     f"CloudBees dejó de dar soporte a CodeShip en enero de 2026 y recomienda migrar a CloudBees Unify ({cb('aviso')}); el retiro definitivo "
     f"estaba anunciado para el 31 de enero de 2026 ({cb('resumen')}). Antes, CloudBees había anunciado el retiro del plan gratuito para el 1 de marzo de 2023 ({cb('free')}). "
     "Su aplicación, app.codeship.com, redirige al aviso de fin de vida.",
     "No se puede integrar, porque el servicio ya no existe. Se reemplaza por una herramienta equivalente."],
    ["CloudBees Unify",
     f"Es el producto al que CloudBees recomienda migrar. Su precio se obtiene con una cotización personalizada ({cb('precio')}) y su "
     f"página ofrece una prueba gratuita ({cb('unify')}).",
     "Se evaluó como reemplazo de Codeship."],
    ["GitHub Actions",
     f"Su uso es gratuito con los ejecutores estándar de GitHub en repositorios públicos ({gh('facturacion')}).",
     "Se evaluó como reemplazo de Codeship."],
], [1.1, 3.4, 2.0])
figura("t01-travis-prueba-gratuita.png", "Aviso de la documentación de Travis CI sobre la prueba gratuita",
       "La sección Free Trial Plan de la documentación de facturación de Travis CI empieza con el aviso Free Trial Plans are currently "
       "unavailable, es decir, que por ahora no se ofrecen planes de prueba gratuita. El texto siguiente describe cómo era esa prueba: "
       "10.000 créditos, equivalentes a 1.000 minutos de Linux, válidos por 14 días, y una validación con tarjeta de crédito mediante un "
       "cobro de autorización de un dólar que luego se devuelve. Por eso la activación de Travis CI queda como una decisión del grupo.")
figura("t02-codeship-fin-de-vida.png", "Aviso de fin de vida de CloudBees CodeShip",
       "La dirección de la documentación de CodeShip ya no muestra la documentación del producto: redirige a esta página, en la que "
       "CloudBees anuncia que desde enero de 2026 no da soporte a CloudBees CodeShip y recomienda migrar a CloudBees Unify.", ancho=4.6)
figura("t03-codeship-cierre.png", "Comprobación del cierre de Codeship desde la terminal",
       "Después de la fecha y la hora de la consulta, el primer curl pide la aplicación de Codeship, app.codeship.com, donde se configuraban "
       "los proyectos, y el servidor responde con una redirección permanente (301) hacia el aviso de fin de vida. El segundo hace lo mismo con "
       "la documentación y el tercero confirma el texto del aviso. El último pide el instalador de Jet, la herramienta de línea de comandos con la que CodeShip Pro permitía "
       f"ejecutar su configuración en el equipo local ({cb('jet')}), y el servidor responde 403 Forbidden: tampoco hay una forma local de "
       "ejecutar una configuración de Codeship.")

titulo("Arquitectura de la integración continua del proyecto", 2)
parrafo(f"Con esos hallazgos, el proyecto queda integrado con tres herramientas que leen el mismo repositorio: Jenkins, de la segunda "
        f"parte; Travis CI; y GitHub Actions, en lugar de Codeship. La Figura {sig('fig')} resume la arquitectura y la Tabla "
        f"{sig('tab')} muestra que las tres ejecutan las mismas etapas.", keep=True)
figura("diagrama-tres-ci.png", "Arquitectura de la integración continua en la Entrega 3",
       "Cada integrante sube sus cambios con git push al repositorio de GitHub, donde están los tres archivos de configuración: Jenkinsfile, "
       ".travis.yml y .github/workflows/integracion-continua.yml. Jenkins revisa la rama cada dos minutos desde el equipo del proyecto; Travis CI "
       "y GitHub Actions se activan con cada push y ejecutan el trabajo en máquinas de la nube. Abajo están las seis etapas comunes. Codeship "
       "aparece como retirado y reemplazado por GitHub Actions.")
tabla("Etapas equivalentes en las tres herramientas", ["Etapa", "Jenkins (Jenkinsfile)", "Travis CI (.travis.yml)", "GitHub Actions (integracion-continua.yml)"], [
    ["Descargar el código", "checkout scm", "Travis CI clona el repositorio en la máquina virtual antes de empezar", "actions/checkout@v7"],
    ["Pruebas unitarias", "Contenedor node:22-alpine, npm ci y npm test, con límite de 3 minutos",
     "Etapa pruebas unitarias: Node.js 22 de la máquina virtual, npm ci y timeout 180 npm test",
     "Job pruebas-unitarias: contenedor node:22-alpine, npm ci y npm test, con timeout-minutes: 3"],
    ["Construir la imagen", "docker compose build app", "docker compose build app, en before_script", "Paso docker compose build app"],
    ["Levantar los dos contenedores", ".env con la credencial de Jenkins y docker compose up -d --wait",
     ".env con la variable CLAVE_DB o una clave al azar, y docker compose up -d --wait", ".env con una clave al azar y docker compose up -d --wait"],
    ["Verificar /salud", "curl a http://docker:8095/salud", "curl a http://localhost:8095/salud, con reintentos", "Igual que en Travis CI"],
    ["Pruebas de integración", "docker compose exec -T app npm test", "Igual que en Jenkins", "Igual que en Jenkins"],
    ["Limpieza", "Sección post, always: docker compose down -v y borrar el .env", "after_script, que corre siempre", "Último paso, con if: always()"],
    ["Si fallan las unitarias", "Las etapas siguientes se saltan", "Travis CI cancela la etapa contenedores", "El job contenedores se salta (needs)"],
], [1.2, 1.7, 1.8, 1.8])

titulo("Travis CI", 2)
titulo("Qué es Travis CI", 3)
parrafo(f"Travis CI es un servicio de integración continua que se usa para probar y construir proyectos alojados en GitHub, Assembla, "
        f"Bitbucket o GitLab ({tr('inicio')}). Para cada job crea una máquina virtual con el ambiente de compilación, clona el repositorio en "
        f"ella y ejecuta las fases definidas en el archivo .travis.yml ({tr('ciclo')}). A diferencia de Jenkins, no hay que instalar ni mantener "
        "un servidor: la configuración vive en el repositorio y la ejecución ocurre en la infraestructura de Travis CI. El tutor la nombró, "
        "junto con GitHub, Docker y Jenkins, entre las herramientas que el proyecto debía permitir implementar (Reyes Corredor, 2026).")
titulo("Cómo se conecta el repositorio de GitHub con Travis CI", 3)
parrafo(f"La guía de inicio de Travis CI describe la conexión en seis pasos, que la Tabla {sig('tab')} adapta al proyecto ({tr('inicio')}).", keep=True)
TABLA_TRAVIS = sig("tab")
tabla("Pasos para conectar el repositorio con Travis CI", ["Paso", "Acción", "Detalle"], [
    ["1", "Entrar a travis-ci.com e iniciar sesión con la cuenta de GitHub", "El registro se hace con la cuenta del servicio donde está el repositorio, en este caso la propietaria de Owito/poli-ic-proyecto1."],
    ["2", "Aceptar la autorización de Travis CI", "GitHub muestra la solicitud de autorización de Travis CI y el propietario la acepta."],
    ["3", "Elegir un plan",
     f"Por ahora la prueba gratuita no está disponible ({tr('facturacion')}). Las alternativas sin pago son pedir al soporte créditos para "
     f"código abierto, indicando el nombre de la cuenta y los créditos que se necesitan ({tr('faq')}), o el plan del GitHub Student Developer "
     f"Pack, que según la guía puede no asignarse solo y requerir una solicitud al soporte ({tr('inicio')})."],
    ["4", "Registrar los datos de facturación", f"La guía indica que, para iniciar el plan elegido, hay que registrar los datos de facturación y una tarjeta de crédito válida ({tr('inicio')})."],
    ["5", "Settings y Manage Repositories", "Desde el perfil, en Settings, el botón Manage Repositories lleva a GitHub para elegir los repositorios que usará Travis CI; ahí se selecciona Owito/poli-ic-proyecto1."],
    ["6", "Subir un commit", f"Travis CI solo construye los commits que se suben después de agregar el archivo .travis.yml ({tr('inicio')}); el resultado se consulta en la página del repositorio en Travis CI."],
], [0.5, 2.0, 4.0])
parrafo("Estos pasos todavía no se ejecutaron: crear la cuenta y elegir el plan son decisiones del grupo, porque implican aceptar las "
        "condiciones de un tercero y, en un plan de pago, un costo. Mientras tanto, el archivo se validó y se ejecutó en el equipo del "
        "proyecto, como se explica más adelante.")
titulo("El archivo .travis.yml explicado", 3)
parrafo(f"El Listado {sig('lis')} contiene el archivo completo, guardado en la raíz del repositorio.", keep=True)
listado("Archivo .travis.yml del proyecto", ".travis.yml")
parrafo(f"**version: ~> 1.0** activa la validación de la configuración, que revisa el archivo contra una especificación formal en JSON "
        f"Schema y muestra sus mensajes en la pestaña View Config de cada ejecución ({tr('validacion')}). **os: linux** y **dist: noble** piden "
        f"una máquina con Ubuntu 24.04, que según la documentación trae Docker 28.5.1, docker-compose 2.40.2 y Node.js 22.14.0 ({tr('noble')}). "
        f"**language: node_js** con **node_js: \"22\"** usa la última versión 22.x de Node.js ({tr('node')}), la misma versión mayor de la imagen "
        "node:22-alpine con la que se construye la API.")
parrafo(f"**jobs.include** define dos jobs, cada uno en su propia etapa (stage). Travis CI ejecuta las etapas una después de la otra y solo "
        f"pasa a la siguiente si todos los jobs de la anterior terminaron bien; si uno falla, cancela los jobs de las etapas siguientes y la "
        f"ejecución queda fallida ({tr('etapas')}). Así se reproduce la regla del Jenkinsfile: si las pruebas unitarias fallan, no se construye "
        "la imagen ni se levantan los contenedores.")
parrafo(f"Cada job pasa por fases con reglas distintas ({tr('ciclo')}). Si falla un comando de install o de before_script, el job queda en "
        "estado errored y se detiene. Si falla un comando de script, los demás comandos de script se ejecutan igual y el job queda failed. "
        "after_script se ejecuta siempre y no cambia el resultado. Por eso la construcción de la imagen y el arranque de los contenedores "
        "están en before_script, ya que sin ellos no tiene sentido probar nada; la consulta a /salud y las pruebas de integración están en "
        "script, y si /salud fallara, las pruebas de integración se ejecutarían igual y también fallarían, lo que da más información sin "
        "cambiar el resultado; y la limpieza está en after_script.")
parrafo(f"En la segunda etapa, **services: docker** activa Docker en la máquina virtual ({tr('docker')}) e **install: skip** omite la "
        f"instalación de dependencias en la máquina ({tr('ciclo')}), porque las pruebas de integración corren dentro del contenedor de la API. "
        f"La línea que empieza con docker compose version cubre una diferencia de nombres: la documentación de Travis CI habla del programa "
        f"docker-compose ({tr('docker')}), mientras el proyecto usa el subcomando docker compose; si el subcomando no existe, se enlaza el "
        "programa como complemento de Docker.")
parrafo("Tres decisiones vienen de lo aprendido con Jenkins. La clave de la base de datos no está en el archivo: se toma de la variable "
        "CLAVE_DB si el grupo la define en la configuración del repositorio en Travis CI y, si no, se genera al azar con openssl en cada "
        "ejecución, porque la base solo existe dentro de la máquina virtual y se borra al final. Las pruebas unitarias tienen un tiempo "
        "límite de tres minutos con el comando timeout, la lección de la ejecución 2 de Jenkins. Y la consulta a /salud reintenta hasta "
        "diez veces, cada dos segundos, porque en docker-compose.yml solo la base de datos tiene comprobación de salud y la API puede tardar "
        "unos instantes en aceptar conexiones después de arrancar.")
titulo("Ejecución en Travis CI", 3)
parrafo("[[COMPLETAR: captura del build en Travis una vez Carlos conecte la cuenta (las dos etapas en verde y la consola de la etapa "
        "contenedores con la respuesta de /salud y las pruebas de integración)]]")

titulo("Codeship y su reemplazo", 2)
titulo("Qué era Codeship", 3)
parrafo(f"CodeShip era el servicio de integración continua en la nube de CloudBees y tenía dos modalidades. CodeShip Basic ejecutaba los "
        f"comandos de cada proyecto, configurados en la propia aplicación, dentro de un contenedor con Ubuntu 20.04 ({cb('free')}). CodeShip Pro "
        f"funcionaba con Docker y se configuraba con dos archivos del repositorio: codeship-services.yml, donde cada servicio necesario para la "
        f"ejecución corre en un contenedor de Docker y cuya sintaxis es muy parecida, y en buena parte compatible, con la de docker-compose.yml "
        f"({cb('services')}), y codeship-steps.yml, con la lista de pasos que se ejecutan en orden, cada uno con el servicio en el que corre y su "
        f"comando ({cb('steps')}). Jet, su herramienta de línea de comandos, permitía ejecutar esa misma configuración en el equipo local ({cb('free')}).")
titulo("Por qué no se puede integrar", 3)
parrafo(f"Codeship ya no existe como servicio. CloudBees anunció que lo descontinuaría el 31 de enero de 2026 y que después de esa fecha "
        f"quedaría inaccesible ({cb('resumen')}), y hoy su documentación y su aplicación redirigen al aviso de fin de vida (Figuras "
        f"{F('t02-codeship-fin-de-vida')} y {F('t03-codeship-cierre')}). Además, CloudBees había anunciado el retiro del plan gratuito para el "
        f"1 de marzo de 2023 y tenía cerrado su programa para código abierto ({cb('free')}), y el instalador de Jet ya no se puede descargar. Por lo tanto, no hay forma "
        "de crear una cuenta ni de ejecutar una configuración de Codeship, ni en la nube ni en el equipo local. Escribir los archivos "
        "codeship-services.yml y codeship-steps.yml habría dejado en el repositorio una configuración que nadie puede ejecutar, así que no se agregaron.")
titulo("Elección del reemplazo", 3)
parrafo(f"Siguiendo la indicación del tutor, se buscó una herramienta que cumpliera las mismas funciones de Codeship y que fuera libre o "
        f"gratuita (Reyes Corredor, 2026). Se compararon dos candidatas: CloudBees Unify, el producto al que CloudBees recomienda migrar, y "
        f"GitHub Actions, el servicio de integración continua de GitHub, que el tutor mencionó en el primer encuentro como ejemplo de un sistema "
        f"que se activa y ejecuta pruebas automáticamente sobre el código que se sube al repositorio (Reyes Corredor, 2026). La Tabla "
        f"{sig('tab')} resume la comparación y la Figura {sig('fig')} muestra la condición de costo de GitHub Actions.", keep=True)
tabla("Comparación de CodeShip Pro con las alternativas", ["Criterio", "CodeShip Pro", "CloudBees Unify", "GitHub Actions"], [
    ["En servicio en noviembre de 2026", "No, retirado en enero de 2026", "Sí", "Sí"],
    ["Dónde se ejecuta", "Nube de CloudBees", f"Infraestructura de CloudBees Unify o ejecutores propios ({cb('flujos')})", f"Ejecutores de GitHub o propios ({gh('facturacion')})"],
    ["Configuración en el repositorio", "codeship-services.yml y codeship-steps.yml", f"Archivos YAML en .cloudbees/workflows/ ({cb('flujos')})", f"Archivos YAML en .github/workflows/ ({gh('sintaxis')})"],
    ["Trabajo con Docker", "Cada servicio es un contenedor", f"Motor de CI/CD basado en Tekton ({cb('faq')})", "El ejecutor ubuntu-24.04 trae Docker 28.0.4 y Docker Compose 2.38.2 (GitHub, 2026)"],
    ["Costo para un repositorio público", f"Plan gratuito retirado en 2023 ({cb('free')})", f"Cotización personalizada; prueba gratuita ({cb('unify')}; {cb('precio')})", f"Gratuito con los ejecutores estándar ({gh('facturacion')})"],
    ["Cuenta adicional", "Sí", "Sí", "No, usa la cuenta de GitHub del repositorio"],
    ["Mencionado por el tutor", "No", "No", "Sí"],
], [1.5, 1.5, 1.8, 1.7])
figura("t04-actions-gratis-publicos.png", "Condición de costo de GitHub Actions para repositorios públicos",
       "La documentación de facturación de GitHub Actions dice que su uso es gratuito con los ejecutores estándar alojados por GitHub en "
       "repositorios públicos y con ejecutores propios. El repositorio del proyecto es público, así que las ejecuciones no tienen costo.", ancho=5.6)
parrafo(f"GitHub Actions es el equivalente más fiel por cuatro razones. Primera, conserva lo esencial de Codeship: un servicio en la nube, "
        "sin servidor propio, que se activa con cada cambio y se configura con archivos YAML versionados junto al código. Segunda, como "
        "CodeShip Pro, ejecuta los pasos con Docker: las pruebas unitarias corren dentro de un contenedor node:22-alpine y las de integración "
        "usan el mismo docker-compose.yml, que cumple el papel que tenía codeship-services.yml. Tercera, es gratuito para un repositorio "
        "público; CloudBees Unify, aunque es el sucesor oficial y también usa flujos YAML en el repositorio, es una plataforma empresarial con "
        "precio por cotización, lo que va en contra de la indicación del tutor. Cuarta, no exige crear otra cuenta, porque funciona con la de "
        f"GitHub. Además, CloudBees Unify puede mostrar los flujos y las ejecuciones de GitHub Actions ({cb('gha')}), de modo que la elección no "
        "cierra la puerta al producto que recomienda CloudBees.")
titulo("Equivalencia entre CodeShip Pro y GitHub Actions", 3)
parrafo(f"La Tabla {sig('tab')} muestra cómo se traduce cada concepto de CodeShip Pro al flujo de GitHub Actions del proyecto.", keep=True)
tabla("Conceptos de CodeShip Pro y su equivalente en el proyecto", ["CodeShip Pro", "GitHub Actions", "En este proyecto"], [
    [f"codeship-services.yml: servicios que corren como contenedores, con sintaxis parecida a docker-compose.yml ({cb('services')})",
     "Contenedores que el flujo levanta con Docker", "El docker-compose.yml de la primera parte, sin cambios"],
    [f"codeship-steps.yml: pasos en orden, cada uno con su servicio y su comando ({cb('steps')})", "jobs y steps con run",
     "Dos jobs, pruebas-unitarias y contenedores, con sus pasos"],
    ["Paso que corre dentro de un servicio", "docker run o docker compose exec",
     "Pruebas unitarias dentro de node:22-alpine y pruebas de integración dentro de tareas-app"],
    [f"Tiempo límite por paso ({cb('steps')})", f"timeout-minutes por job o por paso ({gh('sintaxis')})", "15 minutos por job y 3 para las pruebas unitarias"],
    [f"Pasos que se ejecutan ante una falla ({cb('steps')})", f"Condición if con always() o failure() ({gh('expresiones')})", "La limpieza, con if: always()"],
    [f"Jet, para ejecutar la configuración en el equipo local ({cb('free')})", "No aplica", "El script tools/simular_ci.py, descrito más adelante"],
], [2.4, 1.9, 2.2])
titulo("El flujo de GitHub Actions explicado", 3)
parrafo(f"Un flujo de trabajo es un proceso automatizado configurable, formado por uno o más jobs, que se define en un archivo YAML guardado "
        f"en la carpeta .github/workflows del repositorio ({gh('sintaxis')}). El Listado {sig('lis')} contiene el flujo del proyecto.", keep=True)
listado("Archivo .github/workflows/integracion-continua.yml", r".github\workflows\integracion-continua.yml")
parrafo(f"**on** indica cuándo se ejecuta: con cada push a cualquier rama, con cada pull request y a mano (workflow_dispatch); el disparador "
        f"manual solo funciona cuando el archivo ya está en la rama por defecto, main ({gh('sintaxis')}). **permissions** limita el acceso del "
        f"token GITHUB_TOKEN a leer el contenido del repositorio ({gh('sintaxis')}), porque el flujo no necesita escribir nada. **concurrency** agrupa las ejecuciones de una "
        f"misma rama y, con cancel-in-progress, cancela la que esté en curso cuando llega un commit nuevo ({gh('sintaxis')}).")
parrafo(f"El job **pruebas-unitarias** corre en un ejecutor ubuntu-24.04, descarga el código con actions/checkout y ejecuta npm ci y npm test "
        f"dentro de un contenedor node:22-alpine, igual que la etapa de Jenkins, con un tiempo límite de tres minutos. El job **contenedores** "
        f"tiene needs: pruebas-unitarias, así que GitHub Actions solo lo ejecuta si el primero termina bien; si el primero falla o se salta, "
        f"este también se salta ({gh('sintaxis')}). Sus pasos escriben el .env con una clave de un solo uso, construyen la imagen, levantan los "
        f"dos contenedores, consultan /salud y ejecutan las pruebas de integración. Cada paso se ejecuta solo si los anteriores tuvieron éxito, "
        f"salvo que tenga otra condición ({gh('expresiones')}); el último tiene if: always(), que lo ejecuta siempre ({gh('expresiones')}), así que "
        "baja los contenedores y borra el .env aunque algo haya fallado, como la sección post del Jenkinsfile.")
parrafo(f"**timeout-minutes** fija un máximo de 15 minutos por job; sin él, GitHub cancela un job después de 360 minutos ({gh('sintaxis')}). "
        "El ejecutor ubuntu-24.04 trae Docker 28.0.4, Docker Compose 2.38.2 y Node.js 22.23.4, según la documentación de su imagen "
        "20261111.370.1 (GitHub, 2026).")
titulo("Ejecución en GitHub Actions", 3)
parrafo("GitHub Actions está habilitado en el repositorio, según la consulta a la API de GitHub del 19 de noviembre de 2026, y el "
        "disparador push aplica a cualquier rama; por eso el flujo se ejecutará solo en cuanto se publique en GitHub la rama de esta entrega.")
parrafo("[[COMPLETAR: captura de la ejecución en GitHub Actions después de publicar la rama (los dos jobs en verde y el paso Verificar /salud)]]")

titulo("Validación local de las dos configuraciones", 2)
parrafo("Antes de conectar los servicios se comprobó en el equipo del proyecto que los dos archivos están bien escritos y que sus comandos "
        "funcionan con el proyecto real. Se hizo en dos niveles.")
titulo("Revisión de la sintaxis", 3)
parrafo(f"El script tools/validar_ci.py lee los dos archivos, los valida contra el esquema oficial de Travis CI, publicado en el repositorio "
        f"travis-ci/travis-yml (Travis CI, 2026), y contra el esquema de flujos de GitHub Actions de SchemaStore (SchemaStore, s. f.), y "
        f"comprueba que ambos contengan los pasos del Jenkinsfile en el mismo orden. Además, el flujo de GitHub Actions se revisó con "
        f"actionlint 1.7.12 (rhysd, 2026), que analiza también con ShellCheck los comandos de Bash de cada paso, y los comandos de .travis.yml "
        f"se revisaron con el mismo ShellCheck. La Figura {sig('fig')} muestra el resultado.", keep=True)
figura("t05-validacion-ci.png", "Validación de la sintaxis de .travis.yml y del flujo de GitHub Actions",
       "Los diez controles del validador están en OK, con cero fallas: los dos archivos son YAML válido, cumplen sus esquemas, tienen las "
       "etapas en orden, contienen los pasos del Jenkinsfile y no escriben ninguna clave. actionlint revisa el único flujo del repositorio y "
       "no encuentra errores, y ShellCheck no reporta hallazgos en los comandos de Travis CI.")
parrafo("El esquema de Travis CI tiene un límite que conviene conocer. En una prueba con errores introducidos a propósito, detectó un "
        "sistema operativo inexistente y una distribución con formato inválido, pero aceptó un bloque script escrito como mapa y un valor "
        "services numérico, porque el esquema admite varias formas de escribir la configuración que Travis CI luego normaliza (Travis CI, "
        f"2026). Por eso la revisión de la sintaxis se complementó con la ejecución de los comandos, y la revisión definitiva será la pestaña "
        f"View Config del primer build en Travis CI ({tr('validacion')}).")
titulo("Ejecución de los comandos en el equipo", 3)
parrafo(f"El script tools/simular_ci.py lee .travis.yml o el flujo de GitHub Actions y ejecuta sus comandos en orden, con Bash y con el "
        f"Docker del equipo, respetando las reglas de cada servicio: en Travis CI, las fases y las etapas descritas antes; en GitHub Actions, "
        f"que un paso fallido salta los siguientes salvo los que tienen if: always(), y que un job fallido salta los que dependen de él. No "
        f"reemplaza la ejecución en la nube, porque no usa las máquinas de Travis CI ni de GitHub, pero demuestra que los comandos escritos en "
        f"los archivos funcionan con el proyecto real. Las Figuras {sig('fig')} y {sig('fig') + 1} muestran los resultados.")
figura("t06-travis-local.png", "Ejecución en el equipo de los comandos de .travis.yml",
       "Primero corre la etapa pruebas unitarias: npm ci instala 82 paquetes y npm test termina con cuatro pruebas aprobadas y una omitida, la "
       "de integración, porque en esa etapa no hay base de datos. Después corre la etapa contenedores: se construye la imagen, Docker Compose "
       "reporta tareas-db y tareas-app como Healthy, /salud responde con la versión de PostgreSQL leída de la base y las cinco pruebas pasan "
       "dentro del contenedor de la API. La limpieza borra el volumen y la red, y las dos etapas terminan en passed.", ancho=5.0)
figura("t07-actions-local.png", "Ejecución en el equipo de los comandos del flujo de GitHub Actions",
       "El job pruebas-unitarias omite actions/checkout, porque el código ya está en la carpeta, y corre las pruebas dentro de node:22-alpine, "
       f"con Node.js 22.23.5: cuatro aprobadas y una omitida. Como terminó bien, corre el job contenedores, con el mismo resultado que en la "
       f"Figura {F('t06-travis-local')}: imagen construida, contenedores saludables, /salud con PostgreSQL 16.14 y cinco pruebas aprobadas. El "
       "último paso baja los contenedores. Los dos jobs terminan en success.", ancho=5.0)

titulo("Historial de cambios", 2)
autores = sorted({f[3] for f in HISTORIAL})
parrafo(f"El historial de cambios del proyecto es el historial de Git del repositorio. La Tabla {sig('tab')} lo reproduce completo, del primer "
        f"commit al más reciente, tal como lo entrega el comando git log al generar este documento: {len(HISTORIAL)} commits. Cada fila tiene "
        "el identificador corto del commit, la fecha, el autor, el tipo de cambio según la convención usada desde la primera parte (feat para "
        "funciones nuevas, fix para correcciones, test para pruebas, ci para integración continua y docs para documentación), la descripción "
        "y la entrega a la que pertenece.", keep=True)
tabla("Historial de commits del repositorio", ["N.º", "Commit", "Fecha", "Autor", "Tipo", "Descripción", "Entrega"], HISTORIAL,
      [0.35, 0.7, 0.9, 0.85, 0.45, 2.65, 0.6])
parrafo("Tres observaciones ayudan a leerlo. Los commits del 8 de noviembre con un número entre paréntesis, de (#1) a (#3), son solicitudes "
        "de cambio (pull requests) integradas en la rama main desde GitHub; la solicitud #4, que completa un nombre en la portada, seguía "
        "abierta en la consulta a GitHub del 19 de noviembre de 2026. Los commits 335a28e, 34592c9 y 8e708a8 son la prueba de la segunda "
        "parte: el error introducido a propósito, la corrección de las pruebas y del pipeline, y la reversión. Y los nombres de autor "
        f"{' y '.join(autores)} corresponden a la misma dirección de correo en Git, la de Carlos Fernando Guerra Arango. Los commits que se "
        "hagan después de generar el documento no aparecen en la tabla. El historial de ejecuciones de Jenkins está en "
        "la Tabla 8 de la segunda parte.")

titulo("Sugerencias para la solución de problemas", 2)
parrafo(f"La Tabla {sig('tab')} reúne los problemas que aparecieron durante las tres entregas, con la causa encontrada, cómo se resolvió y "
        "una sugerencia para que no se repita. Los dos últimos son preventivos: no ocurrieron en el proyecto, pero la documentación oficial "
        "advierte sobre ellos.", keep=True)
tabla("Problemas encontrados, causas, soluciones y sugerencias", ["Problema", "Causa", "Solución aplicada", "Sugerencia"], [
    ["La ejecución 2 de Jenkins quedó esperando siete minutos y hubo que abortarla.",
     "La prueba que falló dejaba abierto el servidor de prueba: la instrucción que lo cerraba estaba después de la aserción.",
     "Commit 34592c9: cierre con t.after() y tiempos límite en el Jenkinsfile.",
     "Liberar los recursos de cada prueba en t.after() o en finally y poner tiempo límite a cada etapa; Travis CI y GitHub Actions ya lo tienen."],
    ["La ejecución 3 de Jenkins terminó en rojo en las pruebas unitarias.",
     "El commit 335a28e rompió a propósito la regla del título obligatorio: la API respondía 500 donde la prueba esperaba 400.",
     "Commit 8e708a8, que revierte el cambio.",
     "Leer en la consola la línea not ok, con el valor esperado y el obtenido; si un cambio rompe una regla del negocio, revertirlo y corregirlo en un commit nuevo, sin desactivar la prueba."],
    ["El puerto 8080, el de Jenkins por defecto, estaba ocupado.", "Otra aplicación del equipo lo usaba.",
     "Jenkins se publicó en el 8181, un parámetro del script de instalación.",
     "Dejar los puertos como parámetros de los scripts y no escritos en varios lugares."],
    ["Jenkins no podía leer la rama de la Entrega 2, que no estaba publicada en GitHub.",
     "El plugin Git bloquea por seguridad los repositorios locales (SECURITY-2478).",
     "Copia local de solo lectura con la propiedad ALLOW_LOCAL_CHECKOUT, como medida temporal.",
     "Publicar la rama, apuntar el trabajo a https://github.com/Owito/poli-ic-proyecto1.git y retirar la propiedad."],
    ["Al iniciar Docker Desktop arrancaron solos el contenedor jenkins y contenedores de otros proyectos.",
     "Esos contenedores tienen políticas de reinicio; la de jenkins es on-failure.",
     "Se detuvo jenkins con docker stop y, al terminar el trabajo, se detuvo Docker Desktop.",
     "Después de iniciar Docker Desktop, revisar docker ps y detener lo que no se necesita, para no consumir memoria ni ocupar puertos."],
    ["ShellCheck reportó el error SC1017 en todas las líneas de los comandos de Travis CI.",
     "El archivo intermedio se generó en Windows con finales de línea CRLF, que Bash no acepta.",
     "El validador ahora escribe esos comandos con finales de línea LF.",
     "En Windows, guardar los scripts de Bash con finales de línea LF."],
    ["Travis CI no ofrece hoy una prueba gratuita.",
     f"Cambio en los planes de Travis CI, que además pide validar una tarjeta de crédito ({tr('facturacion')}).",
     "La configuración quedó lista y validada en el equipo, a la espera de la cuenta.",
     "Pedir al soporte créditos para código abierto o el plan del GitHub Student Developer Pack antes de considerar un plan de pago."],
    ["Codeship no se puede usar.",
     f"CloudBees anunció su retiro definitivo para el 31 de enero de 2026 ({cb('resumen')}); hoy ni la aplicación ni Jet están disponibles.",
     "Reemplazo por GitHub Actions, con las mismas etapas.",
     "Antes de planear el uso de una herramienta, revisar su estado en la documentación oficial."],
    ["El esquema de Travis CI aceptó errores introducidos a propósito.",
     "El esquema admite varias formas de escribir la configuración, porque Travis CI la normaliza.",
     "Se complementó con la ejecución de los comandos en el equipo.",
     f"Revisar los mensajes de la pestaña View Config del primer build ({tr('validacion')})."],
    ["Preventivo: límite de descargas de Docker Hub en Travis CI.",
     f"La documentación de Travis CI advierte que, por ese límite, habrá que agregar credenciales propias a la configuración ({tr('docker')}).",
     "No se ha necesitado.",
     "Si una ejecución falla al descargar node:22-alpine o postgres:16-alpine, iniciar sesión en Docker Hub con variables de la configuración de Travis CI."],
    ["Preventivo: el subcomando docker compose no existe en la máquina de Travis CI.",
     f"La documentación de Travis CI menciona el programa docker-compose ({tr('docker')}).",
     "El .travis.yml enlaza el programa como complemento de Docker si hace falta.",
     "Imprimir docker --version y docker compose version al inicio de cada ejecución para detectar el cambio pronto."],
], [1.5, 1.7, 1.5, 1.8])

titulo("Responsabilidades y opiniones", 2)
parrafo(f"El enunciado pide consolidar las responsabilidades y las opiniones de los integrantes (Politécnico Grancolombiano, s. f.). La Tabla "
        f"{sig('tab')} las reúne. Las responsabilidades de Carlos Fernando Guerra Arango salen del historial de Git, en el que todos los commits "
        "están registrados con su usuario; las de los demás integrantes y todas las opiniones las escribe cada persona. Según la consulta a la "
        "API de GitHub del 19 de noviembre de 2026, el repositorio tiene, además de la cuenta propietaria, tres colaboradores con permiso de escritura.", keep=True)
INTEGRANTES = ["Leonardo Espinosa Rivera", "Daniel Santiago Estupiñán Cruz", "Anyelo Brayan Galindo Delgado", "Jerlinson Fabián Ramos Galíndez"]
filas_resp = [[n, f"[[COMPLETAR: responsabilidades de {n}]]", f"[[COMPLETAR: opinión de {n}]]"] for n in INTEGRANTES]
filas_resp.append(["Carlos Fernando Guerra Arango",
                   f"Autor de los {len(HISTORIAL)} commits del historial: la API y sus pruebas, los contenedores, la publicación en GitHub, "
                   "Jenkins, Travis CI, GitHub Actions y el documento de las tres partes.",
                   "[[COMPLETAR: opinión de Carlos Fernando Guerra Arango]]"])
tabla("Responsabilidades y opiniones de los integrantes", ["Integrante", "Responsabilidades", "Opinión"], filas_resp, [1.5, 2.6, 2.4])

titulo("Cómo reproducir la tercera parte", 2)
parrafo(f"Para Travis CI se crea la cuenta, se conecta el repositorio con los pasos de la Tabla {TABLA_TRAVIS} y se sube un commit. Para "
        "GitHub Actions basta con publicar la rama en GitHub: el flujo se ejecuta solo y su resultado se consulta en la pestaña Actions del "
        "repositorio. Para revisar los archivos sin conectar nada se ejecuta python tools/validar_ci.py con los dos esquemas y, con Docker "
        "Desktop en ejecución, python tools/simular_ci.py travis o python tools/simular_ci.py actions.")

# ================================================================ CONCLUSIONES
salto()
titulo("Conclusiones", 1)
parrafo("Se construyó una aplicación de dos contenedores, una API en Node.js y una base de datos PostgreSQL, que se ejecuta con un solo comando y que se comunica a través de una red definida por el proyecto. "
        "La comunicación se demostró con evidencia verificable: la red contiene a ambos contenedores, el nombre db se resuelve desde la API, y los datos que devuelve la API coinciden con los de la base de datos. "
        "Además, como el repositorio de GitHub guarda la definición de los contenedores y no las imágenes, cualquier integrante obtiene el mismo ambiente "
        "a partir de una copia del repositorio y un solo comando, que es la conexión entre GitHub y Docker en la que se apoyan la segunda y la tercera parte.")
parrafo("El uso de healthcheck y depends_on con condición evita un error común de arranque, en el que la aplicación intenta conectarse antes de que la base de datos esté lista. El volumen asegura que los datos "
        "persistan, y la separación de configuración en .env permite publicar el proyecto sin exponer claves. El Dockerfile multietapa con usuario sin privilegios produce una imagen más pequeña y más segura.")
parrafo("Las pruebas automatizadas, escritas en la primera parte para que pudieran ejecutarse desde cualquier herramienta, son las que ahora ejecuta "
        "Jenkins, sin cambios en la aplicación. Como aprendizaje del grupo, la integración continua depende de contar primero con un entorno "
        "reproducible; Docker aporta esa base.")
parrafo("Con Jenkins, la integración continua pasó de ser una intención a ser un proceso que se ejecuta solo: cada commit en la rama se prueba, "
        "se construye en contenedores y se verifica de punta a punta, y el historial de ejecuciones muestra en qué etapa falló cada cambio. "
        "Escribir el pipeline en un Jenkinsfile versionado le da al proceso de integración el mismo control de cambios que tiene el código.")
parrafo("La prueba con un error introducido a propósito mostró el valor de la herramienta en dos niveles. Detuvo la integración antes de "
        "construir y desplegar un cambio que rompía una regla del negocio, y además reveló un defecto en las propias pruebas, que dejaban la "
        "ejecución esperando sin fin. La lección para el grupo es que un pipeline necesita tiempos límite y pruebas que siempre liberen sus recursos.")
parrafo("Instalar Jenkins con Docker fue coherente con la primera parte: Jenkins, su motor de Docker y los contenedores del proyecto se levantan "
        "a partir de archivos del repositorio. Además, Jenkins se integra con Docker en dos niveles: ejecuta etapas dentro de contenedores y "
        "construye y levanta los contenedores de la aplicación.")
parrafo("Como limitaciones, Jenkins corre en un equipo local que no es accesible desde internet, por lo que no puede recibir avisos de GitHub y "
        "revisa el repositorio cada dos minutos; la rama de la Entrega 2 se leyó de una copia local mientras no estaba publicada en GitHub; y la "
        "instalación tiene un solo nodo.")
parrafo("En la tercera parte, la misma secuencia de verificación quedó escrita para dos servicios en la nube, Travis CI y GitHub Actions, "
        "sin cambiar una línea de la aplicación, del Dockerfile ni del docker-compose.yml. Eso confirma la idea que guió el proyecto desde la "
        "primera parte: si el ambiente es reproducible con contenedores y las pruebas no dependen de una herramienta, pasar de una herramienta "
        "de integración continua a otra consiste en traducir unas pocas etapas a otro archivo de configuración.")
parrafo("El caso de Codeship dejó una lección práctica: los servicios de integración continua en la nube cambian sus planes y pueden "
        "desaparecer, así que antes de comprometer el proyecto con uno conviene revisar su estado en la fuente oficial. Cuando una herramienta "
        "ya no existe o exige pagar, la indicación del tutor de buscar una alternativa libre o gratuita que cumpla la misma función permitió cumplir el "
        "objetivo: el proyecto queda integrado con contenedores, Jenkins, Travis CI y un reemplazo fiel de Codeship, todos con las mismas etapas.")
parrafo("Como limitaciones de la tercera parte, la ejecución en Travis CI depende de que el grupo active una cuenta con créditos, y la de "
        "GitHub Actions, de que se publique la rama en GitHub; mientras tanto, la evidencia de ejecución de las dos configuraciones es la "
        "ejecución de sus comandos en el equipo del proyecto.")
# Párrafo del grupo (experiencia y distribución del trabajo): lo escribe el grupo y va aquí. Si no llega antes del
# cierre, se omite y la sección queda completa sin él.

# ================================================================ REFERENCIAS
salto()
titulo("Referencias", 1)
WB = "https://web.archive.org/web/"
CB_DOCS = "https://docs.cloudbees.com/docs/cloudbees-codeship/latest/"
R19 = "Recuperado el 19 de noviembre de 2026, de "
REF = [
    f"CloudBees. (s. f.-a). *CloudBees CodeShip Free - EOL*. CloudBees Docs. {R19}{WB}20250528223552/{CB_DOCS}general-account/free-eol",
    f"CloudBees. (s. f.-b). *CloudBees CodeShip product end-of-life announcement*. CloudBees Docs. {R19}https://docs.cloudbees.com/docs/cloudbees-common/latest/cloudbees-codeship-eol",
    f"CloudBees. (s. f.-c). *CloudBees CodeShip retirement FAQ*. CloudBees Docs. {R19}{WB}20251206050201/{CB_DOCS}r4-finalization/faq",
    f"CloudBees. (s. f.-d). *CloudBees Unify is the control plane for your enterprise software factory*. {R19}https://www.cloudbees.com/unify",
    f"CloudBees. (s. f.-e). *Display GitHub Actions workflows and runs*. CloudBees Unify Docs. {R19}https://docs.cloudbees.com/docs/cloudbees-unify/latest/continuous-integration/how-to-guides/display-github-actions-workflows",
    f"CloudBees. (s. f.-f). *Executive summary* [CloudBees CodeShip retirement]. CloudBees Docs. {R19}{WB}20251203162638/{CB_DOCS}r1-codeship-retirement/",
    f"CloudBees. (s. f.-g). *Get a custom CloudBees Unify pricing quote*. {R19}https://www.cloudbees.com/pricing",
    f"CloudBees. (s. f.-h). *Installing the Jet CLI*. CloudBees Docs. {R19}{WB}20251015113940/{CB_DOCS}pro-jet-cli/installation",
    f"CloudBees. (s. f.-i). *Services configuration*. CloudBees Docs. {R19}{WB}20250612120933/{CB_DOCS}pro-builds-and-configuration/services",
    f"CloudBees. (s. f.-j). *Steps configuration*. CloudBees Docs. {R19}{WB}20250602104044/{CB_DOCS}pro-builds-and-configuration/steps",
    f"CloudBees. (s. f.-k). *Understanding workflows*. CloudBees Unify Docs. {R19}https://docs.cloudbees.com/docs/cloudbees-unify/latest/workflows/explanation/understanding-workflows",
    "Docker Inc. (s. f.-a). *Compose file reference*. Docker Docs. Recuperado el 5 de noviembre de 2026, de https://docs.docker.com/reference/compose-file/",
    "Docker Inc. (s. f.-b). *Docker Compose*. Docker Docs. Recuperado el 5 de noviembre de 2026, de https://docs.docker.com/compose/",
    "Docker Inc. (s. f.-c). *Dockerfile reference*. Docker Docs. Recuperado el 5 de noviembre de 2026, de https://docs.docker.com/reference/dockerfile/",
    "Docker Inc. (s. f.-d). *Networking overview*. Docker Docs. Recuperado el 5 de noviembre de 2026, de https://docs.docker.com/engine/network/",
    "Docker Inc. (s. f.-e). *postgres: Official image*. Docker Hub. Recuperado el 5 de noviembre de 2026, de https://hub.docker.com/_/postgres",
    "Docker Inc. (s. f.-f). *Volumes*. Docker Docs. Recuperado el 5 de noviembre de 2026, de https://docs.docker.com/engine/storage/volumes/",
    "Express. (s. f.). *Express: Node.js web application framework*. Recuperado el 5 de noviembre de 2026, de https://expressjs.com/",
    f"GitHub. (s. f.-a). *Billing and usage* [GitHub Actions]. GitHub Docs. {R19}https://docs.github.com/en/actions/concepts/billing-and-usage",
    f"GitHub. (s. f.-b). *Evaluate expressions in workflows and actions*. GitHub Docs. {R19}https://docs.github.com/en/actions/reference/workflows-and-actions/expressions",
    f"GitHub. (s. f.-c). *GitHub Student Developer Pack*. GitHub Education. {R19}https://education.github.com/pack",
    "GitHub. (s. f.-d). *Inviting collaborators to a personal repository*. GitHub Docs. Recuperado el 5 de noviembre de 2026, de "
    "https://docs.github.com/en/account-and-profile/setting-up-and-managing-your-personal-account-on-github/managing-access-to-your-personal-repositories/inviting-collaborators-to-a-personal-repository",
    "GitHub. (s. f.-e). *Repositories documentation*. GitHub Docs. Recuperado el 5 de noviembre de 2026, de https://docs.github.com/en/repositories",
    f"GitHub. (s. f.-f). *Workflow syntax for GitHub Actions*. GitHub Docs. {R19}https://docs.github.com/en/actions/reference/workflows-and-actions/workflow-syntax",
    f"GitHub. (2026). *Ubuntu 24.04* [Documentación de la imagen del ejecutor, versión 20261111.370.1]. Repositorio actions/runner-images. {R19}"
    "https://github.com/actions/runner-images/blob/main/images/ubuntu/Ubuntu2404-Readme.md",
    "Jenkins. (s. f.-a). *Docker*. Jenkins User Documentation: Installing Jenkins. Recuperado el 17 de noviembre de 2026, de https://www.jenkins.io/doc/book/installing/docker/",
    "Jenkins. (s. f.-b). *Git*. Jenkins Plugins. Recuperado el 17 de noviembre de 2026, de https://plugins.jenkins.io/git/",
    "Jenkins. (s. f.-c). *Jenkins: Build great things at any scale*. Recuperado el 17 de noviembre de 2026, de https://www.jenkins.io/",
    "Jenkins. (s. f.-d). *Pipeline*. Jenkins User Documentation. Recuperado el 17 de noviembre de 2026, de https://www.jenkins.io/doc/book/pipeline/",
    "Jenkins. (s. f.-e). *Pipeline syntax*. Jenkins User Documentation. Recuperado el 17 de noviembre de 2026, de https://www.jenkins.io/doc/book/pipeline/syntax/",
    "Jenkins. (s. f.-f). *Using credentials*. Jenkins User Documentation. Recuperado el 17 de noviembre de 2026, de https://www.jenkins.io/doc/book/using/using-credentials/",
    "Jenkins. (s. f.-g). *Using Docker with Pipeline*. Jenkins User Documentation. Recuperado el 17 de noviembre de 2026, de https://www.jenkins.io/doc/book/pipeline/docker/",
    "Node.js. (s. f.). *Test runner*. Node.js v22 documentation. Recuperado el 5 de noviembre de 2026, de https://nodejs.org/docs/latest-v22.x/api/test.html",
    "Politécnico Grancolombiano. (s. f.). *Proyecto de software basado en herramientas de integración continua* [Enunciado del proyecto del "
    "módulo Énfasis Profesional I (Integración Continua)]. Institución Universitaria Politécnico Grancolombiano.",
    "PostgreSQL Global Development Group. (s. f.). *PostgreSQL 16 documentation*. Recuperado el 5 de noviembre de 2026, de https://www.postgresql.org/docs/16/",
    "Reyes Corredor, E. (2026, 19 y 26 de octubre, y 5 y 9 de noviembre). *Encuentros sincrónicos de Énfasis Profesional I (Integración Continua)* [Grabaciones de clase]. Politécnico Grancolombiano.",
    f"rhysd. (2026). *actionlint* (versión 1.7.12) [Software]. GitHub. {R19}https://github.com/rhysd/actionlint",
    f"SchemaStore. (s. f.). *github-workflow.json* [Esquema JSON de los flujos de GitHub Actions]. {R19}https://json.schemastore.org/github-workflow.json",
    f"Travis CI. (s. f.-a). *Billing FAQ*. Travis CI Docs. {R19}https://docs.travis-ci.com/user/billing-faq/",
    f"Travis CI. (s. f.-b). *Billing overview*. Travis CI Docs. {R19}https://docs.travis-ci.com/user/billing-overview/",
    f"Travis CI. (s. f.-c). *Build a JavaScript and Node.js project*. Travis CI Docs. {R19}https://docs.travis-ci.com/user/languages/javascript-with-nodejs/",
    f"Travis CI. (s. f.-d). *Build config validation*. Travis CI Docs. {R19}https://docs.travis-ci.com/user/build-config-validation/",
    f"Travis CI. (s. f.-e). *Build stages*. Travis CI Docs. {R19}https://docs.travis-ci.com/user/build-stages/",
    f"Travis CI. (s. f.-f). *CI/CD pricing plans for any need*. {R19}https://www.travis-ci.com/pricing/",
    f"Travis CI. (s. f.-g). *Job lifecycle*. Travis CI Docs. {R19}https://docs.travis-ci.com/user/job-lifecycle/",
    f"Travis CI. (s. f.-h). *Travis CI onboarding*. Travis CI Docs. {R19}https://docs.travis-ci.com/user/onboarding/",
    f"Travis CI. (s. f.-i). *The Ubuntu 24.04 (Noble Numbat) build environment*. Travis CI Docs. {R19}https://docs.travis-ci.com/user/reference/noble/",
    f"Travis CI. (s. f.-j). *Use Docker in builds*. Travis CI Docs. {R19}https://docs.travis-ci.com/user/docker/",
    f"Travis CI. (2026). *travis-yml* [Repositorio con el esquema JSON de la configuración de Travis CI]. GitHub. {R19}https://github.com/travis-ci/travis-yml",
]
for r in REF:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(-0.5)
    rico(p, r)

# La fuente de las referencias a GitHub (a) no se cita en el texto: se cita en la sección de repositorio.
doc.save(SALIDA)
print("Guardado", SALIDA, "figuras:", contador["fig"], "listados:", contador["lis"], "tablas:", contador["tab"])
