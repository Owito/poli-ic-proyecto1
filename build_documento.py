"""Genera docs/Entrega2_IC_Proyecto.docx (APA 7): el documento del proyecto con la primera parte (Docker, Entrega 1)
y la segunda parte (Jenkins, Entrega 2), a partir de las evidencias reales.

Uso:
    python tools/diagrama.py; python tools/diagrama_jenkins.py      (diagramas)
    bash tools/capturar.sh            (evidencias de la primera parte; requiere Docker)
    bash tools/capturar_jenkins.sh ...   (evidencias de terminal de la segunda parte; ver el script)
    python tools/jenkins_capturas.py ... (capturas de la interfaz de Jenkins; requiere Playwright)
    python build_documento.py         (arma el .docx)
    powershell -File tools/exportar_pdf.ps1   (actualiza la tabla de contenido y exporta el PDF con Word)

Los campos resaltados en amarillo son datos que debe confirmar el grupo.
"""
import re
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
SALIDA = RAIZ / "docs" / "Entrega2_IC_Proyecto.docx"
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
         "j18-github-ejecucion", "j23-consola-github"]
# Evidencias de la segunda parte que son salida de terminal (las demás "j" son capturas de la interfaz de Jenkins).
J_TERMINAL = {"j01-levantar", "j09-cliente-docker", "j13-espejo"}
J_CONSOLA = {"j21-consola-verde", "j22-consola-rojo", "j23-consola-github"}


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
        pr.add_run(e).bold = True
    for fila in filas:
        cs = t.add_row().cells
        for i, v in enumerate(fila):
            cs[i].text = ""
            pr = cs[i].paragraphs[0]; pr.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            rico(pr, v)
    for fila in t.rows:
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
centrado("Proyecto de software basado en herramientas de integración continua: API de tareas del equipo con Docker y Jenkins", negrita=True)
centrado("Entrega 2, Proyecto 2, Escenario 5")
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
centrado("Noviembre de 2026")

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
parrafo("Este documento reúne las dos primeras entregas del proyecto del módulo. Como indicó el tutor, cada entrega se agrega como una "
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
parrafo("Cada captura de pantalla del documento es la salida real de un comando o una vista real de la interfaz de Jenkins, obtenida durante "
        "la elaboración del trabajo, y va acompañada de una explicación de lo que muestra y de por qué importa para el proyecto.")
titulo("Objetivo general", 2)
parrafo("Construir una aplicación de dos contenedores, una API en Node.js y una base de datos PostgreSQL, y automatizar su integración "
        "continua con Jenkins, de modo que cada cambio en el repositorio se pruebe, se construya y se verifique sin pasos manuales, "
        "documentando el proceso con evidencias.")
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
    parrafo("Los repositorios de GitHub se describen en su documentación oficial (GitHub, s. f.-b). El proyecto está publicado en un repositorio público de "
            "GitHub, https://github.com/Owito/poli-ic-proyecto1, enlazado como remoto origin del repositorio local. Se creó como privado y se cambió a público "
            "por recomendación del tutor, para que él pueda entrar a validarlo y para que las herramientas de integración continua muestren los errores donde "
            f"los encuentren (Reyes Corredor, 2026). La Figura {F("19-github")} muestra la consulta del repositorio publicado y su historial en GitHub.", keep=True)
    figura("19-github.png", "Repositorio publicado en GitHub",
           "La primera consulta confirma el nombre, la visibilidad pública, la rama principal main y la dirección del repositorio. La segunda muestra que la "
           "rama main de GitHub apunta al mismo commit que el repositorio local, y la tercera lista el historial tal como quedó publicado. La cuarta pide la "
           "página del repositorio sin iniciar sesión y recibe la respuesta HTTP 200: cualquier persona con el enlace, incluido el tutor, puede abrirlo.")
else:
    parrafo("Los repositorios de GitHub se describen en su documentación oficial (GitHub, s. f.-b). El proyecto se publicó en un repositorio privado de "
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
               "indicó para conformar el equipo de desarrollo y que GitHub documenta en su procedimiento de invitación (GitHub, s. f.-a).")
else:
    ACCESO += (" Para que cada integrante pueda subir sus propios cambios, se le registra como colaborador en Settings, Collaborators con su "
               "usuario de GitHub, la opción que el tutor indicó para conformar el equipo de desarrollo y que GitHub documenta en su "
               "procedimiento de invitación (GitHub, s. f.-a).")
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
parrafo("Estas mismas pruebas son las que ejecuta Jenkins en la segunda parte de este documento, y las que ejecutarán Travis CI y Codeship en la "
        "Entrega 3, ya que el comando npm test no depende de ninguna herramienta específica.")

titulo("Conexión entre GitHub y Docker", 2)
parrafo("El tutor precisó que el alcance de esta entrega es trabajar con GitHub, construir los dos contenedores con Docker y conectar Docker con GitHub "
        "(Reyes Corredor, 2026). En este proyecto la conexión se da a través de los archivos que definen los contenedores. GitHub no guarda las imágenes ni "
        "los contenedores, sino las instrucciones para construirlos: el Dockerfile de la API, el archivo docker-compose.yml con los dos servicios, la red y el "
        f"volumen, el script db/init.sql y la plantilla .env.example, todos visibles en la Figura {F("02-estructura")}. A partir de esos archivos, Docker construye "
        "la imagen de la API en el equipo de quien ejecuta el proyecto y descarga de Docker Hub la imagen oficial postgres:16-alpine (Docker Inc., s. f.-e).")
parrafo("Por eso el repositorio es el punto de encuentro del equipo: cualquier integrante, o el tutor, clona el repositorio y ejecuta docker compose up -d --build "
        "para obtener los mismos dos contenedores, con la misma red y la misma configuración. Cuando alguien modifica el Dockerfile o el docker-compose.yml, el "
        "cambio queda en el historial de commits y los demás lo reciben al actualizar su copia, lo que evita que cada integrante trabaje con una versión distinta "
        "del ambiente. Es también la conexión que aprovecha Jenkins en la segunda parte, y la que aprovecharán Travis CI y Codeship en la Entrega 3, "
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
    ["1", "A mano", "35489a5, el Jenkinsfile", "SUCCESS", "1 min 5 s", "El pipeline completo funciona; la primera vez descarga la imagen de Node.js."],
    ["2", "Cambio en el repositorio", "d06d969, error introducido a propósito", "ABORTED", "7 min 6 s", "La prueba falló, pero la ejecución no terminaba; se abortó a mano."],
    ["3", "Cambio en el repositorio", "25bcac6, pruebas y tiempo límite corregidos", "FAILURE", "7,8 s", "El error sigue y el pipeline se detiene en las pruebas unitarias."],
    ["4", "Cambio en el repositorio", "53d6fe3, revierte el error", "SUCCESS", "19 s", "Todo en verde otra vez."],
], [0.4, 1.0, 1.5, 0.9, 0.8, 1.9])
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
       "git push envía el commit d06d969 a la copia local del repositorio; git ls-remote confirma que la rama apunta a ese commit. Desde ese "
       "momento, la próxima revisión de pollSCM encuentra un commit que no se ha construido y lanza el pipeline.")
titulo("El pipeline detiene la integración cuando una prueba falla", 3)
parrafo("El commit d06d969 se hizo a propósito para comprobar que Jenkins detecta errores: cambia la API para que una tarea sin título reciba un "
        "título por defecto, lo que rompe la regla de negocio de que el título es obligatorio. La prueba unitaria POST /tareas sin titulo espera "
        "la respuesta HTTP 400 y, con el cambio, recibió 500.")
parrafo("En la ejecución 2, Jenkins detectó el fallo, pero la ejecución no terminó: la prueba fallida dejaba abierto el servidor de prueba, porque "
        "la instrucción que lo cerraba estaba después de la aserción y nunca se ejecutaba, y Node.js se quedaba esperando. Tras siete minutos se "
        "abortó a mano. El hallazgo llevó a dos correcciones en el commit 25bcac6: las pruebas cierran el servidor con t.after(), que se ejecuta "
        "aunque la prueba falle, y el Jenkinsfile tiene ahora un tiempo límite general de quince minutos y uno de tres minutos para las pruebas "
        f"unitarias. Con eso, la ejecución 3 terminó en rojo en 7,8 segundos (Figuras {sig('fig')} y {sig('fig') + 1}).")
figura("j17-etapas-rojo.png", "Etapas de la ejecución 3: falla en las pruebas unitarias",
       "La etapa de dependencias y pruebas unitarias está en rojo. Las cuatro siguientes aparecen saltadas: Jenkins no construyó la imagen ni "
       "levantó los contenedores, porque no tiene sentido desplegar código que no pasa sus pruebas. Las acciones finales sí se ejecutaron. A la "
       "derecha, el resumen de npm test: cinco pruebas, tres aprobadas, una fallida y una omitida, y el código de salida 1 que marca la falla.")
figura("j22-consola-rojo.png", "Consola de la ejecución 3",
       "La prueba 1 reporta not ok, con el valor esperado, 400, y el obtenido, 500. Después, cada etapa siguiente aparece como skipped due to "
       "earlier failure(s), se imprime el mensaje de falla definido en post y la ejecución termina con Finished: FAILURE.")
parrafo(f"Finalmente se revirtió el commit con el error (commit 53d6fe3) y la ejecución 4 volvió a quedar en verde. La Figura {sig('fig')} "
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

# ================================================================ CONCLUSIONES
salto()
titulo("Conclusiones", 1)
parrafo("Se construyó una aplicación de dos contenedores, una API en Node.js y una base de datos PostgreSQL, que se ejecuta con un solo comando y que se comunica a través de una red definida por el proyecto. "
        "La comunicación se demostró con evidencia verificable: la red contiene a ambos contenedores, el nombre db se resuelve desde la API, y los datos que devuelve la API coinciden con los de la base de datos. "
        "Además, como el repositorio de GitHub guarda la definición de los contenedores y no las imágenes, cualquier integrante obtiene el mismo ambiente "
        "a partir de una copia del repositorio y un solo comando, que es la conexión entre GitHub y Docker en la que se apoya la segunda parte y se apoyará la Entrega 3.")
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
        "revisa el repositorio cada dos minutos; la rama de esta entrega se leyó de una copia local mientras no estaba publicada en GitHub; y la "
        "instalación tiene un solo nodo. Travis CI y Codeship, que completan la integración del proyecto, corresponden a la Entrega 3.")
# Párrafo del grupo (experiencia y distribución del trabajo): lo escribe el grupo y va aquí. Si no llega antes del
# cierre, se omite y la sección queda completa sin él.

# ================================================================ REFERENCIAS
salto()
titulo("Referencias", 1)
REF = [
    "Docker Inc. (s. f.-a). *Compose file reference*. Docker Docs. Recuperado el 5 de noviembre de 2026, de https://docs.docker.com/reference/compose-file/",
    "Docker Inc. (s. f.-b). *Docker Compose*. Docker Docs. Recuperado el 5 de noviembre de 2026, de https://docs.docker.com/compose/",
    "Docker Inc. (s. f.-c). *Dockerfile reference*. Docker Docs. Recuperado el 5 de noviembre de 2026, de https://docs.docker.com/reference/dockerfile/",
    "Docker Inc. (s. f.-d). *Networking overview*. Docker Docs. Recuperado el 5 de noviembre de 2026, de https://docs.docker.com/engine/network/",
    "Docker Inc. (s. f.-e). *postgres: Official image*. Docker Hub. Recuperado el 5 de noviembre de 2026, de https://hub.docker.com/_/postgres",
    "Docker Inc. (s. f.-f). *Volumes*. Docker Docs. Recuperado el 5 de noviembre de 2026, de https://docs.docker.com/engine/storage/volumes/",
    "Express. (s. f.). *Express: Node.js web application framework*. Recuperado el 5 de noviembre de 2026, de https://expressjs.com/",
    "GitHub. (s. f.-a). *Inviting collaborators to a personal repository*. GitHub Docs. Recuperado el 5 de noviembre de 2026, de "
    "https://docs.github.com/en/account-and-profile/setting-up-and-managing-your-personal-account-on-github/managing-access-to-your-personal-repositories/inviting-collaborators-to-a-personal-repository",
    "GitHub. (s. f.-b). *Repositories documentation*. GitHub Docs. Recuperado el 5 de noviembre de 2026, de https://docs.github.com/en/repositories",
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
]
for r in REF:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(-0.5)
    rico(p, r)

# La fuente de las referencias a GitHub (a) no se cita en el texto: se cita en la sección de repositorio.
doc.save(SALIDA)
print("Guardado", SALIDA, "figuras:", contador["fig"], "listados:", contador["lis"], "tablas:", contador["tab"])
