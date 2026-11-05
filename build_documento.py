"""Genera docs/Entrega1_IC_Proyecto1.docx (APA 7) a partir de las evidencias reales.

Uso:
    python tools/diagrama.py          (diagrama)
    bash tools/capturar.sh            (ejecuta los comandos y genera los PNG; requiere Docker)
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
SALIDA = RAIZ / "docs" / "Entrega1_IC_Proyecto1.docx"
FUENTE = "Times New Roman"

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
         "12-crud", "13-psql-verifica", "14-validacion", "17-persistencia", "15-pruebas"]


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
    pn.add_run("Elaboración propia." if png.startswith("diagrama") else
               "Captura de la salida real del comando, generada durante la elaboración de este trabajo.")
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
centrado("Proyecto de software con Docker: API de tareas del equipo en dos contenedores comunicados", negrita=True)
centrado("Entrega 1, Proyecto 1, Escenario 3")
doc.add_paragraph()
centrado("Leonardo Espinosa Rivera")
centrado("Daniel Santiago Estupiñán Cruz")
centrado("Anyelo Brayan Galindo Delgado")
centrado("Jerlinson Ramos Galindez")
centrado("Carlos Fernando Guerra Arango")
centrado("[CONFIRMAR: en la portada solo van quienes trabajaron en la entrega; borrar los nombres que no apliquen]", resaltar=True, simple=True)
doc.add_paragraph()
centrado("Códigos estudiantiles: [CÓDIGO LEONARDO] [CÓDIGO DANIEL] [CÓDIGO ANYELO] [CÓDIGO JERLINSON]", resaltar=True, simple=True)
centrado("Carlos Fernando Guerra Arango: 100360894", simple=True)
doc.add_paragraph()
centrado("Programa de Ingeniería de Software")
centrado("Facultad de Ingeniería, Diseño e Innovación")
centrado("Politécnico Grancolombiano")
centrado("TIC 42711: Integración Continua, grupo B04")
centrado("Grupo de trabajo colaborativo 14")
centrado("Tutor: Edward Reyes Corredor")
centrado("Noviembre de 2026")
salto()

# ================================================================ TABLA DE CONTENIDO
p = doc.add_paragraph(style="Heading 1"); p.add_run("Tabla de contenido")
# El título anterior no debe entrar al índice: se le quita el nivel de esquema.
ppr = p._p.get_or_add_pPr(); ol = OxmlElement("w:outlineLvl"); ol.set(qn("w:val"), "9"); ppr.append(ol)
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
parrafo("Este documento corresponde a la Entrega 1 del proyecto del módulo. El enunciado indica que el proyecto se crea en un "
        "repositorio de GitHub y que, con Docker, se construyen dos contenedores comunicados entre sí. Para cumplirlo se desarrolló "
        "una API REST de registro de tareas del equipo con Node.js y Express, que guarda sus datos en una base de datos PostgreSQL. "
        "La API corre en un contenedor y la base de datos en otro; ambos se orquestan con Docker Compose. El dominio es deliberadamente "
        "sencillo, porque las entregas siguientes (Jenkins, Travis CI y Codeship) necesitan una aplicación con pruebas automatizadas "
        "sobre la cual construir los flujos de integración continua.")
parrafo("Cada captura de pantalla que aparece en el documento es la salida real de un comando ejecutado durante la elaboración del trabajo, "
        "y va acompañada de una explicación de lo que muestra y de por qué importa para el proyecto.")
titulo("Objetivo general", 2)
parrafo("Construir y ejecutar con Docker una aplicación de dos contenedores, una API en Node.js y una base de datos PostgreSQL, que se "
        "comuniquen entre sí, y documentar con evidencias el proceso para que sirva de base a las entregas de integración continua.")
titulo("Objetivos específicos", 2)
parrafo("Primero, organizar el proyecto en un repositorio de control de versiones con una estructura clara, listo para publicarse en GitHub.")
parrafo("Segundo, definir un Dockerfile para la API y un archivo de Docker Compose que levante los dos servicios en una red propia, con volumen "
        "persistente para la base de datos, comprobación de salud y orden de arranque.")
parrafo("Tercero, demostrar con comandos reales que los dos contenedores se comunican: que la API consulta la base de datos y que los datos "
        "devueltos provienen de ella.")
parrafo("Cuarto, escribir pruebas automatizadas que puedan ejecutarse más adelante desde las herramientas de integración continua.")

# ================================================================ DESARROLLO
salto()
titulo("Desarrollo", 1)
titulo("Repositorio en GitHub y estructura del proyecto", 2)
titulo("Elección del proyecto y de las herramientas", 3)
parrafo("El tutor indicó que cada grupo escoge libremente el proyecto, siempre que permita aplicar las herramientas del módulo, y que se "
        "prefieren herramientas libres, instalando solo las necesarias para el proyecto (Reyes Corredor, 2026). "
        "Por eso se eligió un registro de tareas del equipo: tiene un modelo de datos de una sola tabla, un CRUD completo y suficiente "
        "lógica de validación para escribir pruebas útiles. Todas las herramientas usadas son libres o de uso gratuito: Git para el control de versiones, "
        "GitHub como servicio del repositorio remoto, Docker y Docker Compose para los contenedores, Node.js 22 con el marco Express para la API "
        "(Express, s. f.) y PostgreSQL 16 como base de datos (PostgreSQL Global Development Group, s. f.). Node.js y Git se instalan en el equipo "
        "porque se usan para escribir el código y las pruebas; PostgreSQL no se instala en el equipo, solo corre dentro de su contenedor.")
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
parrafo("Los repositorios de GitHub se describen en su documentación oficial (GitHub, s. f.-b). El proyecto se publicó en un repositorio privado de "
        f"GitHub, https://github.com/Owito/poli-ic-proyecto1, enlazado como remoto origin del repositorio local. La Figura {F("19-github")} muestra la "
        "consulta del repositorio publicado y su historial en GitHub. [[CONFIRMAR: si otros integrantes agregan commits, el historial mostrará el aporte de cada uno.]]", keep=True)
figura("19-github.png", "Repositorio publicado en GitHub",
       "La primera consulta confirma el nombre, la visibilidad privada, la rama principal main y la dirección del repositorio. La segunda muestra que la "
       "rama main de GitHub apunta al mismo commit que el repositorio local, y la tercera lista el historial tal como quedó publicado.")
parrafo("Para que el tutor pueda revisar el proyecto, el repositorio debe darle acceso como colaborador, es decir, como usuario de desarrollo, tal como lo "
        "pidió en el encuentro del 26 de octubre de 2026 (Reyes Corredor, 2026). GitHub documenta el procedimiento de invitación (GitHub, s. f.-a). "
        "[[PENDIENTE DEL GRUPO: invitar al tutor Edward Reyes Corredor con el usuario de GitHub que él indique y comprobar que puede entrar.]]")

titulo("Arquitectura de los dos contenedores", 2)
parrafo("La solución tiene dos servicios que Docker Compose crea y conecta en una misma red. El primero, app, ejecuta la API de Node.js. "
        "El segundo, db, ejecuta PostgreSQL 16. Solo la API se publica hacia el equipo anfitrión, por el puerto 8095; la base de datos queda "
        "accesible únicamente dentro de la red, lo que reduce la superficie expuesta.", keep=True)
figura("diagrama-arquitectura.png", "Diagrama de la arquitectura de la solución",
       "El cliente (navegador o curl) habla por HTTP con el contenedor tareas-app, que atiende en el puerto interno 3000 y está publicado en el 8095 del equipo. "
       "Ese contenedor se conecta por SQL con tareas-db usando el nombre de servicio db como dirección. Ambos viven en la red red-tareas. La base de datos "
       "guarda sus archivos en el volumen tareas-datos y recibe el script db/init.sql en el primer arranque. La flecha de depends_on indica que la API no arranca hasta que la base "
       "de datos pase su comprobación de salud.")
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
parrafo("Estas mismas pruebas se ejecutarán desde Jenkins en la Entrega 2 y desde Travis CI y Codeship en la Entrega 3, ya que el comando npm test no depende de ninguna herramienta específica.")

titulo("Cómo reproducir el proyecto", 2)
parrafo("El archivo README.md del repositorio lo explica en español. En resumen: se copia .env.example como .env y se cambia la clave, se ejecuta docker compose up -d --build, se verifica con docker compose ps "
        "y se consultan http://localhost:8095/salud y http://localhost:8095/tareas. Para detenerlo se usa docker compose down.")

# ================================================================ CONCLUSIONES
salto()
titulo("Conclusiones", 1)
parrafo("Se construyó una aplicación de dos contenedores, una API en Node.js y una base de datos PostgreSQL, que se ejecuta con un solo comando y que se comunica a través de una red definida por el proyecto. "
        "La comunicación se demostró con evidencia verificable: la red contiene a ambos contenedores, el nombre db se resuelve desde la API, y los datos que devuelve la API coinciden con los de la base de datos.")
parrafo("El uso de healthcheck y depends_on con condición evita un error común de arranque, en el que la aplicación intenta conectarse antes de que la base de datos esté lista. El volumen asegura que los datos "
        "persistan, y la separación de configuración en .env permite publicar el proyecto sin exponer claves. El Dockerfile multietapa con usuario sin privilegios produce una imagen más pequeña y más segura.")
parrafo("Las pruebas automatizadas quedaron listas para integrarse en las siguientes entregas. Como el comando npm test funciona dentro y fuera de Docker, Jenkins, Travis CI y Codeship podrán ejecutarlo sin cambios en la aplicación. "
        "Como aprendizaje del grupo, la integración continua depende de contar primero con un entorno reproducible; Docker aporta esa base.")
parrafo("Como limitaciones, el proyecto aún no tiene un flujo de integración continua, que corresponde a la Entrega 2, y las pruebas de integración dependen de que la base de datos esté disponible. "
        "[[PENDIENTE DEL GRUPO: agregar una conclusión con la experiencia del equipo y la distribución del trabajo, y completar las secciones de GitHub marcadas.]]")

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
    "Node.js. (s. f.). *Test runner*. Node.js v22 documentation. Recuperado el 5 de noviembre de 2026, de https://nodejs.org/docs/latest-v22.x/api/test.html",
    "PostgreSQL Global Development Group. (s. f.). *PostgreSQL 16 documentation*. Recuperado el 5 de noviembre de 2026, de https://www.postgresql.org/docs/16/",
    "Reyes Corredor, E. (2026, 19 y 26 de octubre). *Encuentros sincrónicos de Énfasis Profesional I (Integración Continua)* [Grabaciones de clase]. Politécnico Grancolombiano.",
]
for r in REF:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(-0.5)
    rico(p, r)

# La fuente de las referencias a GitHub (a) no se cita en el texto: se cita en la sección de repositorio.
doc.save(SALIDA)
print("Guardado", SALIDA, "figuras:", contador["fig"], "listados:", contador["lis"], "tablas:", contador["tab"])
