"""Revisión mecánica del PDF del documento: páginas, figuras, tablas, listados, citas, referencias y pendientes.

Uso: python tools/verificar_pdf.py docs/Entrega3_IC_Proyecto.pdf [marcadores_esperados]
Imprime cada control con OK o FALLA y termina con el número de fallas.

Los pendientes que debe completar el grupo se escriben en build_documento.py como [[COMPLETAR: ...]]; en el documento
salen resaltados en amarillo y sin corchetes, así que aquí se cuentan por la palabra COMPLETAR.
"""
import re
import sys
from pathlib import Path

from pypdf import PdfReader

ruta = Path(sys.argv[1])
esperados = int(sys.argv[2]) if len(sys.argv) > 2 else 0
paginas = [p.extract_text() or "" for p in PdfReader(str(ruta)).pages]
texto = "\n".join(paginas)
plano = re.sub(r"\s+", " ", texto)
# Una cita partida al final de una línea ("s. f.-" y la letra en la línea siguiente) se vuelve a unir.
plano = re.sub(r"s\. f\.- ([a-z])\b", r"s. f.-\1", plano)
fallas = 0


def control(nombre, ok, detalle=""):
    global fallas
    fallas += 0 if ok else 1
    print(("OK    " if ok else "FALLA ") + nombre + (f": {detalle}" if detalle else ""))


print("Páginas:", len(paginas))

for tipo, total in (("Figura", 51), ("Tabla", 17), ("Listado", 7)):
    rotulos = [int(n) for n in re.findall(rf"(?m)^{tipo} (\d+)\s*$", texto)]
    faltan = sorted(set(range(1, total + 1)) - set(rotulos))
    control(f"{tipo}s rotuladas 1 a {total}", not faltan and len(rotulos) == total, f"rótulos {len(rotulos)}, faltan {faltan}")
    # un elemento está citado si su número aparece en el texto fuera de su propio rótulo ("Qué muestra la Figura N" cuenta)
    sin_cita = [n for n in range(1, total + 1) if plano.count(f"{tipo} {n} ") + plano.count(f"{tipo} {n}.") + plano.count(f"{tipo} {n},") + plano.count(f"{tipo} {n})") + plano.count(f"{tipo}s {n} ") < 2]
    control(f"{tipo}s citadas en el texto", not sin_cita, f"sin cita aparte del rótulo: {sin_cita}")

pendientes = plano.count("COMPLETAR")
control(f"Pendientes COMPLETAR: {pendientes}", pendientes == esperados, f"esperados {esperados}")
control("Sin corchetes dobles sueltos", "[[" not in texto and "]]" not in texto)
control("Sin guion largo (U+2014)", "\u2014" not in texto)
control("Sin '??' (referencias rotas)", "??" not in texto)
control("Sin errores de campo de Word", "Error!" not in texto)
# C\u00f3digos estudiantiles: el documento lleva solo nombres. Un c\u00f3digo del Poli es un n\u00famero de 9 o 10 cifras.
# Se excluye solo el identificador de la ejecución de GitHub Actions citada en el documento (no es un código estudiantil).
IDS_PERMITIDOS = {"45919779652"}
codigos = sorted(set(re.findall(r"(?<![\w.,-])\d{8,11}(?![\w.,-])", texto)) - IDS_PERMITIDOS)
control("Sin n\u00fameros de 8 a 11 cifras (c\u00f3digos estudiantiles)", not codigos, f"encontrados {codigos}")
toc = " ".join(paginas[1:5])  # el índice ocupa de la página 2 a la 5
control("Tabla de contenido con las tres partes y la sustentaci\u00f3n",
        all(p in toc for p in ("Primera parte", "Segunda parte", "Tercera parte", "Sustentaci\u00f3n en video")))

# Referencias: cada entrada debe citarse en el texto (autor y sufijo de fecha).
ini = plano.rfind("Referencias")
cuerpo, refs = plano[:ini], plano[ini:]
citas = [("Docker Inc. (s. f.-a)", "s. f.-a"), ("Politécnico Grancolombiano. (s. f.)", "Politécnico Grancolombiano, s. f."),
         ("Reyes Corredor, E. (2026", "Reyes Corredor, 2026"), ("Node.js. (s. f.)", "Node.js, s. f."), ("Express. (s. f.)", "Express, s. f."),
         ("PostgreSQL Global Development Group", "PostgreSQL Global Development Group, s. f."),
         ("SchemaStore. (s. f.)", "SchemaStore, s. f."), ("rhysd. (2026)", "rhysd, 2026"), ("Travis CI. (2026)", "Travis CI, 2026"),
         ("GitHub. (2026)", "GitHub, 2026")]
for autor, letras in (("Jenkins", "abcdefg"), ("CloudBees", "abcdefghijk"), ("GitHub", "abcdef"), ("Travis CI", "abcdefghij")):
    citas += [(f"{autor}. (s. f.-{l})", f"{autor}, s. f.-{l}") for l in letras]
for entrada, cita in citas:
    control(f"Referencia citada: {entrada}", entrada in refs and cita in cuerpo)

for frase in ("Entrega 3 y sustentación, escenarios 7 y 8", "Jerlinson Fabián Ramos Galíndez", "Tutor: Edward Reyes Corredor"):
    control(f"Portada contiene '{frase}'", frase in paginas[0])

print("Total de fallas:", fallas)
