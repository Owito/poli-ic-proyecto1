"""Revisión mecánica del PDF del documento: páginas, figuras, tablas, listados, citas y referencias.

Uso: python tools/verificar_pdf.py docs/Entrega2_IC_Proyecto.pdf
Imprime cada control con OK o FALLA y termina con el número de fallas.
"""
import re
import sys
from pathlib import Path

from pypdf import PdfReader

ruta = Path(sys.argv[1])
paginas = [p.extract_text() or "" for p in PdfReader(str(ruta)).pages]
texto = "\n".join(paginas)
plano = re.sub(r"\s+", " ", texto)
fallas = 0


def control(nombre, ok, detalle=""):
    global fallas
    fallas += 0 if ok else 1
    print(("OK    " if ok else "FALLA ") + nombre + (f": {detalle}" if detalle else ""))


print("Páginas:", len(paginas))

for tipo, total in (("Figura", 41), ("Tabla", 8), ("Listado", 5)):
    rotulos = [int(n) for n in re.findall(rf"(?m)^{tipo} (\d+)\s*$", texto)]
    faltan = sorted(set(range(1, total + 1)) - set(rotulos))
    control(f"{tipo}s rotuladas 1 a {total}", not faltan and len(rotulos) == total, f"rótulos {len(rotulos)}, faltan {faltan}")
    menciones = {int(n) for n in re.findall(rf"{tipo}s? (\d+)", plano)}
    menciones |= {int(n) for n in re.findall(rf"{tipo}s (\d+) y (\d+)", plano) for n in n}
    # una figura está citada si su número aparece en el texto fuera de su propio rótulo ("Qué muestra la Figura N" cuenta)
    sin_cita = [n for n in range(1, total + 1) if plano.count(f"{tipo} {n} ") + plano.count(f"{tipo} {n}.") + plano.count(f"{tipo} {n},") + plano.count(f"{tipo} {n})") < 2]
    control(f"{tipo}s citadas en el texto", not sin_cita, f"sin cita aparte del rótulo: {sin_cita}")

control("Sin marcadores [[ ]] pendientes", "[[" not in texto and "]]" not in texto)
control("Sin guion largo", "—" not in texto)
control("Tabla de contenido presente", "Tabla de contenido" in texto and "Segunda parte" in paginas[1] + paginas[2])

# Referencias: cada entrada debe citarse en el texto (autor y sufijo de fecha).
ini = plano.rfind("Referencias")
cuerpo, refs = plano[:ini], plano[ini:]
citas = [
    ("Docker Inc. (s. f.-a)", "s. f.-a"), ("Jenkins. (s. f.-a)", "Jenkins, s. f.-a"), ("Jenkins. (s. f.-b)", "Jenkins, s. f.-b"),
    ("Jenkins. (s. f.-c)", "Jenkins, s. f.-c"), ("Jenkins. (s. f.-d)", "Jenkins, s. f.-d"), ("Jenkins. (s. f.-e)", "Jenkins, s. f.-e"),
    ("Jenkins. (s. f.-f)", "Jenkins, s. f.-f"), ("Jenkins. (s. f.-g)", "Jenkins, s. f.-g"),
    ("Politécnico Grancolombiano. (s. f.)", "Politécnico Grancolombiano, s. f."), ("Reyes Corredor, E. (2026", "Reyes Corredor, 2026"),
    ("Node.js. (s. f.)", "Node.js, s. f."), ("Express. (s. f.)", "Express, s. f."),
    ("PostgreSQL Global Development Group", "PostgreSQL Global Development Group, s. f."),
]
for entrada, cita in citas:
    control(f"Referencia citada: {entrada}", entrada in refs and cita in cuerpo.replace("Grancolombiano, s. f.", "Grancolombiano, s. f."))

for frase in ("Entrega 2, Proyecto 2, Escenario 5", "Jerlinson Fabián Ramos Galíndez", "Tutor: Edward Reyes Corredor"):
    control(f"Portada contiene '{frase}'", frase in paginas[0])

print("Total de fallas:", fallas)
