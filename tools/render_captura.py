"""Convierte la salida REAL de un comando (guardada en texto) en una imagen PNG estilo terminal.

Uso: python tools/render_captura.py evidencias/texto/01-xxx.txt
La primera línea del .txt es el comando ("$ comando"); el resto es la salida tal cual.
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

FUENTE = r"C:\Windows\Fonts\consola.ttf"
TAM = 17
MAX_COLS = 130


def envolver(linea):
    if len(linea) <= MAX_COLS:
        return [linea]
    return [linea[i:i + MAX_COLS] for i in range(0, len(linea), MAX_COLS)]


def render(ruta_txt: Path, ruta_png: Path):
    texto = ruta_txt.read_text(encoding="utf-8", errors="replace").replace("\r", "").replace("\t", "    ")
    lineas = []
    for l in texto.rstrip("\n").split("\n"):
        lineas.extend(envolver(l))
    fuente = ImageFont.truetype(FUENTE, TAM)
    ancho_car = fuente.getlength("M")
    alto = TAM + 6
    barra = 34
    margen = 16
    ancho = int(max(len(l) for l in lineas) * ancho_car) + margen * 2
    ancho = max(ancho, 700)
    img = Image.new("RGB", (ancho, barra + margen * 2 + alto * len(lineas)), (24, 26, 31))
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, ancho, barra], fill=(52, 56, 66))
    for i, c in enumerate([(255, 95, 86), (255, 189, 46), (39, 201, 63)]):
        d.ellipse([14 + i * 22, 11, 26 + i * 22, 23], fill=c)
    d.text((100, 8), "Terminal - " + ruta_txt.stem, font=ImageFont.truetype(FUENTE, 15), fill=(200, 204, 214))
    y = barra + margen
    for l in lineas:
        color = (130, 220, 130) if l.startswith("$ ") else (225, 228, 235)
        d.text((margen, y), l, font=fuente, fill=color)
        y += alto
    img.save(ruta_png)


if __name__ == "__main__":
    for a in sys.argv[1:]:
        p = Path(a)
        render(p, p.parent.parent / (p.stem + ".png"))
        print("ok", p.stem)
