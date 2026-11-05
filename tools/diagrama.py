"""Genera evidencias/diagrama-arquitectura.png (diagrama simple de los dos contenedores)."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

SALIDA = Path(__file__).resolve().parent.parent / "evidencias" / "diagrama-arquitectura.png"
F = r"C:\Windows\Fonts\arial.ttf"
FB = r"C:\Windows\Fonts\arialbd.ttf"


def fuente(tam, negrita=False):
    return ImageFont.truetype(FB if negrita else F, tam)


def caja(d, xy, titulo, lineas, relleno, borde):
    d.rounded_rectangle(xy, radius=14, fill=relleno, outline=borde, width=3)
    x0, y0, x1, _ = xy
    d.text(((x0 + x1) / 2, y0 + 16), titulo, font=fuente(22, True), fill=(20, 30, 50), anchor="ma")
    for i, l in enumerate(lineas):
        d.text(((x0 + x1) / 2, y0 + 56 + i * 26), l, font=fuente(18), fill=(40, 50, 70), anchor="ma")


def flecha(d, p0, p1, texto=None, doble=False):
    d.line([p0, p1], fill=(60, 70, 90), width=4)
    def punta(a, b):
        dx = 1 if b[0] > a[0] else -1
        d.polygon([b, (b[0] - 14 * dx, b[1] - 8), (b[0] - 14 * dx, b[1] + 8)], fill=(60, 70, 90))
    punta(p0, p1)
    if doble:
        punta(p1, p0)
    if texto:
        d.text(((p0[0] + p1[0]) / 2, p0[1] - 34), texto, font=fuente(16), fill=(60, 70, 90), anchor="ma")


def main():
    W, H = 1500, 640
    img = Image.new("RGB", (W, H), (255, 255, 255))
    d = ImageDraw.Draw(img)
    # red Docker
    d.rounded_rectangle((330, 90, 1170, 560), radius=22, fill=(240, 246, 255), outline=(70, 110, 180), width=3)
    d.text((750, 100), "Red Docker: red-tareas (bridge)", font=fuente(20, True), fill=(50, 90, 160), anchor="ma")
    caja(d, (30, 200, 280, 400), "Cliente", ["navegador o curl", "en el equipo anfitrión"], (255, 250, 235), (200, 150, 40))
    caja(d, (380, 190, 700, 440), "Contenedor tareas-app", ["servicio: app", "Node.js 22 + Express", "puerto interno 3000", "usuario no root (node)"], (232, 248, 238), (50, 150, 90))
    caja(d, (800, 190, 1120, 440), "Contenedor tareas-db", ["servicio: db", "PostgreSQL 16 (alpine)", "puerto interno 5432", "healthcheck: pg_isready"], (232, 240, 252), (60, 100, 180))
    flecha(d, (280, 300), (380, 300), "HTTP", doble=True)
    d.text((330, 330), "puerto 8095", font=fuente(15), fill=(90, 90, 90), anchor="ma")
    flecha(d, (700, 300), (800, 300), "SQL", doble=True)
    d.text((750, 330), "DB_HOST=db", font=fuente(15), fill=(90, 90, 90), anchor="ma")
    # volumen e init
    d.rounded_rectangle((830, 590, 1090, 630), radius=10, fill=(255, 240, 240), outline=(180, 70, 70), width=3)
    d.text((960, 598), "Volumen tareas-datos", font=fuente(18, True), fill=(120, 40, 40), anchor="ma")
    d.line([(960, 440), (960, 590)], fill=(180, 70, 70), width=4)
    d.text((1000, 500), "/var/lib/postgresql/data", font=fuente(14), fill=(120, 40, 40))
    d.rounded_rectangle((1220, 230, 1480, 300), radius=10, fill=(245, 245, 245), outline=(120, 120, 120), width=3)
    d.text((1350, 240), "db/init.sql", font=fuente(18, True), fill=(60, 60, 60), anchor="ma")
    d.text((1350, 268), "esquema y datos de ejemplo", font=fuente(14), fill=(60, 60, 60), anchor="ma")
    d.line([(1220, 265), (1120, 265)], fill=(120, 120, 120), width=3)
    d.polygon([(1120, 265), (1134, 257), (1134, 273)], fill=(120, 120, 120))
    d.text((750, 505), "depends_on: db (condition: service_healthy)", font=fuente(16), fill=(50, 90, 160), anchor="ma")
    d.text((750, 530), "La app arranca solo cuando la base responde como saludable", font=fuente(14), fill=(50, 90, 160), anchor="ma")
    d.text((750, 20), "Arquitectura de los dos contenedores", font=fuente(26, True), fill=(20, 30, 50), anchor="ma")
    img.save(SALIDA)
    print("ok", SALIDA)


if __name__ == "__main__":
    main()
