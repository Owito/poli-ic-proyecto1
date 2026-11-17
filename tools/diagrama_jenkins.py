"""Genera evidencias/diagrama-jenkins.png: cómo se instaló Jenkins con Docker y dónde corre cada etapa."""
from pathlib import Path

from PIL import Image, ImageDraw

from diagrama import caja, fuente

SALIDA = Path(__file__).resolve().parent.parent / "evidencias" / "diagrama-jenkins.png"
GRIS = (60, 70, 90)


def linea(d, p0, p1, texto=None, color=GRIS, dx_texto=0):
    d.line([p0, p1], fill=color, width=4)
    (x0, y0), (x1, y1) = p0, p1
    if abs(x1 - x0) >= abs(y1 - y0):
        s = 1 if x1 > x0 else -1
        d.polygon([p1, (x1 - 14 * s, y1 - 8), (x1 - 14 * s, y1 + 8)], fill=color)
    else:
        s = 1 if y1 > y0 else -1
        d.polygon([p1, (x1 - 8, y1 - 14 * s), (x1 + 8, y1 - 14 * s)], fill=color)
    if texto:
        d.text(((x0 + x1) / 2 + dx_texto, min(y0, y1) - 26 if y0 == y1 else (y0 + y1) / 2), texto,
               font=fuente(15), fill=color, anchor="ma")


def main():
    W, H = 1600, 820
    img = Image.new("RGB", (W, H), (255, 255, 255))
    d = ImageDraw.Draw(img)
    d.text((W / 2, 18), "Jenkins con Docker: arquitectura de la Entrega 2", font=fuente(26, True), fill=(20, 30, 50), anchor="ma")

    # equipo anfitrión
    d.rounded_rectangle((20, 70, 1580, 800), radius=22, fill=(250, 250, 250), outline=(150, 150, 150), width=3)
    d.text((40, 80), "Equipo anfitrión: Windows 11 con Docker Desktop (contenedores Linux)", font=fuente(18, True), fill=(80, 80, 80))

    caja(d, (50, 150, 330, 330), "Navegador", ["interfaz web de Jenkins", "http://localhost:8181"], (255, 250, 235), (200, 150, 40))
    caja(d, (50, 470, 330, 700), "Repositorio Git", ["copia local (bare) de la rama", "montada en /repos (solo lectura)", "", "después: GitHub", "Owito/poli-ic-proyecto1"], (245, 245, 245), (120, 120, 120))

    # red jenkins
    d.rounded_rectangle((400, 120, 1560, 780), radius=22, fill=(240, 246, 255), outline=(70, 110, 180), width=3)
    d.text((980, 130), "Red Docker: jenkins (bridge)", font=fuente(20, True), fill=(50, 90, 160), anchor="ma")

    caja(d, (440, 190, 780, 470), "Contenedor jenkins", ["imagen tareas-jenkins:2.580.2-1", "Jenkins 2.580.2 sobre Java 21", "cliente de Docker y Compose", "plugins: Pipeline, Git,", "Docker Pipeline, Credentials", "puerto 8080 -> 8181 del equipo"], (232, 248, 238), (50, 150, 90))
    d.rounded_rectangle((870, 180, 1530, 740), radius=16, fill=(232, 240, 252), outline=(60, 100, 180), width=3)
    d.text((1200, 192), "Contenedor jenkins-docker (docker:dind)", font=fuente(21, True), fill=(20, 30, 50), anchor="ma")
    d.text((1200, 222), "motor de Docker propio, alias docker, puerto 2376 con TLS", font=fuente(16), fill=(40, 50, 70), anchor="ma")
    caja(d, (900, 270, 1500, 400), "Etapa de pruebas unitarias", ["contenedor temporal node:22-alpine", "npm ci y npm test sobre el espacio de trabajo"], (255, 255, 255), (60, 100, 180))
    caja(d, (900, 430, 1500, 600), "Proyecto levantado por el pipeline", ["tareas-app (imagen construida) + tareas-db", "red red-tareas y volumen temporal", "puerto 8095 publicado en el motor docker"], (255, 255, 255), (60, 100, 180))
    d.text((1200, 625), "Al terminar: docker compose down -v", font=fuente(15), fill=(50, 90, 160), anchor="ma")
    d.text((1200, 650), "(cada ejecución empieza con la base limpia)", font=fuente(15), fill=(50, 90, 160), anchor="ma")

    d.rounded_rectangle((470, 650, 780, 710), radius=10, fill=(255, 240, 240), outline=(180, 70, 70), width=3)
    d.text((625, 660), "Volumen jenkins-data", font=fuente(17, True), fill=(120, 40, 40), anchor="ma")
    d.text((625, 684), "/var/jenkins_home (compartido)", font=fuente(14), fill=(120, 40, 40), anchor="ma")
    d.line([(610, 470), (610, 650)], fill=(180, 70, 70), width=4)
    d.line([(780, 690), (900, 690)], fill=(180, 70, 70), width=4)
    d.text((825, 662), "mismo", font=fuente(13), fill=(120, 40, 40), anchor="ma")
    d.text((825, 702), "volumen", font=fuente(13), fill=(120, 40, 40), anchor="ma")

    linea(d, (330, 240), (440, 240), "HTTP", dx_texto=-22)
    linea(d, (330, 585), (440, 420), None)
    d.text((300, 380), "checkout", font=fuente(15), fill=GRIS)
    linea(d, (780, 300), (870, 300), "comandos", dx_texto=0)
    d.text((825, 320), "TCP 2376", font=fuente(13), fill=(90, 90, 90), anchor="ma")
    linea(d, (780, 420), (900, 515), None)
    d.text((790, 505), "curl /salud", font=fuente(14), fill=GRIS)
    img.save(SALIDA)
    print("ok", SALIDA)


if __name__ == "__main__":
    main()
