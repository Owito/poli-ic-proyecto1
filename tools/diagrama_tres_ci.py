"""Genera evidencias/diagrama-tres-ci.png: un repositorio, tres herramientas de integración continua y las mismas etapas."""
from pathlib import Path

from PIL import Image, ImageDraw

from diagrama import caja, fuente
from diagrama_jenkins import linea

SALIDA = Path(__file__).resolve().parent.parent / "evidencias" / "diagrama-tres-ci.png"
OSCURO = (20, 30, 50)


def etiqueta(d, centro, texto):
    """Texto con fondo blanco, para que se lea encima de una flecha."""
    f = fuente(15)
    x0, y0, x1, y1 = d.textbbox(centro, texto, font=f, anchor="mm")
    d.rectangle((x0 - 6, y0 - 4, x1 + 6, y1 + 4), fill=(255, 255, 255))
    d.text(centro, texto, font=f, fill=(60, 70, 90), anchor="mm")


def main():
    W, H = 1600, 1010
    img = Image.new("RGB", (W, H), (255, 255, 255))
    d = ImageDraw.Draw(img)
    d.text((W / 2, 18), "Un repositorio, tres integraciones continuas: arquitectura de la Entrega 3", font=fuente(26, True), fill=OSCURO, anchor="ma")

    caja(d, (30, 90, 330, 240), "Integrante del equipo", ["hace commit y git push", "a una rama del repositorio"], (255, 250, 235), (200, 150, 40))
    caja(d, (30, 330, 470, 700), "Repositorio en GitHub", ["Owito/poli-ic-proyecto1 (público)", "", "Jenkinsfile", ".travis.yml", ".github/workflows/", "   integracion-continua.yml", "Dockerfile y docker-compose.yml", "tests/ (unitarias e integración)"], (245, 245, 245), (120, 120, 120))
    linea(d, (180, 240), (180, 330), None)
    d.text((195, 272), "git push", font=fuente(16), fill=(60, 70, 90))

    caja(d, (640, 80, 1130, 280), "Jenkins (Entrega 2)", ["contenedor en el equipo del proyecto", "revisa la rama cada dos minutos", "lee el Jenkinsfile", "etapas en un solo pipeline"], (232, 248, 238), (50, 150, 90))
    caja(d, (640, 330, 1130, 530), "Travis CI (Entrega 3)", ["máquina virtual Ubuntu 24.04", "se activa con cada push", "lee .travis.yml", "dos etapas: unitarias y contenedores"], (252, 240, 232), (190, 110, 50))
    caja(d, (640, 580, 1130, 780), "GitHub Actions (Entrega 3)", ["runner ubuntu-24.04 de GitHub", "push, pull request o a mano", "lee integracion-continua.yml", "dos jobs: unitarias y contenedores"], (232, 240, 252), (60, 100, 180))

    # Codeship retirado
    d.rounded_rectangle((1220, 610, 1570, 750), radius=14, fill=(242, 242, 242), outline=(150, 150, 150), width=3)
    d.text((1395, 626), "Codeship (CloudBees)", font=fuente(21, True), fill=(110, 110, 110), anchor="ma")
    d.text((1395, 664), "retirado el 31 de enero de 2026", font=fuente(17), fill=(110, 110, 110), anchor="ma")
    d.text((1395, 692), "app.codeship.com redirige al", font=fuente(17), fill=(110, 110, 110), anchor="ma")
    d.text((1395, 718), "aviso de fin de vida", font=fuente(17), fill=(110, 110, 110), anchor="ma")
    d.rounded_rectangle((1330, 588, 1460, 616), radius=8, fill=(190, 80, 80))
    d.text((1395, 602), "RETIRADO", font=fuente(15, True), fill=(255, 255, 255), anchor="mm")
    for x in range(1135, 1215, 16):
        d.line([(x, 680), (x + 8, 680)], fill=(150, 150, 150), width=3)
    d.text((1175, 650), "lo reemplaza", font=fuente(14), fill=(110, 110, 110), anchor="ma")

    linea(d, (470, 420), (640, 180), None)
    linea(d, (470, 515), (640, 430), None)
    linea(d, (470, 610), (640, 680), None)
    etiqueta(d, (555, 300), "revisa la rama (pollSCM)")
    etiqueta(d, (555, 472), "push")
    etiqueta(d, (555, 645), "push o pull request")

    # Etapas comunes
    d.rounded_rectangle((30, 820, 1570, 990), radius=18, fill=(240, 246, 255), outline=(70, 110, 180), width=3)
    d.text((800, 832), "Las mismas etapas en las tres: si fallan las unitarias no se construye ni se levanta nada, y la 6 corre siempre", font=fuente(20, True), fill=(50, 90, 160), anchor="ma")
    etapas = ["1. npm ci y npm test", "2. docker compose build app", "3. docker compose up -d --wait", "4. curl /salud", "5. pruebas de integración", "6. docker compose down -v"]
    ancho = (1570 - 30 - 40) / len(etapas)
    for i, e in enumerate(etapas):
        x0 = 50 + i * ancho
        d.rounded_rectangle((x0, 880, x0 + ancho - 20, 960), radius=10, fill=(255, 255, 255), outline=(60, 100, 180), width=2)
        d.text((x0 + (ancho - 20) / 2, 908), e, font=fuente(16), fill=OSCURO, anchor="ma")
        if i < len(etapas) - 1:
            d.polygon([(x0 + ancho - 18, 912), (x0 + ancho - 4, 920), (x0 + ancho - 18, 928)], fill=(60, 100, 180))
    img.save(SALIDA)
    print("ok", SALIDA)


if __name__ == "__main__":
    main()
