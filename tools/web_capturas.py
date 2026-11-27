"""Capturas REALES de páginas públicas de documentación oficial, con un navegador sin ventana (Playwright + Microsoft Edge).

Uso (con el entorno virtual que tiene playwright instalado):
    python tools/web_capturas.py
Cada captura recorta la zona de la página que contiene el texto citado en el documento; no se modifica la página.
"""
from pathlib import Path

from playwright.sync_api import sync_playwright

EVID = Path(__file__).resolve().parent.parent / "evidencias"
VISTA = {"width": 1280, "height": 900}

# (archivo, dirección, texto que debe aparecer, píxeles por encima del texto, alto, x inicial y ancho del recorte)
CAPTURAS = [
    ("t01-travis-prueba-gratuita.png", "https://docs.travis-ci.com/user/billing-overview/",
     "Free Trial Plans are currently unavailable", 70, 450, 365, 880),
    ("t02-codeship-fin-de-vida.png", "https://docs.cloudbees.com/docs/cloudbees-codeship/latest/",
     "As of January 2026, CloudBees no longer supports", 80, 400, 335, 610),
    ("t04-actions-gratis-publicos.png", "https://docs.github.com/en/actions/concepts/billing-and-usage",
     "GitHub Actions usage is free for standard GitHub-hosted runners in public repositories", 100, 300, 425, 760),
]


def capturar(page, archivo, url, texto, arriba, alto, x, ancho):
    page.goto(url, wait_until="networkidle", timeout=60000)
    page.wait_for_timeout(1500)
    el = page.get_by_text(texto, exact=False).first
    el.wait_for(state="visible", timeout=20000)
    el.scroll_into_view_if_needed()
    y = page.evaluate("(e) => e.getBoundingClientRect().top + window.scrollY", el.element_handle())
    clip = {"x": x, "y": max(0, y - arriba), "width": ancho, "height": alto}
    page.screenshot(path=str(EVID / archivo), full_page=True, clip=clip)
    print("captura", archivo, "desde", page.url)


with sync_playwright() as p:
    nav = p.chromium.launch(channel="msedge", headless=True)
    ctx = nav.new_context(viewport=VISTA, locale="en-US")
    page = ctx.new_page()
    for c in CAPTURAS:
        capturar(page, *c)
    nav.close()
