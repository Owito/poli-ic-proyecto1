# Auditoría de la Entrega 1 (Proyecto 1, Escenario 3) - Integración Continua

Fuentes: enunciado `IC_enunciado_proyecto.txt` y las grabaciones de los encuentros E1 (19-oct), E2 (26-oct) y E3 (5-nov-2026)
del tutor Edward Reyes Corredor. Los minutos son los de la transcripción de Teams de cada encuentro.
Actualizada el 6-nov-2026 con el E3. En Canvas (actividad 625707) no hay rúbrica ni descripción publicadas (consultado el 6-nov).

## 1. Indicaciones del tutor y dónde se cumplen

| Indicación | Fuente | Dónde se cumple | Estado |
|---|---|---|---|
| Proyecto creado en un repositorio de GitHub | Enunciado; E2 00:09:34, 00:09:45 | https://github.com/Owito/poli-ic-proyecto1 (remoto `origin`, rama `main`) | Cumplido |
| Paso a paso desde la creación del repositorio en GitHub | E3 00:06:07 | Tabla 2 del documento (comandos de creación y publicación) | Cumplido |
| Con Docker, DOS contenedores comunicados entre sí | Enunciado; E2 00:09:45, 00:39:51; E3 00:57:05 | `docker-compose.yml` (servicios `app` y `db`, red `red-tareas`); sección "Comunicación entre los contenedores" | Cumplido y verificado |
| Conexión entre Docker y GitHub | E3 00:43:35, 01:01:12, 01:04:36 | Sección "Conexión entre GitHub y Docker" y "Cómo reproducir el proyecto" (`git clone`) | Cumplido |
| Instalación de Docker sin el paso detallado | E3 00:06:07 | Figura de versiones y frase que lo aclara | Cumplido |
| Evidencia con pantallazos y explicación debajo | E1 00:15:06, 00:15:27; E2 00:12:14; E3 00:05:46 | Cada figura con número, título, nota y explicación | Cumplido |
| Documento PDF con normas APA: portada, introducción, conclusiones, referencias | E1 00:09:08; E2 00:13:30; E3 00:05:46, 00:06:07, 00:06:47 | `docs/Entrega1_IC_Proyecto1.pdf` | Cumplido al subir el PDF |
| Sin video en esta entrega (va en la final) | E3 00:06:07, 00:07:52 | No aplica | Cumplido |
| Acceso del tutor: no da su usuario; acepta repositorio público, usuario creado o enlace; recomienda público | E3 00:08:55, 00:09:05, 01:02:22 | Repositorio público desde el 6-nov (`gh repo view` -> PUBLIC; la página responde HTTP 200 sin sesión); figura de GitHub recapturada | Cumplido |
| Lo que está en GitHub basta para reconstruir el ambiente | E3 00:43:35, 01:01:12 | Figura de clonación: `git clone` del repo público, proyecto `tareas-clon`, 2 contenedores sanos, `/salud` 200 y 5/5 pruebas (6-nov) | Cumplido y verificado |
| Integrantes como colaboradores (Settings, Collaborators) | E3 00:07:52, 01:02:22 | Interruptor `COLABORADORES_INVITADOS` | Pendiente: usuarios de GitHub del grupo |
| En la portada solo quienes trabajaron; grupo de 5, no individual | E1 00:14:40; E3 00:02:58 | Portada con los 5 y nota resaltada | Pendiente: grupo |
| Herramientas libres; instalar solo lo necesario | E1 00:30:08; E2 00:31:14, 00:42:42 | Git, Docker, Node, Express, PostgreSQL (solo en contenedor) | Cumplido |
| Proyecto proporcional que integre las herramientas | E3 01:06:23 | Sección "Elección del proyecto y de las herramientas" | Cumplido |
| Entregar dentro de la ventana 7 a 10 de noviembre | E1 00:08:29; E2 00:06:00; E3 00:01:41 | Canvas | Pendiente: subir |

## 2. Lo que se verificó corriendo (5-nov-2026)

`docker compose up -d --build`, `docker compose ps` (db healthy, app running), `docker network inspect` (ambos en la red),
resolución de `db` desde la app, `curl` a `/salud`, `/tareas`, POST/PUT/validaciones, `psql` dentro de `db`, `npm test` en el
contenedor (5 pruebas, 5 aprobadas, 0 omitidas), recreación con `down` + `up` conservando datos. Salidas en `evidencias/texto/`.
URLs de referencias: `docs/verificacion_urls.txt`.

## 3. Pendientes del grupo

1. Hecho el 6-nov: repositorio público, figura recapturada y `REPO_PUBLICO = True`.
2. Usuarios de GitHub de los integrantes, invitarlos como colaboradores y que acepten; luego `COLABORADORES_INVITADOS = True`.
3. Portada: quiénes trabajaron y sus códigos.
4. Opcional: un párrafo del grupo con la experiencia y la distribución del trabajo (va al final de las conclusiones).
5. Regenerar: `python build_documento.py` y `powershell -ExecutionPolicy Bypass -File tools\exportar_pdf.ps1`.
6. Carlos sube el PDF a Canvas antes del 10-nov 11:55 p. m., con los nombres del grupo en la portada; el tutor carga la nota
   a cada integrante al calificar (decisión de Carlos, 6-nov).
