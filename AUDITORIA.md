# Auditoría de la Entrega 1 (Proyecto 1, Escenario 3) - Integración Continua

Fuentes: enunciado `IC_enunciado_proyecto.txt` y `NOTAS_IC.md` (encuentros E1 19-oct y E2 26-oct de 2026, tutor Edward Reyes Corredor). Los minutos son los de `NOTAS_IC.md`.

## 1. Indicaciones del tutor y dónde se cumplen

| Indicación | Fuente | Dónde se cumple | Estado |
|---|---|---|---|
| Proyecto creado en un repositorio de GitHub | Enunciado Entrega 1; E2 00:09:34, 00:09:45 | Repo git local listo (`git log`); falta crear el remoto y hacer push | Parcial: paso humano |
| Con Docker, DOS contenedores comunicados entre sí | Enunciado; E2 00:09:45, 00:39:51 | `docker-compose.yml` (servicios `app` y `db`, red `red-tareas`); figuras 9, 10, 11, 12, 13 del documento | Cumplido y verificado |
| Evidencia con pantallazos y documentación de cada uno | E2 00:12:14, 00:13:00 | 18 figuras en `evidencias/`, cada una con "Qué muestra la Figura N" | Cumplido |
| "No pegar y pegar sin comentar"; indicar cada pantallazo | E1 00:15:06, 00:15:27 | Cada figura con número, título, nota y explicación | Cumplido |
| Documento con normas APA: portada, introducción, conclusiones, referencias | E1 00:09:08; E2 00:13:30 | `docs/Entrega1_IC_Proyecto1.docx` | Cumplido (PDF pendiente, ver 3) |
| Tutor como usuario de desarrollo del repositorio | E2 00:12:14, 00:13:00, 00:15:16 | Documento lo explica con marcador | Pendiente: paso humano |
| En la portada solo quienes trabajaron | E1 00:14:40 | Portada con los 5 y nota resaltada para confirmar | Pendiente: grupo |
| Herramientas libres | E1 00:30:08; E2 00:42:42 | Git, Docker, Node, Express, PostgreSQL | Cumplido |
| Instalar solo lo necesario | E2 00:31:14, 00:46:17 | Solo Git, Docker y Node; PostgreSQL solo en contenedor | Cumplido |
| Entregar dentro de la ventana 7 a 10 de noviembre | E1 00:08:29; E2 00:06:00 | Calendario | Pendiente: subir a Canvas |
| Grupal | E1 00:10:29 | Portada con integrantes | Pendiente: confirmar |

## 2. Lo que se verificó corriendo (5-nov-2026)

`docker compose up -d --build`, `docker compose ps` (db healthy, app running), `docker network inspect` (ambos en la red), resolución de `db` desde la app, `curl` a `/salud`, `/tareas`, POST/PUT/validaciones, `psql` dentro de `db`, `npm test` en el contenedor (5 pruebas, 5 aprobadas, 0 omitidas), recreación con `down` + `up` conservando datos. Salidas en `evidencias/texto/`. URLs de referencias: `docs/verificacion_urls.txt` (11 con HTTP 200).

## 3. Obstáculos

- PDF no generado: Word quedó colgado (proceso WINWORD 17128, iniciado por mi script) y no pude cerrarlo (acción denegada). Cierre Word y ejecute: `powershell -ExecutionPolicy Bypass -File tools\exportar_pdf.ps1` (antes borre `docs\~$trega1_IC_Proyecto1.docx` si queda). Eso también actualiza la tabla de contenido.
- Sin captura de GitHub ni de navegador: no hay remoto (regla dura).
- Quién sube a Canvas, formato (PDF o Word), repo público o privado: el tutor no lo dijo (NOTAS_IC, sección 8).
- Códigos estudiantiles, nombres definitivos y conclusión del grupo: resaltados en amarillo en el documento.

## 4. Pasos humanos

1. Confirmar con el grupo quiénes van en la portada y completar los códigos.
2. Crear el repo en GitHub (privado o público, decide el grupo).
3. `git remote add origin https://github.com/USUARIO/REPO.git` y `git push -u origin main`.
4. Invitar al tutor como colaborador (Settings, Collaborators) y comprobar que entra; pedirle su usuario por el chat.
5. Tomar captura del repo en GitHub y pegarla en el documento (marcador resaltado); poner la URL.
6. Cerrar Word, generar el PDF, revisar la tabla de contenido y los resaltados.
7. Decidir quién sube a Canvas (poli.instructure.com, tarea Proyecto 1, escenario 3) antes del 10-nov 11:55 p. m.
8. Opcional: preguntar al tutor formato (PDF/Word/ZIP) y si exige repo privado.
