# Auditoría de la Entrega 2 (Proyecto 2, Escenario 5): Jenkins

Canvas 101015 · actividad 625708 · 150 puntos · grupal · abre sáb 21-nov-2026, vence mar 24-nov-2026 11:55 p. m.
Rama `poli-trabajos/ic-proyecto2-jenkins` · documento `docs/Entrega2_IC_Proyecto.pdf` (generado por `build_documento.py`).
Elaborada el 17-nov-2026. Marcas: ✅ verificado en esta sesión (se dice cómo) · 🟡 inferido · ❓ no se sabe.

Fuentes leídas en esta sesión: enunciado `C:\Users\fgogu\poli-bloque2-material\IC_enunciado_proyecto.txt` (completo);
`C:\Users\fgogu\poli-grabaciones-2s2026\ic\RESUMEN_ESCENARIO4_IC.md` (completo) y la transcripción
`ic-esc4-2026-11-09.txt` (búsquedas puntuales de cada frase atribuida al tutor, líneas citadas abajo como `T:n`);
documentación de jenkins.io (7 páginas, citadas en el documento); Canvas por API (actividad 625708, anuncio 954155,
entrega 625707 y su DocViewer).

## 1. Contrato de la entrega

- ✅ Canvas, 17-nov: la actividad 625708 no tiene descripción ni rúbrica visibles mientras está bloqueada (`description`
  vacío, `rubric` ausente; consulta `/api/v1/courses/101015/assignments?per_page=100`). `group_category_id` 152453.
- ✅ Enunciado, "ENTREGA 2 SEMANA 5": "Es momento de aplicar la integración continua. Esta semana se implementa Jenkins como
  gestor de operaciones. Para esto es necesario entregar un documento con las características requeridas para la
  implementación de esta herramienta."
- ✅ Anuncio 954155 (10-nov) "MATERIAL DE APOYO - ESCENARIO CUATRO - INSTALACION DE JENKINS": adjunto
  `Instalación de Jenkins.pdf` (file 101295510, 654.977 B). Se leyó en memoria del navegador (no se guardó en disco):
  son 40 imágenes sin capa de texto, así que su contenido se toma de lo que el tutor mostró en el encuentro (T:60).
- ✅ Retro de la Entrega 1 (625707): 144/150, un solo comentario genérico ("Favor mirar la rúbrica de evaluación y los
  criterios evaluados ... así como los comentarios relacionados al entregable"), `rubric_assessment` vacío y **sin
  anotaciones** en el PDF (`/annotations` del DocViewer devolvió `{"data":[]}`). No hay nada concreto que corregir de la parte 1.

## 2. Indicaciones y dónde se cumplen

| Indicación | Fuente | Dónde se cumple en el documento | Estado |
|---|---|---|---|
| Implementar Jenkins como gestor de operaciones | Enunciado, Entrega 2 | Segunda parte completa; Jenkins instalado y ejecutando el pipeline (ejecuciones 1 a 4) | ✅ |
| Documento con las características requeridas para la implementación | Enunciado, Entrega 2 | "Características de Jenkins" (Tabla 4) y "Requisitos para implementar Jenkins en el proyecto" (Tabla 5) | ✅ |
| Saber qué hace, qué características tiene y cómo implementarla en el proyecto | Tutor, [00:05:42] T:22 | "Qué es Jenkins y qué papel cumple", Tabla 4, instalación, Jenkinsfile y trabajo | ✅ |
| Explorar, instalar, "cacharrear" | Tutor, [00:04:34] T:19 | Instalación real con Docker, asistente completo, dos trabajos, error a propósito | ✅ |
| Adaptarla al proyecto que se viene trabajando | Tutor, T:60 | El pipeline prueba, construye y levanta la misma API de la parte 1 | ✅ |
| Ver qué plugins sirven | Tutor, T:60 | "Plugins utilizados" (Tabla 6, versiones leídas de la API de Jenkins) | ✅ |
| Integración con Docker | Tutor, T:53 | Instalación con docker:dind, etapa en contenedor `node:22-alpine`, `docker compose` en el pipeline (Figuras 21 a 23, 32 y 33) | ✅ |
| "Ubiquen el plugin J Cloud" | Tutor, T:53 (dictado) | No se identificó el plugin; no se menciona en el documento | ❓ |
| Explicar el paso a paso, no pegar fotos | Tutor, [00:59:34] T:80 | Cada figura lleva "Qué muestra la Figura N"; listados explicados; tablas de etapas y requisitos | ✅ |
| Mismo documento, "segunda parte" | Tutor, [00:56:35] T:62 | La parte 1 se conserva como "Primera parte"; la nueva es "Segunda parte" | ✅ |
| Portada, introducción del proyecto, secuencia del desarrollo | Tutor, [00:58:00] T:66 | Portada actualizada, introducción con ambas partes y objetivos 5 a 7 | ✅ |
| Historial de ejecuciones | Tutor, T:31 y T:53 | Figura 37 (historial por etapa) y Tabla 8 | ✅ |
| Sin video en esta entrega | Tutor, [01:00:23] T:86 | No aplica | ✅ |
| PDF con normas APA | Tutor, entrega 1 (NOTAS_IC.md §10); 🟡 se mantiene | Mismo generador y formato APA 7 de la entrega 1 | 🟡 |

## 3. Qué se verificó y cómo

- ✅ Jenkins 2.580.2 en `http://localhost:8181`, contenedores `jenkins` y `jenkins-docker` (Figura 22, `docker ps`).
- ✅ Asistente completo: desbloqueo, plugins sugeridos (43 en 52 s según `docker logs jenkins`), usuario `admin-ci`, URL.
- ✅ Trabajo `tareas-equipo` (Pipeline script from SCM): `config.xml` leído por API con URL, rama y `scriptPath`.
- ✅ Ejecuciones (API `/job/tareas-equipo/N/api/json` y `consoleText`): #1 SUCCESS (manual, 1 min 5 s), #2 ABORTED
  (SCM, 7 min 6 s, colgada), #3 FAILURE (SCM, 7,8 s), #4 SUCCESS (SCM, 19 s). La #4: 4 unitarias aprobadas y 1 omitida,
  `/salud` con PostgreSQL 16.16, 5 de 5 pruebas de integración.
- ✅ Trabajo `tareas-equipo-github` #1 SUCCESS (8 s): clona `https://github.com/Owito/poli-ic-proyecto1.git`, commit 190bf46.
- ✅ Pruebas locales: `npm test` con el error presente termina (3 aprobadas, 1 fallida, 1 omitida) después del arreglo `t.after`.
- ✅ Frases atribuidas al tutor, contrastadas con la transcripción: gratuito, CI/CD y "code commit build y deploy" (T:53),
  más de 1000 plugins (T:53), Docker como integración valiosa (T:53), nube (T:53), historial (T:31, T:53), JDK 21 o 25,
  puerto 8080, firewall y servicio con sistema local (T:60), "segunda parte" y "todo el proyecto en uno solo" (T:62),
  "no es pegar" (T:80), "explorarlas, instalarlas" (T:19).
- ✅ PDF revisado con `tools/verificar_pdf.py` (resultado en la sección 5).

## 5. Revisión del PDF (✅ 17-nov, `python tools/verificar_pdf.py docs/Entrega2_IC_Proyecto.pdf`)

- 69 páginas. Figuras 1 a 41, tablas 1 a 8 y listados 1 a 5 rotulados en orden y todos citados en el texto (la Tabla 3 de la
  parte 1 no se citaba en la entrega 1; se agregó la frase). Sin marcadores `[[ ]]`, sin guion largo, portada con
  "Entrega 2, Proyecto 2, Escenario 5", "Jerlinson Fabián Ramos Galíndez" (como en el PDF entregado en Canvas) y el tutor.
- Las 13 referencias revisadas por el script están citadas en el texto; las 7 de jenkins.io se leyeron en esta sesión.
- Tabla de contenido: 41 entradas, cada número de página coincide con la página real del título (script con pypdf).
- Revisión visual de las páginas 1 a 4, 35 a 40, 53 a 58 y 66 a 69: figuras legibles, sin desbordes; quedan espacios en
  blanco al final de algunas páginas porque cada figura va pegada a su título (mismo criterio de la entrega 1).
- Mapa de la segunda parte: inicio p. 33 · características 34 · requisitos 35 · instalación 36 · configuración 41 ·
  plugins 46 · credencial 47 · Jenkinsfile 48 · trabajo y repositorio 52 · ejecuciones 54 · prueba que falla 57 ·
  GitHub 60 · dificultades 64 · conclusiones 66 · referencias 68.

## 4. Pendientes humanos ⏰

1. **Publicar la rama** (decisión de Carlos; el repositorio es público): `git -C C:\Users\fgogu\poli-ic-proyecto1 push -u origin poli-trabajos/ic-proyecto2-jenkins`
   y, si el grupo lo aprueba, integrarla a `main` con un PR. Mientras no se publique, el tutor no verá el `Jenkinsfile` en GitHub.
2. Después de publicar, cambiar en Jenkins la URL del trabajo `tareas-equipo` a GitHub (opcional; el documento ya lo explica).
3. Subir `docs/Entrega2_IC_Proyecto.pdf` a la actividad 625708 entre el 21 y el 24-nov, y decidir quién sube (grupal, 152453).
4. Revisar el texto de la actividad y la rúbrica cuando abra el 21-nov; si piden algo distinto, ajustar `build_documento.py`.
5. Preguntar en el encuentro del jueves 19-nov qué es el plugin "J Cloud".

## 🧱 Obstáculos

- ❓ Rúbrica de los 150 puntos: no existe en Canvas hasta que abra la actividad. Paso: `GET /api/v1/courses/101015/assignments/625708` el 21-nov.
- ❓ Plugin "J Cloud" (dictado automático): no identificado. Paso: preguntar al tutor el 19-nov.
- 🟡 El manual de instalación del tutor (40 imágenes) no se transcribió; lo que el documento le atribuye sale del encuentro (T:60).
- 🟡 El documento afirma que el grupo eligió Docker en lugar del instalador de Windows; la elección la hizo esta sesión de trabajo
  por continuidad con la parte 1. Si el grupo prefiere el instalador del tutor, la sección "Por qué con Docker" debe reescribirse.
- La primera ejecución de la captura del asistente se colgó buscando el formulario del primer usuario; se reanudó con el mismo
  script y no afectó a Jenkins (las capturas j02 a j04 son de la primera pasada, j05 y j08 de la reanudación).
- Al arrancar Docker Desktop se levantaron solos los contenedores de otros proyectos con política de reinicio (Credisabe,
  inventario, h2h). No se tocaron.
