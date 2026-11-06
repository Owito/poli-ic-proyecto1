# Tareas del equipo (API REST en dos contenedores)

Proyecto del módulo **Énfasis Profesional I (Integración Continua)**, Politécnico Grancolombiano. Entrega 1: el proyecto vive en un repositorio de GitHub y se ejecuta con Docker en **dos contenedores comunicados entre sí**:

| Servicio | Contenedor | Imagen | Función |
|---|---|---|---|
| `app` | `tareas-app` | construida desde el `Dockerfile` (Node.js 22 + Express) | API REST de tareas del equipo |
| `db` | `tareas-db` | `postgres:16-alpine` | Base de datos PostgreSQL con volumen persistente |

Ambos están en la red `red-tareas`; la API llega a la base por el nombre de servicio `db`. Todos los datos de ejemplo son ficticios.

## Repositorio

https://github.com/Owito/poli-ic-proyecto1

```bash
git clone https://github.com/Owito/poli-ic-proyecto1.git
cd poli-ic-proyecto1
```

El repositorio guarda el `Dockerfile` y el `docker-compose.yml`, no las imágenes: al ejecutar `docker compose up -d --build`, Docker construye la imagen de la API y descarga `postgres:16-alpine` de Docker Hub.

## Requisitos

- Docker Desktop (o Docker Engine) con Docker Compose v2.
- Para ejecutar las pruebas fuera de Docker: Node.js 22 (opcional).

## Cómo levantarlo

```bash
cp .env.example .env        # en PowerShell: Copy-Item .env.example .env
# edite .env y cambie POSTGRES_PASSWORD
docker compose up -d --build
docker compose ps           # db debe quedar "healthy" y app "running"
curl http://localhost:8095/salud
curl http://localhost:8095/tareas
```

Para detener: `docker compose down` (conserva los datos). Para borrar también el volumen de la base: `docker compose down -v`.

## Endpoints

| Método | Ruta | Descripción |
|---|---|---|
| GET | `/salud` | Consulta la base de datos y devuelve su nombre y versión (503 si no hay conexión) |
| GET | `/tareas` | Lista las tareas; filtro opcional `?estado=pendiente\|en_progreso\|hecha` |
| GET | `/tareas/:id` | Detalle de una tarea (404 si no existe) |
| POST | `/tareas` | Crea una tarea. Cuerpo JSON: `titulo` (obligatorio), `responsable`, `estado` |
| PUT | `/tareas/:id` | Actualiza `titulo`, `responsable` o `estado` |
| DELETE | `/tareas/:id` | Borra la tarea (204) |

## Pruebas automatizadas

Usan el ejecutor nativo `node:test`, sin dependencias adicionales. Son las que ejecutarán Jenkins (Entrega 2) y Travis CI y Codeship (Entrega 3).

```bash
npm ci
npm test                                  # pruebas unitarias; la de integración se omite sin DB_HOST
docker compose exec app npm test          # unitarias + integración contra la base real
```

## Estructura

```
Dockerfile              imagen multietapa, usuario no root
docker-compose.yml      servicios app y db, red, volumen, healthcheck
.env.example            variables de entorno sin secretos reales
db/init.sql             esquema y datos de ejemplo
src/                    API (app.js, db.js, server.js)
tests/                  pruebas unitarias y de integración
```

## Variables de entorno

Ver `.env.example`. El archivo `.env` real no se versiona (está en `.gitignore`).
