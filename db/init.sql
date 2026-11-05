-- Script de inicialización: lo ejecuta PostgreSQL la primera vez que se crea el volumen.
CREATE TABLE IF NOT EXISTS tareas (
  id          SERIAL PRIMARY KEY,
  titulo      VARCHAR(120) NOT NULL,
  responsable VARCHAR(80),
  estado      VARCHAR(20) NOT NULL DEFAULT 'pendiente'
              CHECK (estado IN ('pendiente', 'en_progreso', 'hecha')),
  creada_en   TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Datos de ejemplo (ficticios).
INSERT INTO tareas (titulo, responsable, estado) VALUES
  ('Crear el repositorio del proyecto en GitHub', 'Integrante 1', 'hecha'),
  ('Escribir el Dockerfile de la API', 'Integrante 2', 'hecha'),
  ('Definir docker-compose con la API y la base de datos', 'Integrante 3', 'en_progreso'),
  ('Redactar las pruebas automatizadas', 'Integrante 4', 'pendiente'),
  ('Preparar el documento de la Entrega 1', 'Integrante 5', 'pendiente');
