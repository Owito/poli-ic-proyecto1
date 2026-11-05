'use strict';
const express = require('express');

const ESTADOS = ['pendiente', 'en_progreso', 'hecha'];

// Fábrica de la aplicación: recibe el "pool" para poder probarla con una BD real o simulada.
function crearApp(pool) {
  const app = express();
  app.use(express.json());

  // Comprueba que la API alcanza la base de datos (el otro contenedor).
  app.get('/salud', async (_req, res) => {
    try {
      const r = await pool.query('SELECT current_database() AS base, version() AS version, now() AS hora');
      res.json({ estado: 'ok', base_de_datos: r.rows[0].base, version: r.rows[0].version, hora_bd: r.rows[0].hora });
    } catch (e) {
      res.status(503).json({ estado: 'error', detalle: 'No hay conexión con la base de datos' });
    }
  });

  app.get('/tareas', async (req, res, next) => {
    try {
      const { estado } = req.query;
      if (estado && !ESTADOS.includes(estado)) return res.status(400).json({ error: 'estado inválido' });
      const r = estado
        ? await pool.query('SELECT * FROM tareas WHERE estado = $1 ORDER BY id', [estado])
        : await pool.query('SELECT * FROM tareas ORDER BY id');
      res.json(r.rows);
    } catch (e) { next(e); }
  });

  app.get('/tareas/:id(\\d+)', async (req, res, next) => {
    try {
      const r = await pool.query('SELECT * FROM tareas WHERE id = $1', [req.params.id]);
      if (!r.rows.length) return res.status(404).json({ error: 'tarea no encontrada' });
      res.json(r.rows[0]);
    } catch (e) { next(e); }
  });

  app.post('/tareas', async (req, res, next) => {
    try {
      const { titulo, responsable, estado = 'pendiente' } = req.body || {};
      if (typeof titulo !== 'string' || !titulo.trim()) return res.status(400).json({ error: 'titulo es obligatorio' });
      if (!ESTADOS.includes(estado)) return res.status(400).json({ error: 'estado inválido' });
      const r = await pool.query(
        'INSERT INTO tareas (titulo, responsable, estado) VALUES ($1, $2, $3) RETURNING *',
        [titulo.trim(), responsable || null, estado]);
      res.status(201).json(r.rows[0]);
    } catch (e) { next(e); }
  });

  app.put('/tareas/:id(\\d+)', async (req, res, next) => {
    try {
      const { titulo, responsable, estado } = req.body || {};
      if (estado !== undefined && !ESTADOS.includes(estado)) return res.status(400).json({ error: 'estado inválido' });
      if (titulo !== undefined && (typeof titulo !== 'string' || !titulo.trim())) return res.status(400).json({ error: 'titulo inválido' });
      const r = await pool.query(
        `UPDATE tareas SET titulo = COALESCE($1, titulo), responsable = COALESCE($2, responsable),
         estado = COALESCE($3, estado) WHERE id = $4 RETURNING *`,
        [titulo ? titulo.trim() : null, responsable ?? null, estado ?? null, req.params.id]);
      if (!r.rows.length) return res.status(404).json({ error: 'tarea no encontrada' });
      res.json(r.rows[0]);
    } catch (e) { next(e); }
  });

  app.delete('/tareas/:id(\\d+)', async (req, res, next) => {
    try {
      const r = await pool.query('DELETE FROM tareas WHERE id = $1', [req.params.id]);
      if (!r.rowCount) return res.status(404).json({ error: 'tarea no encontrada' });
      res.status(204).end();
    } catch (e) { next(e); }
  });

  app.use((_req, res) => res.status(404).json({ error: 'ruta no encontrada' }));
  // eslint-disable-next-line no-unused-vars
  app.use((err, _req, res, _next) => {
    console.error(err.message);
    res.status(500).json({ error: 'error interno' });
  });
  return app;
}

module.exports = { crearApp, ESTADOS };
