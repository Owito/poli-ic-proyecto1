'use strict';
// Pruebas de integración: usan la base PostgreSQL real (el otro contenedor).
// Se omiten si no hay DB_HOST, así `npm test` también corre sin Docker.
const test = require('node:test');
const assert = require('node:assert/strict');
const { crearApp } = require('../src/app');
const { crearPool } = require('../src/db');

const hayBd = Boolean(process.env.DB_HOST);
const opciones = { skip: hayBd ? false : 'sin DB_HOST: prueba de integración omitida' };

test('integración: salud, crear, leer, actualizar y borrar una tarea', opciones, async () => {
  const pool = crearPool();
  const s = await new Promise((ok) => { const x = crearApp(pool).listen(0, () => ok(x)); });
  const url = `http://127.0.0.1:${s.address().port}`;
  const json = { 'content-type': 'application/json' };
  try {
    const salud = await (await fetch(`${url}/salud`)).json();
    assert.equal(salud.estado, 'ok');

    const lista = await (await fetch(`${url}/tareas`)).json();
    assert.ok(lista.length >= 5, 'deben existir los datos de ejemplo del script SQL');

    const crea = await fetch(`${url}/tareas`, { method: 'POST', headers: json, body: JSON.stringify({ titulo: 'Tarea de prueba', responsable: 'CI' }) });
    assert.equal(crea.status, 201);
    const nueva = await crea.json();
    assert.equal(nueva.estado, 'pendiente');

    const act = await fetch(`${url}/tareas/${nueva.id}`, { method: 'PUT', headers: json, body: JSON.stringify({ estado: 'hecha' }) });
    assert.equal((await act.json()).estado, 'hecha');

    assert.equal((await fetch(`${url}/tareas/${nueva.id}`, { method: 'DELETE' })).status, 204);
    assert.equal((await fetch(`${url}/tareas/${nueva.id}`)).status, 404);
  } finally {
    s.close();
    await pool.end();
  }
});
